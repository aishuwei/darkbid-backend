import os
import io
import numpy as np
from PIL import Image
from paddleocr import PaddleOCR
from docx.oxml.ns import qn
from lxml import etree


# ========== 日志函数 ==========
def log(msg):
    print(f"[OCR引擎] {msg}")


# ========== OCR 引擎类 ==========
class OCREngine:
    def __init__(self):
        self.ocr = None
        self._init_ocr()

    def _init_ocr(self):
        try:
            log("正在初始化 PaddleOCR (CPU模式)...")
            self.ocr = PaddleOCR(
                use_angle_cls=False,
                use_textline_orientation=False,
                lang='ch'
            )
            log("PaddleOCR 初始化成功")
        except Exception as e:
            log(f"❌ PaddleOCR 初始化失败: {e}")

    def extract_images_with_positions(self, file_path):
        """
        从 Word 文档中提取图片及其位置信息
        返回: list[dict] 包含 PIL Image、段落索引、锚点信息等
        """
        import zipfile

        images_info = []
        try:
            docx_zip = zipfile.ZipFile(file_path)
        except Exception as e:
            log(f"无法打开 docx 文件: {e}")
            return []

        # 解析 document.xml
        try:
            with docx_zip.open('word/document.xml') as doc_xml:
                tree = etree.parse(doc_xml)
                root = tree.getroot()
        except Exception as e:
            log(f"解析 document.xml 失败: {e}")
            docx_zip.close()
            return []

        # 命名空间映射
        nsmap = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
            'v': 'urn:schemas-microsoft-com:vml',
            'o': 'urn:schemas-microsoft-com:office:office',
        }

        # 建立段落元素 -> 索引映射
        para_elements = root.xpath('.//w:p', namespaces=nsmap)
        para_idx_map = {elem: i for i, elem in enumerate(para_elements)}

        # ---------- 关键修复：加载 relationships（使用 local-name 无视命名空间）----------
        rels_map = {}
        try:
            with docx_zip.open('word/_rels/document.xml.rels') as rels_xml:
                rels_tree = etree.parse(rels_xml)
                # 用 local-name() 避免命名空间前缀不一致的问题
                rel_nodes = rels_tree.xpath('//*[local-name()="Relationship"]')
                log(f"Relationships 文件中共找到 {len(rel_nodes)} 个关系")
                for rel in rel_nodes:
                    rid = rel.get('Id')
                    target = rel.get('Target')
                    if rid:
                        rels_map[rid] = target
                        log(f"  关系映射: {rid} -> {target}")
        except KeyError:
            log("未找到 word/_rels/document.xml.rels")
        except Exception as e:
            log(f"解析 relationships 出错: {e}")
        # ---------------------------------------------------------------------------

        image_entries = []

        # 1. DrawingML 格式 (w:drawing -> a:blip)
        drawings = root.xpath('.//w:drawing', namespaces=nsmap)
        log(f"找到 {len(drawings)} 个 DrawingML 对象")
        for drawing in drawings:
            blip = drawing.xpath('.//a:blip', namespaces=nsmap)
            if blip:
                embed = blip[0].get(qn('r:embed'))
                link = blip[0].get(qn('r:link'))
                rid = embed or link
                if rid:
                    parent_p = drawing.xpath('ancestor::w:p', namespaces=nsmap)
                    para_idx = para_idx_map.get(parent_p[0], -1) if parent_p else -1
                    image_entries.append({
                        'rid': rid,
                        'para_idx': para_idx,
                        'elem': drawing,
                        'type': 'drawing'
                    })
                    log(f"  DrawingML 图片引用: {rid}, 所在段落: {para_idx}")

        # 2. VML 旧格式 (w:pict -> v:imagedata)
        picts = root.xpath('.//w:pict', namespaces=nsmap)
        log(f"找到 {len(picts)} 个 VML 对象")
        for pict in picts:
            imagedata = pict.xpath('.//v:imagedata', namespaces=nsmap)
            if imagedata:
                rid = imagedata[0].get(qn('r:id'))  # VML 用 r:id
                if rid:
                    parent_p = pict.xpath('ancestor::w:p', namespaces=nsmap)
                    para_idx = para_idx_map.get(parent_p[0], -1) if parent_p else -1
                    image_entries.append({
                        'rid': rid,
                        'para_idx': para_idx,
                        'elem': pict,
                        'type': 'pict'
                    })
                    log(f"  VML 图片引用: {rid}, 所在段落: {para_idx}")

        log(f"从 XML 中解析出 {len(image_entries)} 个图片引用")

        # 遍历提取图片二进制并转为 PIL Image
        for idx, entry in enumerate(image_entries):
            target = rels_map.get(entry['rid'])
            if not target:
                log(f"⚠️  图片引用 {entry['rid']} 在 relationships 中未找到，已跳过")
                continue

            # 处理 Target 路径（可能是相对路径如 media/image1.png，也可能是绝对路径 /word/media/image1.png）
            target = target.lstrip('/')
            if not target.startswith('word/'):
                image_path = 'word/' + target
            else:
                image_path = target

            try:
                image_data = docx_zip.read(image_path)
            except KeyError:
                log(f"压缩包中找不到图片文件: {image_path}")
                continue

            # 转为 PIL Image
            try:
                image = Image.open(io.BytesIO(image_data)).convert('RGB')
            except Exception as e:
                log(f"图片解码失败 {image_path}: {e}")
                continue

            # 提取位置/锚点信息
            anchor_info = {'type': 'inline'}
            if entry['type'] == 'drawing':
                anchor = entry['elem'].xpath('.//wp:anchor', namespaces=nsmap)
                if anchor:
                    anchor_elem = anchor[0]
                    anchor_info = {
                        'type': 'anchor',
                        'relativeFromH': anchor_elem.get(qn('wp:relativeFromH')),
                        'relativeFromV': anchor_elem.get(qn('wp:relativeFromV')),
                        'alignH': anchor_elem.get(qn('wp:alignH')),
                        'alignV': anchor_elem.get(qn('wp:alignV')),
                    }
                    pos_offsets = anchor_elem.xpath('.//wp:posOffset', namespaces=nsmap)
                    if len(pos_offsets) >= 1:
                        anchor_info['posOffsetH'] = pos_offsets[0].text
                    if len(pos_offsets) >= 2:
                        anchor_info['posOffsetV'] = pos_offsets[1].text

            images_info.append({
                'index': idx,
                'paragraph_index': entry['para_idx'],
                'section_index': 0,
                'anchor_info': anchor_info,
                'image': image,           # PIL Image，供 recognize_text 使用
                'image_data': image_data,  # 二进制，备用
                'rId': entry['rid'],
                'image_path': image_path,
            })

        docx_zip.close()
        log(f"成功提取 {len(images_info)} 张图片用于 OCR")
        return images_info

    def recognize_text(self, image):
        if not self.ocr:
            return ""

        try:
            img_array = np.array(image)
            result = self.ocr.ocr(img_array, cls=False)

            text_content = ""
            if result and result[0]:
                for line in result[0]:
                    if isinstance(line, (list, tuple)) and len(line) >= 2:
                        text_info = line[1]
                        if isinstance(text_info, (list, tuple)) and len(text_info) >= 2:
                            text = text_info[0]
                            score = text_info[1]
                            if score > 0.6:
                                text_content += text + " "

            return text_content.strip()

        except Exception as e:
            log(f"图片识别出错: {e}")
            import traceback
            log(f"详细错误: {traceback.format_exc()}")
            return ""

    def scan_document_with_positions(self, file_path):
        log("正在启动 OCR 扫描（含位置提取）...")
        images_info = self.extract_images_with_positions(file_path)
        if not images_info:
            log("文档中未发现图片，跳过 OCR 步骤")
            return []

        ocr_results = []
        for info in images_info:
            text = self.recognize_text(info['image'])
            ocr_results.append({
                'index': info['index'],
                'paragraph_index': info['paragraph_index'],
                'anchor_info': info['anchor_info'],
                'section_index': info['section_index'],
                'text': text,
            })
        log(f"OCR 完成，共处理 {len(ocr_results)} 张图片")
        return ocr_results


# ========== 测试代码 ==========
if __name__ == "__main__":
    print("=== 独立测试 OCR 引擎（含位置提取）===")
    engine = OCREngine()
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)
    test_file = os.path.join(project_root, "test.docx")

    if os.path.exists(test_file):
        results = engine.scan_document_with_positions(test_file)
        print("\n=== 识别结果（含位置信息）===")
        if results:
            for r in results:
                print(f"\n📷 图片 {r['index']}:")
                print(f"   所在段落索引: {r['paragraph_index']}")
                print(f"   所在节索引: {r['section_index']}")
                if r['anchor_info']:
                    print(f"   锚点信息: {r['anchor_info']}")
                else:
                    print("   锚点信息: 内联图片（无浮动定位）")
                print(f"   识别文字: {r['text'] if r['text'] else '(未识别到文字)'}")
                print(f"   文字长度: {len(r['text'])} 字符")
        else:
            print("未识别到任何图片或文字")
    else:
        print(f"测试文件不存在: {test_file}")
        print("请在项目根目录下放置一个包含图片的 test.docx 文件")
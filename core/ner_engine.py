# core/ner_engine.py
import hanlp
import re
from utils.helpers import log


class NEREngine:
    """
    基于 HanLP + 正则的混合实体识别引擎
    职责：从文本中提取暗标合规检查所需的各类敏感实体
    """

    def __init__(self):
        self.ner_model = None
        self.tokenizer = None
        self._load_models()

    def _load_models(self):
        """加载 HanLP 预训练模型（分词 + NER）"""
        try:
            log("正在初始化 HanLP 模型...")
            self.tokenizer = hanlp.load(hanlp.pretrained.tok.COARSE_ELECTRA_SMALL_ZH)
            self.ner_model = hanlp.load(hanlp.pretrained.ner.MSRA_NER_ELECTRA_SMALL_ZH)
            log("HanLP 模型加载成功")
        except Exception as e:
            log(f"❌ HanLP 模型加载失败: {e}")
            import traceback
            traceback.print_exc()

    def extract_entities(self, text_input, source_type="word", position_map=None):
        """
        核心方法：提取所有实体（增强版，带位置追踪）

        :param text_input: 字符串 或 字符串列表
        :param source_type: "word" 或 "ocr"
        :param position_map: 当 source_type="word" 时为段落索引列表；
                            当 source_type="ocr" 时为 ocr_results 中的条目列表
        :return: 实体列表 [{
            "text": "...", "type": "...",
            "source_sentence": "...",
            "source_type": "word/ocr",
            "position": 段落索引或图片索引
        }, ...]
        """
        if not text_input:
            return []

        # 统一转换为句子列表
        if isinstance(text_input, str):
            sentences = [text_input] if text_input.strip() else []
        else:
            sentences = [s for s in text_input if s and s.strip()]

        if not sentences:
            return []

        # 构建位置映射
        if position_map is None:
            position_map = list(range(len(sentences)))

        # 确保 position_map 与 sentences 一一对应
        if len(position_map) != len(sentences):
            position_map = list(range(len(sentences)))

        all_entities = []

        try:
            # --- 第一部分：HanLP 深度学习识别 (人名/地名/机构) ---
            for idx, sentence in enumerate(sentences):
                tokens = self.tokenizer(sentence)
                result = self.ner_model([tokens])
                entities = self._parse_hanlp_result(
                    result, sentence, source_type, position_map[idx]
                )
                all_entities.extend(entities)

            # --- 第二部分：正则匹配 (联系方式、证件号等) ---
            for idx, sentence in enumerate(sentences):
                regex_entities = self._extract_regex_entities(
                    sentence, source_type, position_map[idx]
                )
                all_entities.extend(regex_entities)

        except Exception as e:
            log(f"实体识别推理出错: {e}")

        return all_entities

    def _parse_hanlp_result(self, ner_result, source_sentence, source_type, position):
        """解析 HanLP NER 模型输出（增强位置信息）"""
        entities = []
        if not isinstance(ner_result, list) or len(ner_result) == 0:
            return entities

        sentence_entities = ner_result[0]
        if not isinstance(sentence_entities, list):
            return entities

        for ent in sentence_entities:
            if not isinstance(ent, (list, tuple)) or len(ent) < 4:
                continue

            entity_text = ent[0]
            entity_type = ent[1]

            # 过滤无效实体
            if not entity_text or not isinstance(entity_text, str) or len(entity_text) <= 1:
                continue

            # 类型标准化
            type_upper = entity_type.upper()
            if type_upper in ('PER', 'PERSON', 'NR'):
                normalized_type = 'PER'
            elif type_upper in ('ORG', 'ORGANIZATION', 'NT'):
                normalized_type = 'ORG'
            elif type_upper in ('LOC', 'LOCATION', 'NS'):
                normalized_type = 'LOC'
            else:
                normalized_type = 'OTHER'

            entities.append({
                "text": entity_text,
                "type": normalized_type,
                "source_sentence": source_sentence,
                "source_type": source_type,
                "position": position
            })

        return entities

    def _extract_regex_entities(self, text, source_type, position):
        """
        使用正则表达式提取非结构化敏感信息（严格版）
        """
        entities = []

        # 1. 手机号（中国）- 严格11位
        phone_pattern = r'(?<!\d)1[3-9]\d{9}(?!\d)'
        for match in re.finditer(phone_pattern, text):
            entities.append({
                "text": match.group(),
                "type": "PHONE",
                "source_sentence": text,
                "source_type": source_type,
                "position": position
            })

        # 2. 座机号 - 必须有区号+连字符格式，如 010-12345678
        tel_pattern = r'(?<!\d)0\d{2,3}-[2-9]\d{6,7}(?:-\d{1,4})?(?!\d)'
        for match in re.finditer(tel_pattern, text):
            entities.append({
                "text": match.group(),
                "type": "TEL",
                "source_sentence": text,
                "source_type": source_type,
                "position": position
            })

        # 3. 邮箱
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        for match in re.finditer(email_pattern, text):
            entities.append({
                "text": match.group(),
                "type": "EMAIL",
                "source_sentence": text,
                "source_type": source_type,
                "position": position
            })

        # 4. 身份证号（18位）
        id_pattern = r'(?<!\d)[1-9]\d{5}(?:18|19|20)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])\d{3}[\dXx](?!\d)'
        for match in re.finditer(id_pattern, text):
            entities.append({
                "text": match.group(),
                "type": "ID_CARD",
                "source_sentence": text,
                "source_type": source_type,
                "position": position
            })

        # 5. 统一社会信用代码 - 必须同时包含字母和数字，且18位
        credit_code_pattern = r'(?<![A-Za-z0-9])[0-9A-HJ-NPQRTUWXY]{18}(?![A-Za-z0-9])'
        for match in re.finditer(credit_code_pattern, text):
            code = match.group()
            # 必须同时包含字母和数字
            has_alpha = any(c.isalpha() for c in code)
            has_digit = any(c.isdigit() for c in code)
            if has_alpha and has_digit:
                entities.append({
                    "text": code,
                    "type": "CREDIT_CODE",
                    "source_sentence": text,
                    "source_type": source_type,
                    "position": position
                })

        # 6. 网址
        url_pattern = r'https?://[^\s,，。；;]+'
        for match in re.finditer(url_pattern, text):
            entities.append({
                "text": match.group(),
                "type": "URL",
                "source_sentence": text,
                "source_type": source_type,
                "position": position
            })

        return entities

    def scan_document(self, word_text_list=None, ocr_results=None):
        """
        综合扫描入口（增强版，保留位置信息）

        :param word_text_list: Word 解析出的段落文本列表
        :param ocr_results: OCR 识别结果列表（含位置信息）
        :return: {
            "entities": [完整实体列表，带位置],
            "summary": 按类型分组的字典（兼容旧接口）
        }
        """
        log("正在启动实体识别扫描...")

        all_entities = []

        # 收集 Word 文本（带段落索引）
        if word_text_list:
            valid_items = [(i, t) for i, t in enumerate(word_text_list) if len(t.strip()) > 5]
            if valid_items:
                indices, texts = zip(*valid_items)
                word_entities = self.extract_entities(
                    list(texts), source_type="word", position_map=list(indices)
                )
                all_entities.extend(word_entities)
                log(f"已载入 {len(valid_items)} 段 Word 文本")

        # 收集 OCR 文本（带图片段落索引）
        if ocr_results:
            valid_items = [(item.get("paragraph_index", -1), item.get("text", ""))
                           for item in ocr_results
                           if len(item.get("text", "").strip()) > 5]
            if valid_items:
                positions, texts = zip(*valid_items)
                ocr_entities = self.extract_entities(
                    list(texts), source_type="ocr", position_map=list(positions)
                )
                all_entities.extend(ocr_entities)
                log(f"已载入 {len(valid_items)} 段 OCR 文本")

        if not all_entities:
            log("没有识别到任何实体")
            return {
                "entities": [],
                "summary": {
                    "ORG": [], "PER": [], "LOC": [],
                    "PHONE": [], "TEL": [], "EMAIL": [],
                    "ID_CARD": [], "CREDIT_CODE": [], "URL": []
                }
            }

        # 去重并生成摘要
        result_summary = {
            "ORG": [], "PER": [], "LOC": [],
            "PHONE": [], "TEL": [], "EMAIL": [],
            "ID_CARD": [], "CREDIT_CODE": [], "URL": []
        }

        seen = set()
        unique_entities = []
        for ent in all_entities:
            key = f"{ent['type']}_{ent['text']}_{ent['source_type']}_{ent['position']}"
            if key not in seen:
                seen.add(key)
                unique_entities.append(ent)
                ent_type = ent['type']
                if ent_type in result_summary and ent['text'] not in result_summary[ent_type]:
                    result_summary[ent_type].append(ent['text'])

        return {
            "entities": unique_entities,
            "summary": result_summary
        }


# --- 独立运行测试代码 ---
if __name__ == "__main__":
    print("=== 独立测试 NER 引擎 ===")
    engine = NEREngine()

    test_text = [
        "华为技术有限公司在深圳发布了新产品。",
        "项目经理张三的电话是13800138000，邮箱是zhangsan@example.com。",
        "公司统一社会信用代码为91110000MA001B5M2Q。",
        "如有疑问请拨打010-12345678咨询。",
        "身份证号110101199001011234。",
        "公司网址是https://www.example.com。"
    ]

    # 测试带位置信息的扫描
    ocr_test = [
        {"paragraph_index": 10, "text": "李四有一家公司的董事长叫做王老五。"},
        {"paragraph_index": 15, "text": "华为有限公司在上海。"}
    ]

    results = engine.scan_document(word_text_list=test_text, ocr_results=ocr_test)

    print("\n=== 识别结果（带位置）===")
    for ent in results["entities"]:
        print(f"  [{ent['type']}] '{ent['text']}' @ {ent['source_type']} pos={ent['position']}")

    print("\n=== 摘要 ===")
    for category, items in results["summary"].items():
        if items:
            print(f"  {category}: {items}")
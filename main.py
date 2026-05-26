"""
main.py - 暗标合规检查系统入口（增强版：支持精确位置批注）
"""
import os
import sys
import argparse
import time
import glob

# 将项目根目录加入路径，确保能导入 core 和 utils
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from utils.helpers import log
from utils.file_handler import FileHandler

# ========== 默认配置 ==========
DEFAULT_INPUT_DIR = "clients/input"
DEFAULT_RULES = "clients/config/rules.json"
DEFAULT_OUTPUT = "clients/output"


def get_input_file(input_dir):
    """从输入目录中获取唯一的Word文件"""
    if not os.path.exists(input_dir):
        log(f"输入目录不存在，正在创建: {input_dir}")
        os.makedirs(input_dir, exist_ok=True)
        return None

    docx_files = glob.glob(os.path.join(input_dir, "*.docx"))

    if len(docx_files) == 0:
        log(f"❌ 输入目录中没有找到 .docx 文件: {input_dir}")
        return None

    if len(docx_files) > 1:
        log(f"⚠️  输入目录中发现多个 .docx 文件，将使用第一个: {os.path.basename(docx_files[0])}")
        for f in docx_files[1:]:
            log(f"   忽略: {os.path.basename(f)}")

    return docx_files[0]


class DarkMarkInspector:
    """暗标合规检查主控类"""

    def __init__(self, rules_path=DEFAULT_RULES):
        # 先导入并初始化 NER（依赖 PyTorch）
        log("正在初始化 NER 引擎（优先加载，避免依赖冲突）...")
        from core.ner_engine import NEREngine
        self.ner_engine = NEREngine()
        log("NER 引擎初始化完成")

        # 再导入并初始化 Format Checker
        log("正在初始化格式检查器...")
        from core.format_checker import FormatChecker
        self.format_checker = FormatChecker(rules_path)
        log("格式检查器初始化完成")

        # 最后导入并初始化 OCR
        log("正在初始化 OCR 引擎...")
        from core.ocr_engine import OCREngine
        self.ocr_engine = OCREngine()
        log("OCR 引擎初始化完成")

        # 初始化输出工具
        log("正在初始化输出工具...")
        from utils.report_gen import ReportGenerator
        from utils.annotator import DocumentAnnotator
        self.report_generator = ReportGenerator()
        self.annotator = DocumentAnnotator()
        log("输出工具初始化完成")

    def run_inspection(self, file_path, output_dir=DEFAULT_OUTPUT):
        """执行全流程检查"""
        log("=== 启动暗标合规性全流程检查 ===")
        overall_start_time = time.time()

        file_handler = FileHandler(project_root=PROJECT_ROOT)
        output_dir = file_handler.resolve_input_path(output_dir)
        os.makedirs(output_dir, exist_ok=True)

        # --- 第一步：提取 Word 文本 ---
        log("📝 正在提取 Word 文档文本...")
        step_start = time.time()
        try:
            from docx import Document
            doc = Document(file_path)
            word_texts = [para.text for para in doc.paragraphs if para.text.strip()]
            log(f"成功提取 {len(word_texts)} 段 Word 文本")
        except Exception as e:
            log(f"读取文档失败: {e}")
            return
        step_time = time.time() - step_start
        log(f"⏱️ 文本提取耗时: {step_time:.2f} 秒")

        # --- 第二步：格式检查 ---
        log("\n--- 开始执行格式检查 ---")
        step_start = time.time()
        format_issues = self.format_checker.check_document(file_path)
        step_time = time.time() - step_start
        log(f"⏱️ 格式检查耗时: {step_time:.2f} 秒")

        # 重置错误列表，供下次检查使用
        self.format_checker.errors = []

        # --- 第三步：OCR 图片提取 ---
        log("\n--- 开始执行 OCR 图片识别 ---")
        step_start = time.time()
        ocr_results = self.ocr_engine.scan_document_with_positions(file_path)
        ocr_texts = [item["text"] for item in ocr_results if item["text"].strip()]
        step_time = time.time() - step_start
        log(f"⏱️ OCR 识别耗时: {step_time:.2f} 秒")

        # --- 第四步：NER 敏感信息扫描 ---
        log("\n--- 开始执行敏感实体识别 (NER) ---")
        step_start = time.time()
        # 【关键修改】传入完整的 ocr_results（含位置信息），而非仅文本列表
        ner_full_results = self.ner_engine.scan_document(
            word_text_list=word_texts,
            ocr_results=ocr_results  # 传入完整结果，不是 ocr_texts
        )
        step_time = time.time() - step_start
        log(f"⏱️ NER 识别耗时: {step_time:.2f} 秒")

        # --- 第五步：生成输出文件 ---
        log("\n--- 正在生成输出文件 ---")
        step_start = time.time()

        # 生成 HTML 报告（使用摘要数据，兼容旧格式）
        report_path = self.report_generator.generate_html_report(
            file_path=file_path,
            format_issues=format_issues,
            ner_data=ner_full_results.get("summary", ner_full_results),  # 兼容
            ocr_data=ocr_results,
            output_dir=output_dir,
            suffix="_检查报告"
        )

        # 生成批注版 Word（传入完整实体数据，支持精确位置）
        annotated_path = self.annotator.generate_annotated_copy(
            original_path=file_path,
            format_issues=format_issues,
            ner_data=ner_full_results,  # 传入完整结果（含 entities 和 summary）
            ocr_data=ocr_results,
            output_dir=output_dir,
            suffix="_批注版"
        )

        step_time = time.time() - step_start
        log(f"⏱️ 输出文件生成耗时: {step_time:.2f} 秒")

        # --- 第六步：汇总报告 ---
        total_time = time.time() - overall_start_time
        self._generate_final_report(
            format_issues=format_issues,
            ner_data=ner_full_results.get("summary", ner_full_results),
            ocr_data=ocr_results,
            report_path=report_path,
            annotated_path=annotated_path,
            total_time=total_time,
            file_path=file_path
        )

    def _generate_final_report(self, format_issues, ner_data, ocr_data,
                               report_path, annotated_path, total_time, file_path):
        """生成最终的检查报告"""
        print("\n" + "=" * 60)
        print(" 📋 暗标合规检查最终报告")
        print("=" * 60)

        # 1. 格式检查结果
        print(f"\n📐 一、格式排版检查")
        print("-" * 40)
        if format_issues and len(format_issues) > 0:
            print(f"   ❌ 发现 {len(format_issues)} 处格式问题：")
            for i, issue in enumerate(format_issues, 1):
                print(f"   {i}. {issue}")
        else:
            print("   ✅ 格式检查通过，未发现问题")

        # 2. OCR 检查结果
        print(f"\n🖼️  二、图片内容检查 (OCR)")
        print("-" * 40)
        if ocr_data:
            print(f"   - 共扫描图片: {len(ocr_data)} 张")
            ocr_text_count = sum(len(item.get("text", "")) for item in ocr_data)
            print(f"   - 识别总字数: {ocr_text_count} 字")
            if ocr_text_count > 50:
                print(f"   ⚠️  警告: 图片中包含大量可识别文字，请确认是否符合暗标要求")
        else:
            print(f"   - 未发现图片或图片无法识别")

        # 3. NER 敏感信息结果
        print(f"\n🔍 三、敏感实体检测 (NER)")
        print("-" * 40)
        risk_score = 0
        if ner_data:
            has_risk = False
            if isinstance(ner_data, dict):
                if ner_data.get("ORG"):
                    print(f"   🏢 机构名 (高风险): {ner_data['ORG']}")
                    risk_score += len(ner_data["ORG"]) * 2
                    has_risk = True
                if ner_data.get("PER"):
                    print(f"   👤 人名 (高风险): {ner_data['PER']}")
                    risk_score += len(ner_data["PER"]) * 2
                    has_risk = True
                if ner_data.get("LOC"):
                    print(f"   📍 地名 (中风险): {ner_data['LOC']}")
                    risk_score += len(ner_data["LOC"])
                    has_risk = True
            if not has_risk:
                print(f"   ✅ 未检测到敏感实体")
        else:
            print(f"   ⚠️  NER 模块未返回结果（可能模型加载失败）")

        # 4. 输出文件
        print(f"\n📄 四、输出文件")
        print("-" * 40)
        if report_path and os.path.exists(report_path):
            print(f"   ✅ HTML 报告: {os.path.abspath(report_path)}")
        else:
            print(f"   ❌ HTML 报告生成失败")

        if annotated_path and os.path.exists(annotated_path):
            print(f"   ✅ 批注版 Word: {os.path.abspath(annotated_path)}")
        else:
            print(f"   ❌ 批注版生成失败")

        # 5. 最终判定
        print(f"\n🏆 五、最终判定")
        print("-" * 40)
        has_format_issues = format_issues and len(format_issues) > 0

        if risk_score == 0 and not has_format_issues:
            print("   ✅ 合格: 文档符合暗标排版及内容规范！")
        else:
            print("   ❌ 不合格: 文档存在违规内容或格式问题，请根据上述提示修改。")
            if has_format_issues:
                print(f"      - 格式问题: {len(format_issues)} 处")
            if risk_score > 0:
                print(f"      - 敏感信息风险评分: {risk_score} 分")

        # 6. 性能统计
        print(f"\n⏱️  六、性能统计")
        print("-" * 40)
        print(f"   - 总耗时: {total_time:.2f} 秒")
        print(f"   - 检查时间: {time.strftime('%Y-%m-%d %H:%M:%S')}")

        print("\n" + "=" * 60)
        print("⚠️  注意: 批注副本仅供内部修改参考，严禁用于投标！")
        print("=" * 60)


def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(
        description="暗标合规检查系统 - 检查 Word 文档是否符合暗标排版要求"
    )
    parser.add_argument(
        "--rules", "-r",
        default=DEFAULT_RULES,
        help=f"规则配置文件路径 (默认: {DEFAULT_RULES})"
    )
    parser.add_argument(
        "--output", "-o",
        default=DEFAULT_OUTPUT,
        help=f"输出目录路径 (默认: {DEFAULT_OUTPUT})"
    )
    args = parser.parse_args()

    file_handler = FileHandler(project_root=PROJECT_ROOT)
    input_dir = file_handler.resolve_input_path(DEFAULT_INPUT_DIR)
    rules_path = file_handler.resolve_input_path(args.rules)

    input_path = get_input_file(input_dir)
    if not input_path:
        print(f"\n❌ 错误: 无法在 input/ 目录中找到 .docx 文件")
        print(f"   请将要检查的 Word 文件放入: {input_dir}")
        print(f"   目录中只能放一个 .docx 文件")
        sys.exit(1)

    log(f"检测到输入文件: {os.path.basename(input_path)}")

    if not os.path.exists(rules_path):
        print(f"❌ 错误: 规则配置文件不存在: {rules_path}")
        print(f"请使用 --rules 指定正确的规则文件")
        sys.exit(1)

    inspector = DarkMarkInspector(rules_path=rules_path)
    inspector.run_inspection(file_path=input_path, output_dir=args.output)


if __name__ == "__main__":
    main()
"""
format_only.py - 暗标格式专项检查入口（轻量版）
===============================================
仅执行格式排版检查，不加载 NER/OCR 引擎。
检测报告包含：
  1. 检查维度清单（展示本次检查了哪些内容、是否通过）
  2. 问题明细表
  3. 输出文件清单（HTML 报告 + 批注版 Word）
"""

import os
import sys
import argparse
import time
import glob
import json
import re
from datetime import datetime
from docx.enum.text import WD_ALIGN_PARAGRAPH

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

# 兼容：若项目尚未提供 helpers / file_handler，则使用内置 fallback
try:
    from utils.helpers import log
except ImportError:
    def log(msg):
        print(f"[检查中] {msg}")

try:
    from utils.file_handler import FileHandler
except ImportError:
    class FileHandler:
        def __init__(self, project_root=None):
            self.project_root = project_root or PROJECT_ROOT

        def resolve_input_path(self, path):
            if os.path.isabs(path):
                return path
            return os.path.join(self.project_root, path)

from core.format_checker import FormatChecker

# ========== 【新增】导入批注生成器（仅依赖 python-docx，仍为轻量） ==========
try:
    from utils.annotator import DocumentAnnotator
except ImportError:
    DocumentAnnotator = None
    log("警告: 未找到 DocumentAnnotator，将跳过批注版 Word 生成")

DEFAULT_INPUT_DIR = "clients/blind_bid_tec_doc"
DEFAULT_RULES = "clients/config/rules.json"
DEFAULT_OUTPUT = "clients/output"


def get_input_file(input_dir):
    """从输入目录中获取唯一的 Word 文件"""
    if not os.path.exists(input_dir):
        log(f"输入目录不存在，正在创建: {input_dir}")
        os.makedirs(input_dir, exist_ok=True)
        return None

    docx_files = glob.glob(os.path.join(input_dir, "*.docx"))
    if len(docx_files) == 0:
        log(f"❌ 输入目录中没有找到 .docx 文件: {input_dir}")
        return None

    if len(docx_files) > 1:
        log(f"⚠️  发现多个 .docx 文件，将使用第一个: {os.path.basename(docx_files[0])}")
        for f in docx_files[1:]:
            log(f"   忽略: {os.path.basename(f)}")

    return docx_files[0]


class FormatOnlyInspector:
    """
    暗标格式专项检查器
    只执行 FormatChecker，不依赖 PyTorch / PaddleOCR 等重型库。
    """

    # 修改：将 "object_check" 改为 "table_check"，与配置文件键名保持一致
    CHECK_ITEMS = {
        "page_check": ("纸张大小与页数", "检查文档是否为 A4 纸张，以及页数是否可能超标"),
        "margin_check": ("页边距", "检查上/下/左/右边距是否符合暗标要求的固定值"),
        "font_check": ("字体、字号及样式", "检查中文字体、字号、颜色、加粗、倾斜、下划线及字符间空格"),
        "paragraph_check": ("段落格式", "检查对齐方式、行距模式与数值、首行缩进、左右缩进、段前段后间距及空段落"),
        "table_check": ("表格与图片格式", "检查表格整体对齐、表内文字样式、图片居中对齐等"),
        "structure_check": ("文档结构", "检查页眉、页脚、页码等是否违规存在"),
        "punctuation_check": ("标点符号规范", "检查是否存在英文标点混用在中文语境中的情况"),
    }

    def __init__(self, rules_path=DEFAULT_RULES):
        log("正在初始化格式检查器...")
        self.format_checker = FormatChecker(rules_path)
        log("格式检查器初始化完成")

        # 【新增】初始化批注生成器
        if DocumentAnnotator is not None:
            log("正在初始化批注生成器...")
            self.annotator = DocumentAnnotator()
            log("批注生成器初始化完成")
        else:
            self.annotator = None

        self.rules_path = rules_path
        with open(rules_path, "r", encoding="utf-8") as f:
            self.rules = json.load(f)

    def run_inspection(self, file_path, output_dir=DEFAULT_OUTPUT):
        """执行格式检查并生成增强报告 + 批注版 Word（先出报告，后出批注）"""
        log("=== 启动暗标格式专项检查 ===")
        overall_start = time.time()

        file_handler = FileHandler(project_root=PROJECT_ROOT)
        output_dir = file_handler.resolve_input_path(output_dir)
        os.makedirs(output_dir, exist_ok=True)

        # 1. 执行格式检查
        log("📝 正在提取并检查 Word 文档格式...")
        step_start = time.time()
        self.format_checker.errors = []
        format_issues = self.format_checker.check_document(file_path)
        step_time = time.time() - step_start
        log(f"⏱️ 格式检查耗时: {step_time:.2f} 秒，发现 {len(format_issues)} 处问题")

        # 2. 按维度归类问题
        classified = self._classify_issues(format_issues)

        # 3. 生成检查清单
        checklist = self._build_checklist(classified)

        # 4. 【修改】先生成 HTML 报告（不等待批注版）
        log("\n--- 正在生成 HTML 报告 ---")
        report_path = self._generate_report(
            file_path=file_path,
            format_issues=format_issues,
            checklist=checklist,
            annotated_path=None,      # 此时批注版尚未生成
            output_dir=output_dir,
        )

        # 5. 【新增】后台生成批注版 Word（不影响报告）
        annotated_path = None
        if self.annotator is not None:
            log("\n--- 正在后台生成批注版 Word（不会阻塞报告查看） ---")
            step_start = time.time()
            try:
                annotated_path = self.annotator.generate_annotated_copy(
                    original_path=file_path,
                    format_issues=format_issues,
                    ner_data=None,
                    ocr_data=None,
                    output_dir=output_dir,
                    suffix="_批注版"
                )
                step_time = time.time() - step_start
                log(f"⏱️ 批注版生成耗时: {step_time:.2f} 秒")
                log(f"📁 批注版 Word 保存至: {annotated_path}")
            except Exception as e:
                log(f"❌ 批注版生成失败: {e}")
        else:
            log("⚠️  批注生成器未加载，跳过批注版 Word 输出")

        # 6. 控制台汇总（包含报告路径和批注版路径）
        total_time = time.time() - overall_start
        self._print_summary(format_issues, checklist, report_path, annotated_path, total_time)

        return report_path

    # ------------------------------------------------------------------
    # 问题归类
    # ------------------------------------------------------------------
    def _classify_issues(self, issues):
        """将错误列表按检查维度归类"""
        buckets = {key: [] for key in self.CHECK_ITEMS.keys()}
        buckets["other"] = []

        for issue in issues:
            target = self._classify_single_issue(issue)
            buckets[target].append(issue)
        return buckets

    def _classify_single_issue(self, issue):
        """根据错误文本特征判断所属维度"""
        if any(k in issue for k in ["纸张", "页数", "A4"]):
            return "page_check"
        if "边距" in issue:
            return "margin_check"
        if any(k in issue for k in ["字体", "字号", "颜色", "加粗", "倾斜", "下划线", "字符间有空格"]):
            return "font_check"
        if any(k in issue for k in ["对齐方式", "行距", "缩进", "间距", "空段落", "空格或", "Tab"]):
            return "paragraph_check"
        if any(k in issue for k in ["表格", "图片"]):
            return "table_check"
        if any(k in issue for k in ["页眉", "页脚", "页码"]):
            return "structure_check"
        if "标点" in issue:
            return "punctuation_check"
        return "other"

    # ------------------------------------------------------------------
    # 检查清单构建
    # ------------------------------------------------------------------
    def _build_checklist(self, classified):
        checklist = []
        for key, (name, desc) in self.CHECK_ITEMS.items():
            enabled = self.rules.get(key, {}).get("enabled", False)
            issues = classified.get(key, [])
            count = len(issues)
            checklist.append({
                "key": key,
                "name": name,
                "enabled": enabled,
                "passed": (count == 0),
                "issue_count": count,
                "issues": issues[:3],
                "desc": desc,
            })
        return checklist

    # ------------------------------------------------------------------
    # HTML 报告生成（【修改】支持 annotated_path 为 None 的情况）
    # ------------------------------------------------------------------
    def _generate_report(self, file_path, format_issues, checklist, annotated_path, output_dir):
        file_name = os.path.basename(file_path)
        check_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        total_issues = len(format_issues)
        total_checks = sum(1 for c in checklist if c["enabled"])
        passed_checks = sum(1 for c in checklist if c["enabled"] and c["passed"])
        failed_checks = total_checks - passed_checks
        passed = (total_issues == 0)

        checklist_rows = self._render_checklist_rows(checklist)

        if format_issues:
            issue_rows = self._render_issue_rows(format_issues)
            issues_section = f"""
            <div class="section">
                <div class="section-title">🔍 问题明细</div>
                <div style="overflow-x: auto;">
                    <table>
                        <thead>
                            <tr>
                                <th style="width:40px">#</th>
                                <th style="width:90px">位置</th>
                                <th>问题描述</th>
                                <th style="width:130px">当前值</th>
                                <th style="width:130px">要求值</th>
                            </tr>
                        </thead>
                        <tbody>{issue_rows}</tbody>
                    </table>
                </div>
            </div>
            """
        else:
            issues_section = """
            <div class="section">
                <div class="section-title">🔍 问题明细</div>
                <div class="empty-state">
                    <div class="empty-icon">🎉</div>
                    <div>未发现任何格式问题，所有启用检查项均通过！</div>
                </div>
            </div>
            """

        # 【修改】输出文件区域 HTML：根据 annotated_path 状态显示不同内容
        if annotated_path and os.path.exists(annotated_path):
            anno_basename = os.path.basename(annotated_path)
            output_files_html = f"""
            <div class="section">
                <div class="section-title">📁 输出文件</div>
                <div class="file-list">
                    <div class="file-item">
                        <span class="file-icon">📄</span>
                        <div>
                            <div class="file-name">HTML 检查报告</div>
                            <div class="file-path">{os.path.basename(file_name)}_格式检查报告.html</div>
                        </div>
                    </div>
                    <div class="file-item">
                        <span class="file-icon">📝</span>
                        <div>
                            <div class="file-name">批注版 Word（仅含格式批注）</div>
                            <div class="file-path">{anno_basename}</div>
                        </div>
                    </div>
                </div>
            </div>
            """
        elif annotated_path is None:
            # 情况：批注版尚未生成（正在后台生成中）
            output_files_html = f"""
            <div class="section">
                <div class="section-title">📁 输出文件</div>
                <div class="file-list">
                    <div class="file-item">
                        <span class="file-icon">📄</span>
                        <div>
                            <div class="file-name">HTML 检查报告</div>
                            <div class="file-path">{os.path.basename(file_name)}_格式检查报告.html</div>
                        </div>
                    </div>
                    <div class="file-item">
                        <span class="file-icon">⏳</span>
                        <div>
                            <div class="file-name">批注版 Word</div>
                            <div class="file-path">正在后台生成中，请稍后查看输出目录...</div>
                        </div>
                    </div>
                </div>
            </div>
            """
        else:
            # 生成失败或未加载批注器
            output_files_html = f"""
            <div class="section">
                <div class="section-title">📁 输出文件</div>
                <div class="file-list">
                    <div class="file-item">
                        <span class="file-icon">📄</span>
                        <div>
                            <div class="file-name">HTML 检查报告</div>
                            <div class="file-path">{os.path.basename(file_name)}_格式检查报告.html</div>
                        </div>
                    </div>
                    <div class="file-item disabled">
                        <span class="file-icon">⚠️</span>
                        <div>
                            <div class="file-name">批注版 Word</div>
                            <div class="file-path">未生成（批注生成器未加载或生成失败）</div>
                        </div>
                    </div>
                </div>
            </div>
            """

        status_icon = "✅" if passed else "❌"
        status_text = "检查通过" if passed else "发现违规"
        status_color = "#27ae60" if passed else "#e74c3c"

        html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>暗标格式专项检查报告 - {file_name}</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            background: #f5f6fa;
            color: #2c3e50;
            line-height: 1.6;
            padding: 20px;
        }}
        .container {{
            max-width: 960px;
            margin: 0 auto;
            background: #fff;
            border-radius: 12px;
            box-shadow: 0 2px 12px rgba(0,0,0,0.08);
            overflow: hidden;
        }}
        .header {{
            background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }}
        .header h1 {{ font-size: 24px; margin-bottom: 8px; }}
        .header .meta {{ font-size: 14px; opacity: 0.9; }}
        .header .notice {{
            margin-top: 8px;
            font-size: 12px;
            opacity: 0.85;
            background: rgba(255,255,255,0.15);
            display: inline-block;
            padding: 4px 12px;
            border-radius: 20px;
        }}
        .status-card {{
            padding: 25px 30px;
            text-align: center;
            border-bottom: 1px solid #eee;
        }}
        .status-badge {{
            display: inline-block;
            padding: 10px 28px;
            border-radius: 30px;
            font-size: 18px;
            font-weight: bold;
            color: white;
            background: {status_color};
            margin-bottom: 15px;
        }}
        .stats {{
            display: flex;
            justify-content: center;
            gap: 30px;
            flex-wrap: wrap;
        }}
        .stat-item {{ text-align: center; min-width: 80px; }}
        .stat-num {{ font-size: 28px; font-weight: bold; color: {status_color}; }}
        .stat-label {{ font-size: 13px; color: #7f8c8d; margin-top: 4px; }}
        .section {{
            padding: 25px 30px;
            border-bottom: 1px solid #eee;
        }}
        .section-title {{
            font-size: 16px;
            font-weight: bold;
            color: #2c3e50;
            margin-bottom: 12px;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .section-desc {{
            font-size: 13px;
            color: #7f8c8d;
            margin-bottom: 15px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
        }}
        th {{
            background: #f8f9fa;
            padding: 12px 10px;
            text-align: left;
            font-weight: 600;
            color: #555;
            border-bottom: 2px solid #e0e0e0;
        }}
        td {{
            padding: 12px 10px;
            border-bottom: 1px solid #f0f0f0;
            vertical-align: top;
        }}
        tr:hover {{ background: #fafbfc; }}
        .current {{ color: #e74c3c; font-weight: 500; }}
        .required {{ color: #27ae60; font-weight: 500; }}
        
        /* 检查清单样式 */
        .checklist-table tr.passed {{
            border-left: 4px solid #27ae60;
            background: #f6fff9;
        }}
        .checklist-table tr.failed {{
            border-left: 4px solid #e74c3c;
            background: #fff6f6;
        }}
        .checklist-table tr.disabled {{
            border-left: 4px solid #bdc3c7;
            background: #fafafa;
            color: #95a5a6;
        }}
        .tag {{
            display: inline-block;
            padding: 2px 10px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: 600;
        }}
        .tag-pass {{ background: #d4edda; color: #155724; }}
        .tag-fail {{ background: #f8d7da; color: #721c24; }}
        .tag-disabled {{ background: #e2e3e5; color: #383d41; }}
        .issue-preview {{
            font-size: 12px;
            color: #666;
            margin-top: 4px;
            line-height: 1.4;
        }}
        .issue-preview code {{
            background: #f1f2f6;
            padding: 1px 4px;
            border-radius: 3px;
            font-family: monospace;
            color: #e74c3c;
        }}
        
        /* 【新增】输出文件列表样式 */
        .file-list {{
            display: flex;
            flex-direction: column;
            gap: 12px;
        }}
        .file-item {{
            display: flex;
            align-items: center;
            gap: 12px;
            padding: 12px 16px;
            background: #f8f9fa;
            border-radius: 8px;
            border-left: 4px solid #667eea;
        }}
        .file-item.disabled {{
            border-left-color: #bdc3c7;
            opacity: 0.7;
        }}
        .file-icon {{ font-size: 24px; }}
        .file-name {{ font-weight: 600; font-size: 14px; color: #2c3e50; }}
        .file-path {{ font-size: 12px; color: #7f8c8d; margin-top: 2px; font-family: monospace; }}
        
        .empty-state {{
            text-align: center;
            padding: 40px 20px;
            color: #27ae60;
            font-size: 15px;
            font-weight: 500;
        }}
        .empty-icon {{ font-size: 40px; margin-bottom: 10px; }}
        .footer {{
            padding: 20px 30px;
            text-align: center;
            font-size: 12px;
            color: #95a5a6;
            background: #fafbfc;
        }}
        @media (max-width: 600px) {{
            body {{ padding: 10px; }}
            .header {{ padding: 20px; }}
            .header h1 {{ font-size: 20px; }}
            .section {{ padding: 15px; }}
            th, td {{ padding: 8px 6px; font-size: 13px; }}
            .stats {{ gap: 15px; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📐 暗标格式专项检查报告</h1>
            <div class="meta">{file_name} &nbsp;|&nbsp; 检查时间: {check_time}</div>
            <div class="notice">ℹ️ 本报告仅包含格式排版检查，未进行 OCR 图片识别与 NER 敏感信息扫描</div>
        </div>

        <div class="status-card">
            <div class="status-badge">{status_icon} {status_text}</div>
            <div class="stats">
                <div class="stat-item">
                    <div class="stat-num">{total_issues}</div>
                    <div class="stat-label">格式问题</div>
                </div>
                <div class="stat-item">
                    <div class="stat-num">{total_checks}</div>
                    <div class="stat-label">检查维度</div>
                </div>
                <div class="stat-item">
                    <div class="stat-num">{passed_checks}</div>
                    <div class="stat-label">通过项</div>
                </div>
                <div class="stat-item">
                    <div class="stat-num">{failed_checks}</div>
                    <div class="stat-label">未通过项</div>
                </div>
            </div>
        </div>

        <div class="section">
            <div class="section-title">📋 检查维度清单</div>
            <p class="section-desc">以下列出本次检查启用的全部维度及结果。灰色项表示在 rules.json 中未启用，不在本次检查范围内。</p>
            <table class="checklist-table">
                <thead>
                    <tr>
                        <th style="width:160px">检查维度</th>
                        <th style="width:70px">状态</th>
                        <th style="width:70px">问题数</th>
                        <th>说明 / 违规预览</th>
                    </tr>
                </thead>
                <tbody>
                    {checklist_rows}
                </tbody>
            </table>
        </div>

        {issues_section}

        {output_files_html}

        <div class="footer">
            本报告由暗标合规检查系统自动生成（格式专项版）&nbsp;|&nbsp; 仅供参考使用
        </div>
    </div>
</body>
</html>"""

        base_name = os.path.splitext(file_name)[0]
        output_path = os.path.join(output_dir, f"{base_name}_格式检查报告.html")
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html)

        log(f"HTML 报告已生成: {output_path}")
        return output_path

    def _render_checklist_rows(self, checklist):
        rows = []
        for item in checklist:
            if not item["enabled"]:
                status_html = '<span class="tag tag-disabled">未启用</span>'
                count_html = "-"
                preview = "本次检查未启用该维度"
                row_class = "disabled"
            elif item["passed"]:
                status_html = '<span class="tag tag-pass">✅ 通过</span>'
                count_html = "0"
                preview = item["desc"]
                row_class = "passed"
            else:
                status_html = '<span class="tag tag-fail">❌ 未通过</span>'
                count_html = str(item["issue_count"])
                snippets = [f"<code>{self._escape_html(iss[:60])}</code>" for iss in item["issues"]]
                preview = f"{item['desc']}<br><div class='issue-preview'>典型问题：" + " | ".join(snippets) + "</div>"
                row_class = "failed"

            rows.append(f"""
            <tr class="{row_class}">
                <td><strong>{item['name']}</strong></td>
                <td>{status_html}</td>
                <td>{count_html}</td>
                <td>{preview}</td>
            </tr>
            """)
        return "\n".join(rows)

    def _render_issue_rows(self, issues):
        rows = []
        for idx, issue in enumerate(issues, 1):
            loc = self._extract_location(issue)
            desc = self._extract_description(issue)
            cur, req = self._extract_values(issue)
            rows.append(f"""
            <tr>
                <td>{idx}</td>
                <td>{loc}</td>
                <td>{self._escape_html(desc)}</td>
                <td class="current">{self._escape_html(cur)}</td>
                <td class="required">{self._escape_html(req)}</td>
            </tr>
            """)
        return "\n".join(rows)

    @staticmethod
    def _extract_location(issue):
        match = re.search(r"段落\s+(\d+)", issue)
        if match:
            return f"段落 {match.group(1)}"
        match = re.search(r"第\s+(\d+)\s+节", issue)
        if match:
            return f"第 {match.group(1)} 节"
        match = re.search(r"表格\s+(\d+)", issue)
        if match:
            return f"表格 {match.group(1)}"
        return "全文"

    @staticmethod
    def _extract_description(issue):
        desc = re.sub(r"^(段落\s+\d+:\s*)", "", issue)
        desc = re.sub(r"^(第\s+\d+\s+节[：:]\s*)", "", desc)
        desc = re.sub(r"^(表格\s+\d+\s+)", "", desc)
        return desc

    @staticmethod
    def _extract_values(issue):
        match = re.search(r"当前[：:]\s*([^,，]+)[,，]\s*(?:要求|应为)[：:]\s*(.+)", issue)
        if match:
            return match.group(1).strip(), match.group(2).strip()
        match = re.search(r"当前[：:]\s*([^)]+)\)", issue)
        if match:
            return match.group(1).strip(), "—"
        match = re.search(r"发现\s+(.+)", issue)
        if match:
            return match.group(1).strip(), "不得出现"
        return "—", "—"

    @staticmethod
    def _escape_html(text):
        return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    # ------------------------------------------------------------------
    # 【修改】控制台汇总增加 annotated_path 参数
    # ------------------------------------------------------------------
    def _print_summary(self, format_issues, checklist, report_path, annotated_path, total_time):
        print("\n" + "=" * 60)
        print(" 📐 暗标格式专项检查报告（控制台版）")
        print("=" * 60)

        print(f"\n📋 检查维度清单（共 {sum(1 for c in checklist if c['enabled'])} 项启用）")
        print("-" * 60)
        for item in checklist:
            if not item["enabled"]:
                continue
            icon = "✅" if item["passed"] else "❌"
            print(f"   {icon} {item['name']:<12}  问题数: {item['issue_count']}")

        print(f"\n🔍 问题明细")
        print("-" * 60)
        if format_issues:
            for i, issue in enumerate(format_issues, 1):
                print(f"   {i}. {issue}")
        else:
            print("   ✅ 未发现任何格式问题")

        # 【新增】输出文件区域
        print(f"\n📁 输出文件")
        print("-" * 60)
        if report_path and os.path.exists(report_path):
            print(f"   ✅ HTML 报告: {os.path.abspath(report_path)}")
        else:
            print(f"   ❌ HTML 报告生成失败")

        if annotated_path and os.path.exists(annotated_path):
            print(f"   ✅ 批注版 Word: {os.path.abspath(annotated_path)}")
        else:
            print(f"   ⚠️  批注版 Word 未生成")

        print(f"\n⏱️  性能统计")
        print("-" * 60)
        print(f"   - 总耗时: {total_time:.2f} 秒")
        print(f"   - 检查时间: {time.strftime('%Y-%m-%d %H:%M:%S')}")

        print("\n" + "=" * 60)
        print("⚠️  提示: 本脚本仅检查格式，未扫描图片文字及敏感实体")
        print("=" * 60)


def main():
    parser = argparse.ArgumentParser(
        description="暗标格式专项检查 - 仅检查 Word 文档排版格式，不加载 AI 模型"
    )
    parser.add_argument("--rules", "-r", default=DEFAULT_RULES, help="规则配置文件路径")
    parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT, help="输出目录路径")
    args = parser.parse_args()

    file_handler = FileHandler(project_root=PROJECT_ROOT)
    input_dir = file_handler.resolve_input_path(DEFAULT_INPUT_DIR)
    rules_path = file_handler.resolve_input_path(args.rules)

    input_path = get_input_file(input_dir)
    if not input_path:
        print(f"\n❌ 错误: 无法在 input/ 目录中找到 .docx 文件")
        sys.exit(1)

    log(f"检测到输入文件: {os.path.basename(input_path)}")

    if not os.path.exists(rules_path):
        print(f"❌ 错误: 规则配置文件不存在: {rules_path}")
        sys.exit(1)

    inspector = FormatOnlyInspector(rules_path=rules_path)
    inspector.run_inspection(file_path=input_path, output_dir=args.output)


if __name__ == "__main__":
    main()
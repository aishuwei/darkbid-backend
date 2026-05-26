"""
utils/report_gen.py - HTML 报告生成器
职责：将检查结果生成简洁的 HTML 报告，手机/电脑浏览器均可打开
"""
import os
import re
from datetime import datetime


def log(msg):
    print(f"[检查中] {msg}")


class ReportGenerator:
    """
    HTML 报告生成器
    生成单文件 HTML，内嵌 CSS，无需外部依赖，手机浏览器直接打开。
    """

    def __init__(self):
        pass

    def generate_html_report(self, file_path, format_issues, ner_data=None, ocr_data=None,
                             output_dir=None, suffix="_检查报告"):
        """
        生成 HTML 报告
        :param file_path: 原始文档路径（用于显示文件名）
        :param format_issues: 格式检查错误列表
        :param ner_data: NER 结果（可选）
        :param ocr_data: OCR 结果（可选）
        :param output_dir: 输出目录
        :param suffix: 输出文件名后缀
        :return: 输出文件路径
        """
        if output_dir is None:
            output_dir = os.path.dirname(os.path.abspath(file_path))
        os.makedirs(output_dir, exist_ok=True)

        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_path = os.path.join(output_dir, f"{base_name}{suffix}.html")

        # 统计数据
        format_count = len(format_issues) if format_issues else 0
        ner_count = self._count_ner_issues(ner_data)
        ocr_count = len(ocr_data) if ocr_data else 0
        total_issues = format_count + ner_count

        # 构建问题明细表格行
        table_rows = self._build_table_rows(format_issues, ner_data, ocr_data)

        # 生成 HTML
        html = self._build_html(
            file_name=os.path.basename(file_path),
            check_time=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            format_count=format_count,
            ner_count=ner_count,
            ocr_count=ocr_count,
            total_issues=total_issues,
            table_rows=table_rows,
            passed=(total_issues == 0)
        )

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html)

        log(f"HTML 报告已生成: {output_path}")
        return output_path

    def _count_ner_issues(self, ner_data):
        if not ner_data:
            return 0
        count = 0
        for key in ["ORG", "PER", "LOC"]:
            if ner_data.get(key):
                count += len(ner_data[key])
        return count

    def _build_table_rows(self, format_issues, ner_data, ocr_data):
        """构建 HTML 表格行"""
        rows = []
        idx = 1

        # 格式问题
        if format_issues:
            for issue in format_issues:
                location = self._extract_location(issue)
                description = self._extract_description(issue)
                current_val, required_val = self._extract_values(issue)

                rows.append("""
                <tr>
                    <td>%d</td>
                    <td>%s</td>
                    <td>%s</td>
                    <td class="current">%s</td>
                    <td class="required">%s</td>
                </tr>
                """ % (idx, location, description, current_val, required_val))
                idx += 1

        # NER 问题
        if ner_data:
            for key, label in [("ORG", "机构名"), ("PER", "人名"), ("LOC", "地名")]:
                if ner_data.get(key):
                    for entity in ner_data[key]:
                        rows.append("""
                        <tr>
                            <td>%d</td>
                            <td>全文</td>
                            <td>敏感实体检测：发现 %s</td>
                            <td class="current">%s</td>
                            <td class="required">不得出现</td>
                        </tr>
                        """ % (idx, label, entity))
                        idx += 1

        # OCR 问题
        if ocr_data:
            total_chars = sum(len(item.get("text", "")) for item in ocr_data)
            if total_chars > 50:
                rows.append("""
                <tr>
                    <td>%d</td>
                    <td>图片区域</td>
                    <td>图片中包含可识别文字</td>
                    <td class="current">%d 张图片 / %d 字</td>
                    <td class="required">建议核实</td>
                </tr>
                """ % (idx, len(ocr_data), total_chars))
                idx += 1

        if not rows:
            rows.append("""
            <tr>
                <td colspan="5" style="text-align:center; color: #27ae60; font-weight: bold;">
                    ✅ 未发现任何问题，文档符合暗标排版要求！
                </td>
            </tr>
            """)

        return "\n".join(rows)

    def _extract_location(self, issue):
        """从错误文本中提取位置信息"""
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

    def _extract_description(self, issue):
        """提取问题描述（去掉位置前缀）"""
        desc = re.sub(r"^(段落\s+\d+:\s*)", "", issue)
        desc = re.sub(r"^(第\s+\d+\s+节[：:]\s*)", "", desc)
        desc = re.sub(r"^(表格\s+\d+\s+)", "", desc)
        return desc

    def _extract_values(self, issue):
        """从错误文本中提取当前值和要求值"""
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

    def _build_html(self, file_name, check_time, format_count, ner_count, ocr_count,
                    total_issues, table_rows, passed):
        """构建完整 HTML 页面"""
        status_icon = "✅" if passed else "❌"
        status_text = "检查通过" if passed else "发现违规"
        status_color = "#27ae60" if passed else "#e74c3c"

        html_template = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>暗标合规检查报告 - %s</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            background: #f5f6fa;
            color: #2c3e50;
            line-height: 1.6;
            padding: 20px;
        }
        .container {
            max-width: 900px;
            margin: 0 auto;
            background: #fff;
            border-radius: 12px;
            box-shadow: 0 2px 12px rgba(0,0,0,0.08);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #667eea 0%%, #764ba2 100%%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        .header h1 { font-size: 24px; margin-bottom: 8px; }
        .header .meta { font-size: 14px; opacity: 0.9; }
        .status-card {
            padding: 25px 30px;
            text-align: center;
            border-bottom: 1px solid #eee;
        }
        .status-badge {
            display: inline-block;
            padding: 12px 30px;
            border-radius: 30px;
            font-size: 18px;
            font-weight: bold;
            color: white;
            background: %s;
            margin-bottom: 15px;
        }
        .stats {
            display: flex;
            justify-content: center;
            gap: 30px;
            flex-wrap: wrap;
        }
        .stat-item { text-align: center; }
        .stat-num { font-size: 28px; font-weight: bold; color: %s; }
        .stat-label { font-size: 13px; color: #7f8c8d; margin-top: 4px; }
        .section {
            padding: 25px 30px;
            border-bottom: 1px solid #eee;
        }
        .section-title {
            font-size: 16px;
            font-weight: bold;
            color: #2c3e50;
            margin-bottom: 15px;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        table {
            width: 100%%;
            border-collapse: collapse;
            font-size: 14px;
        }
        th {
            background: #f8f9fa;
            padding: 12px 10px;
            text-align: left;
            font-weight: 600;
            color: #555;
            border-bottom: 2px solid #e0e0e0;
            position: sticky;
            top: 0;
        }
        td {
            padding: 12px 10px;
            border-bottom: 1px solid #f0f0f0;
            vertical-align: top;
        }
        tr:hover { background: #fafbfc; }
        .current { color: #e74c3c; }
        .required { color: #27ae60; }
        .footer {
            padding: 20px 30px;
            text-align: center;
            font-size: 12px;
            color: #95a5a6;
            background: #fafbfc;
        }
        @media (max-width: 600px) {
            body { padding: 10px; }
            .header { padding: 20px; }
            .header h1 { font-size: 20px; }
            .section { padding: 15px; }
            th, td { padding: 8px 6px; font-size: 13px; }
            .stats { gap: 15px; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📋 暗标合规检查报告</h1>
            <div class="meta">%s &nbsp;|&nbsp; 检查时间: %s</div>
        </div>

        <div class="status-card">
            <div class="status-badge">%s %s</div>
            <div class="stats">
                <div class="stat-item">
                    <div class="stat-num">%d</div>
                    <div class="stat-label">总问题数</div>
                </div>
                <div class="stat-item">
                    <div class="stat-num">%d</div>
                    <div class="stat-label">格式问题</div>
                </div>
                <div class="stat-item">
                    <div class="stat-num">%d</div>
                    <div class="stat-label">敏感信息</div>
                </div>
                <div class="stat-item">
                    <div class="stat-num">%d</div>
                    <div class="stat-label">图片数量</div>
                </div>
            </div>
        </div>

        <div class="section">
            <div class="section-title">🔍 问题明细</div>
            <div style="overflow-x: auto;">
                <table>
                    <thead>
                        <tr>
                            <th style="width:40px">#</th>
                            <th style="width:80px">位置</th>
                            <th>问题描述</th>
                            <th style="width:120px">当前值</th>
                            <th style="width:120px">要求值</th>
                        </tr>
                    </thead>
                    <tbody>
                        %s
                    </tbody>
                </table>
            </div>
        </div>

        <div class="footer">
            本报告由暗标合规检查系统自动生成，仅供参考使用
        </div>
    </div>
</body>
</html>"""

        return html_template % (
            file_name, status_color, status_color,
            file_name, check_time, status_icon, status_text,
            total_issues, format_count, ner_count, ocr_count,
            table_rows
        )

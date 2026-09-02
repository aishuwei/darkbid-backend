"""
utils/report_gen.py - HTML 报告生成器（规则驱动版）
职责：将检查结果生成简洁的 HTML 报告，包含规则清单和问题明细
"""
import os
import re
from datetime import datetime


def log(msg):
    print(f"[检查中] {msg}")


class ReportGenerator:
    def __init__(self):
        pass

    def generate_html_report(self, file_path, format_issues, rule_summary=None, ner_data=None, ocr_data=None,
                             output_dir=None, suffix="_检查报告"):
        """
        生成 HTML 报告
        :param rule_summary: 规则清单列表（由 format_checker 提供）
        """
        if output_dir is None:
            output_dir = os.path.dirname(os.path.abspath(file_path))
        os.makedirs(output_dir, exist_ok=True)

        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_path = os.path.join(output_dir, f"{base_name}{suffix}.html")

        format_count = len(format_issues) if format_issues else 0
        ner_count = self._count_ner_issues(ner_data)
        ocr_count = len(ocr_data) if ocr_data else 0
        total_issues = format_count + ner_count

        # 构建规则清单 HTML
        rule_html = self._build_rule_section(rule_summary)

        # 构建问题明细表格行
        table_rows = self._build_table_rows(format_issues, ner_data, ocr_data)

        html = self._build_html(
            file_name=os.path.basename(file_path),
            check_time=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            format_count=format_count,
            ner_count=ner_count,
            ocr_count=ocr_count,
            total_issues=total_issues,
            rule_html=rule_html,
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
        if "entities" in ner_data:
            valid = {"ORG", "PER", "LOC", "PHONE", "TEL", "EMAIL", "ID_CARD", "CREDIT_CODE", "URL"}
            return sum(1 for e in ner_data["entities"] if e.get("type") in valid)
        return sum(len(ner_data.get(k, [])) for k in ["ORG", "PER", "LOC"])

    def _build_rule_section(self, rule_summary):
        if not rule_summary:
            return '<div class="section"><div class="section-title">📜 检测规则</div><p>未提供规则清单，请检查配置。</p></div>'
        rows = []
        for line in rule_summary:
            if line.startswith("==========") or line.startswith("规则名称") or line.startswith("规则描述") or line.startswith("纸张规格"):
                continue
            # 解析每行：✅ 或 ❌ 开头
            if "✅" in line or "❌" in line:
                parts = line.split(":", 1)
                if len(parts) == 2:
                    status = parts[0].strip()
                    detail = parts[1].strip()
                    icon = "✅" if "✅" in status else "❌"
                    status_text = "启用" if "✅" in status else "未启用"
                    rows.append(f'<tr><td>{icon}</td><td>{detail}</td><td>{status_text}</td></tr>')
        if not rows:
            return '<div class="section"><div class="section-title">📜 检测规则</div><p>无法解析规则清单</p></div>'
        table = f"""
        <div class="section">
            <div class="section-title">📜 检测规则清单</div>
            <div style="overflow-x: auto;">
                <table>
                    <thead><tr><th style="width:60px">状态</th><th>检测项</th><th style="width:80px">是否启用</th></tr></thead>
                    <tbody>{"".join(rows)}</tbody>
                </table>
            </div>
        </div>
        """
        return table

    def _build_table_rows(self, format_issues, ner_data, ocr_data):
        rows = []
        idx = 1

        if format_issues:
            for issue in format_issues:
                location = self._extract_location(issue)
                description = self._extract_description(issue)
                current_val, required_val = self._extract_values(issue)
                rows.append(f"""
                <tr>
                    <td>{idx}</td>
                    <td>{location}</td>
                    <td>{description}</td>
                    <td class="current">{current_val}</td>
                    <td class="required">{required_val}</td>
                </tr>
                """)
                idx += 1

        if ner_data:
            if "entities" in ner_data:
                valid_types = {"ORG", "PER", "LOC", "PHONE", "TEL", "EMAIL", "ID_CARD", "CREDIT_CODE", "URL"}
                for ent in ner_data["entities"]:
                    if ent["type"] not in valid_types:
                        continue
                    loc = f"段落 {ent.get('position', '?')}" if ent["source_type"] == "word" else "图片区域"
                    rows.append(f"""
                    <tr>
                        <td>{idx}</td>
                        <td>{loc}</td>
                        <td>敏感实体检测：发现 {ent['type']}</td>
                        <td class="current">{ent['text']}</td>
                        <td class="required">不得出现</td>
                    </tr>
                    """)
                    idx += 1
            else:
                for key, label in [("ORG", "机构名"), ("PER", "人名"), ("LOC", "地名")]:
                    if ner_data.get(key):
                        for entity in ner_data[key][:5]:
                            rows.append(f"""
                            <tr>
                                <td>{idx}</td>
                                <td>全文</td>
                                <td>敏感实体检测：发现 {label}</td>
                                <td class="current">{entity}</td>
                                <td class="required">不得出现</td>
                            </tr>
                            """)
                            idx += 1

        if ocr_data:
            total_chars = sum(len(item.get("text", "")) for item in ocr_data)
            if total_chars > 50:
                rows.append(f"""
                <tr>
                    <td>{idx}</td>
                    <td>图片区域</td>
                    <td>图片中包含可识别文字</td>
                    <td class="current">{len(ocr_data)} 张图片 / {total_chars} 字</td>
                    <td class="required">建议核实</td>
                </tr>
                """)

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
        desc = re.sub(r"^(段落\s+\d+:\s*)", "", issue)
        desc = re.sub(r"^(第\s+\d+\s+节[：:]\s*)", "", desc)
        desc = re.sub(r"^(表格\s+\d+\s+)", "", desc)
        return desc

    def _extract_values(self, issue):
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
                    total_issues, rule_html, table_rows, passed):
        status_icon = "✅" if passed else "❌"
        status_text = "检查通过" if passed else "发现违规"
        status_color = "#27ae60" if passed else "#e74c3c"

        html_template = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>暗标合规检查报告 - {file_name}</title>
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
            max-width: 1000px;
            margin: 0 auto;
            background: #fff;
            border-radius: 12px;
            box-shadow: 0 2px 12px rgba(0,0,0,0.08);
            overflow: hidden;
        }}
        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }}
        .header h1 {{ font-size: 24px; margin-bottom: 8px; }}
        .header .meta {{ font-size: 14px; opacity: 0.9; }}
        .status-card {{
            padding: 25px 30px;
            text-align: center;
            border-bottom: 1px solid #eee;
        }}
        .status-badge {{
            display: inline-block;
            padding: 12px 30px;
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
        .stat-item {{ text-align: center; }}
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
            margin-bottom: 15px;
            display: flex;
            align-items: center;
            gap: 8px;
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
        .current {{ color: #e74c3c; }}
        .required {{ color: #27ae60; }}
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
            <h1>📋 暗标合规检查报告</h1>
            <div class="meta">{file_name} &nbsp;|&nbsp; 检查时间: {check_time}</div>
        </div>

        <div class="status-card">
            <div class="status-badge">{status_icon} {status_text}</div>
            <div class="stats">
                <div class="stat-item"><div class="stat-num">{total_issues}</div><div class="stat-label">总问题数</div></div>
                <div class="stat-item"><div class="stat-num">{format_count}</div><div class="stat-label">格式问题</div></div>
                <div class="stat-item"><div class="stat-num">{ner_count}</div><div class="stat-label">敏感信息</div></div>
                <div class="stat-item"><div class="stat-num">{ocr_count}</div><div class="stat-label">图片数量</div></div>
            </div>
        </div>

        {rule_html}

        <div class="section">
            <div class="section-title">🔍 问题明细</div>
            <div style="overflow-x: auto;">
                <table>
                    <thead><tr><th style="width:40px">#</th><th style="width:80px">位置</th><th>问题描述</th><th style="width:120px">当前值</th><th style="width:120px">要求值</th></tr></thead>
                    <tbody>{table_rows}</tbody>
                </table>
            </div>
        </div>

        <div class="footer">本报告由暗标合规检查系统自动生成，仅供参考使用</div>
    </div>
</body>
</html>"""
        return html_template
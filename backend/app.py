"""
backend/app.py - Flask后端主入口 (简洁版)
"""
import sys
import os
import shutil
import uuid
import time
import json
from datetime import datetime
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

# --- 项目路径配置 ---
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

# --- 导入业务逻辑模块 ---
from backend.config import UPLOAD_DIR, OUTPUT_DIR, MAX_FILE_SIZE, MAX_PAGES, VALID_CODES, PERMANENT_CODES
from core.format_checker import FormatChecker
from utils.annotator import DocumentAnnotator
from utils.report_gen import ReportGenerator

app = Flask(__name__)
CORS(app)

# --- 内存任务状态 ---
tasks = {}

# --- 激活码管理 ---
USED_CODES_FILE = '/tmp/used_codes.json'

def load_used_codes():
    if os.path.exists(USED_CODES_FILE):
        try:
            with open(USED_CODES_FILE, 'r') as f:
                return set(json.load(f))
        except Exception:
            return set()
    return set()

def save_used_codes(used_set):
    try:
        with open(USED_CODES_FILE, 'w') as f:
            json.dump(list(used_set), f)
    except Exception as e:
        print(f'[激活码] 保存失败: {e}')

used_codes = load_used_codes()
available_codes = set(VALID_CODES) - used_codes

# --- 工具函数 ---
def generate_task_id():
    return 'task_' + str(int(time.time())) + '_' + uuid.uuid4().hex[:6]

def get_task_dir(task_id):
    task_dir = os.path.join(UPLOAD_DIR, task_id)
    os.makedirs(task_dir, exist_ok=True)
    return task_dir

def get_output_dir(task_id):
    out_dir = os.path.join(OUTPUT_DIR, task_id)
    os.makedirs(out_dir, exist_ok=True)
    return out_dir

def validate_document(file_path):
    size = os.path.getsize(file_path)
    if size > MAX_FILE_SIZE:
        return False, f"文件大小{size/1024/1024:.2f}MB，超过限制1MB"
    try:
        from docx import Document
        doc = Document(file_path)
        para_count = len([p for p in doc.paragraphs if p.text.strip()])
        estimated_pages = para_count // 35 + 1
        if estimated_pages > MAX_PAGES:
            return False, f"估算页数约{estimated_pages}页，超过限制{MAX_PAGES}页"
    except Exception as e:
        return False, f"文件解析失败: {str(e)}"
    return True, "校验通过"

# ========== API接口 ==========

@app.route('/api/health', methods=['GET'])
def health_check():
    return jsonify({'status': 'ok', 'message': '服务运行中'})

@app.route('/api/verify', methods=['POST'])
def verify_code():
    global available_codes, used_codes
    data = request.get_json() or {}
    code = data.get('code', '').strip().upper()

    if not code:
        return jsonify({'success': False, 'message': '请输入激活码'}), 400

    if code in PERMANENT_CODES:
        return jsonify({'success': True, 'message': '验证成功（永久码）', 'permanent': True})

    if code not in available_codes:
        return jsonify({'success': False, 'message': '激活码无效或已被使用'}), 401

    available_codes.remove(code)
    used_codes.add(code)
    save_used_codes(used_codes)

    return jsonify({'success': True, 'message': '验证成功', 'permanent': False})

@app.route('/api/upload-req-text', methods=['POST'])
def upload_requirement_text():
    data = request.get_json() or {}
    text = data.get('text', '').strip()
    if not text:
        return jsonify({'success': False, 'message': '文本不能为空'}), 400
    if len(text) > 5000:
        return jsonify({'success': False, 'message': '文本超过5000字限制'}), 400

    task_id = generate_task_id()
    task_dir = get_task_dir(task_id)
    req_path = os.path.join(task_dir, 'requirement.txt')

    with open(req_path, 'w', encoding='utf-8') as f:
        f.write(text)

    tasks[task_id] = {
        'status': 'req_uploaded',
        'requirement_type': 'text',
        'requirement_path': req_path,
        'created_at': datetime.now().isoformat()
    }
    return jsonify({'success': True, 'task_id': task_id, 'message': '格式要求已接收'})

@app.route('/api/upload-req-file', methods=['POST'])
def upload_requirement_file():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': '没有上传文件'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'message': '文件名为空'}), 400

    task_id = generate_task_id()
    task_dir = get_task_dir(task_id)
    ext = os.path.splitext(file.filename)[1].lower()
    req_path = os.path.join(task_dir, f'requirement{ext}')
    file.save(req_path)

    tasks[task_id] = {
        'status': 'req_uploaded',
        'requirement_type': 'file',
        'requirement_path': req_path,
        'file_name': file.filename,
        'created_at': datetime.now().isoformat()
    }
    return jsonify({'success': True, 'task_id': task_id, 'message': '文件已接收'})

@app.route('/api/generate-rules', methods=['POST'])
def generate_rules():
    data = request.get_json() or {}
    task_id = data.get('taskId') or data.get('task_id')

    if not task_id or task_id not in tasks:
        return jsonify({'success': False, 'message': '任务不存在'}), 404

    task = tasks[task_id]
    req_path = task.get('requirement_path')

    if not req_path or not os.path.exists(req_path):
        return jsonify({'success': False, 'message': '格式要求文件不存在'}), 400

    # 模拟生成规则
    default_rules = {"document_info": {"name": "模拟规则", "generated_date": "2026-07-16"}}

    task_dir = get_task_dir(task_id)
    rules_path = os.path.join(task_dir, 'rules.json')
    with open(rules_path, 'w', encoding='utf-8') as f:
        json.dump(default_rules, f, ensure_ascii=False, indent=2)

    tasks[task_id]['status'] = 'rules_generated'
    tasks[task_id]['rules_path'] = rules_path

    return jsonify({'success': True, 'message': '配置已生成', 'task_id': task_id})

@app.route('/api/upload-doc', methods=['POST'])
def upload_document():
    task_id = request.form.get('taskId') or request.form.get('task_id')
    if not task_id or task_id not in tasks:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': '没有上传文件'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'message': '文件名为空'}), 400
    if not file.filename.lower().endswith('.docx'):
        return jsonify({'success': False, 'message': '仅支持 .docx 格式'}), 400

    file.seek(0, os.SEEK_END)
    file_size = file.tell()
    file.seek(0)

    if file_size > MAX_FILE_SIZE:
        return jsonify({'success': False, 'message': f'文件大小{file_size/1024/1024:.2f}MB，超过限制1MB'}), 400

    task_dir = get_task_dir(task_id)
    doc_path = os.path.join(task_dir, 'document.docx')
    file.save(doc_path)

    valid, msg = validate_document(doc_path)
    if not valid:
        os.remove(doc_path)
        return jsonify({'success': False, 'message': msg}), 400

    tasks[task_id]['status'] = 'doc_uploaded'
    tasks[task_id]['doc_path'] = doc_path
    tasks[task_id]['doc_name'] = file.filename

    return jsonify({'success': True, 'message': '文件上传成功', 'task_id': task_id})

@app.route('/api/start-check', methods=['POST'])
def start_check():
    data = request.get_json() or {}
    task_id = data.get('taskId') or data.get('task_id')

    if not task_id or task_id not in tasks:
        return jsonify({'success': False, 'message': '任务不存在'}), 404

    task = tasks[task_id]
    if task.get('status') != 'doc_uploaded':
        return jsonify({'success': False, 'message': '请先上传技术文件'}), 400

    doc_path = task.get('doc_path')
    rules_path = task.get('rules_path')

    if not os.path.exists(doc_path):
        return jsonify({'success': False, 'message': '技术文件不存在'}), 400

    if not rules_path or not os.path.exists(rules_path):
        rules_path = os.path.join(PROJECT_ROOT, 'clients', 'config', 'rules.json')
        if not os.path.exists(rules_path):
            return jsonify({'success': False, 'message': '找不到规则配置文件'}), 500

    tasks[task_id]['status'] = 'checking'

    try:
        # 1. 执行检查
        checker = FormatChecker(rules_path)
        format_issues = checker.check_document(doc_path)

        # 2. 生成批注版
        annotator = DocumentAnnotator()
        out_dir = get_output_dir(task_id)
        annotated_path = annotator.generate_annotated_copy(
            original_path=doc_path,
            format_issues=format_issues,
            output_dir=out_dir,
            suffix='_批注版'
        )

        # 3. 生成报告
        report_gen = ReportGenerator()
        report_path = report_gen.generate_html_report(
            file_path=doc_path,
            format_issues=format_issues,
            output_dir=out_dir,
            suffix='_检查报告'
        )

        # 4. 更新状态
        tasks[task_id]['status'] = 'completed'
        tasks[task_id]['format_issues'] = format_issues
        tasks[task_id]['issue_count'] = len(format_issues)
        tasks[task_id]['annotated_path'] = annotated_path
        tasks[task_id]['report_path'] = report_path
        tasks[task_id]['completed_at'] = datetime.now().isoformat()

        return jsonify({
            'success': True,
            'message': '检查完成',
            'task_id': task_id,
            'issue_count': len(format_issues),
            'passed': len(format_issues) == 0
        })

    except Exception as e:
        import traceback
        tasks[task_id]['status'] = 'failed'
        tasks[task_id]['error'] = str(e)
        tasks[task_id]['traceback'] = traceback.format_exc()
        return jsonify({'success': False, 'message': f'检查失败: {str(e)}'}), 500

@app.route('/api/status/<task_id>', methods=['GET'])
def get_status(task_id):
    if not task_id or task_id not in tasks:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    task = tasks[task_id]
    return jsonify({
        'success': True,
        'task_id': task_id,
        'status': task.get('status'),
        'issue_count': task.get('issue_count', 0)
    })

@app.route('/api/download/<task_id>/<file_type>', methods=['GET'])
def download_result(task_id, file_type):
    if not task_id or task_id not in tasks:
        return jsonify({'success': False, 'message': '任务不存在'}), 404

    task = tasks[task_id]
    if task.get('status') != 'completed':
        return jsonify({'success': False, 'message': '检查尚未完成'}), 400

    file_path = None
    file_name = ""

    if file_type == 'report':
        file_path = task.get('report_path')
        file_name = "格式检查报告.html"
    elif file_type == 'annotated':
        file_path = task.get('annotated_path')
        file_name = "批注版.docx"
    else:
        return jsonify({'success': False, 'message': '无效的文件类型'}), 400

    if not file_path or not os.path.exists(file_path):
        return jsonify({'success': False, 'message': '文件已丢失或不存在'}), 404

    try:
        return send_file(
            file_path,
            as_attachment=True,
            download_name=file_name,
        )
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

if __name__ == '__main__':
    print("=" * 50)
    print(" 追标猎手 - 暗标格式检查后端服务")
    print("=" * 50)
    print(f" 项目根目录: {PROJECT_ROOT}")
    print(f" 上传目录: {UPLOAD_DIR}")
    print(f" 输出目录: {OUTPUT_DIR}")
    print(f" 存储模式: 本地临时文件")
    print("=" * 50)
    app.run(host='0.0.0.0', port=5000, debug=False)

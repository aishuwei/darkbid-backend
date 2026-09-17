import sys
import os
import uuid
import time
import json
import requests
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

# ============================================================
# 云开发配置
# ============================================================
ENV_ID = 'darkbid-d8gxpsbued4a2867'  # ⚠️ 你的云开发环境ID
API_KEY = os.environ.get('CLOUDBASE_APIKEY')  # 从云托管环境变量读取

# ✅ 修正1: 使用正确的云开发数据库 API 域名
BASE_URL = f'https://{ENV_ID}.api.cloudbase.net'

# 通用请求头
def get_headers():
    return {
        'Authorization': f'Bearer {API_KEY}',
        'Content-Type': 'application/json'
    }

app = Flask(__name__)
CORS(app)

# ============================================================
# 【核心】云开发数据库操作 (修正版)
# ============================================================

def db_request(method, path, data=None):
    """调用云开发 HTTP API"""
    url = f'{BASE_URL}{path}'
    headers = get_headers()
    try:
        if method == 'GET':
            resp = requests.get(url, headers=headers, timeout=10)
        elif method == 'POST':
            resp = requests.post(url, headers=headers, json=data, timeout=10)
        elif method == 'PATCH':
            resp = requests.patch(url, headers=headers, json=data, timeout=10)
        elif method == 'DELETE':
            resp = requests.delete(url, headers=headers, timeout=10)
        else:
            raise ValueError(f'Unsupported method: {method}')
        
        # 打印响应以便调试
        # print(f"[数据库] {method} {url} -> {resp.status_code}: {resp.text}")
        
        return resp.json()
    except Exception as e:
        print(f'[数据库] 请求失败: {e}')
        return None

def create_task(task_id, task_data):
    """创建任务到数据库"""
    task_data['task_id'] = task_id
    task_data['created_at'] = datetime.now().isoformat()
    # ✅ 修正2: 使用正确的 API 路径
    result = db_request('POST', '/v1/database/instances/(default)/databases/(default)/collections/tasks/documents', {'data': task_data})
    return result is not None and 'id' in result

def get_task(task_id):
    """从数据库获取任务"""
    # ✅ 修正3: 使用正确的 where 查询参数
    path = f'/v1/database/instances/(default)/databases/(default)/collections/tasks/documents?where={{"task_id":"{task_id}"}}'
    result = db_request('GET', path)
    if result and 'data' in result and len(result['data']) > 0:
        # 将 _id 字段重命名为 id，以兼容后续逻辑
        doc = result['data'][0]
        doc['id'] = doc.get('_id')
        return doc
    return None

def update_task(task_id, update_data):
    """更新任务到数据库"""
    update_data['updated_at'] = datetime.now().isoformat()
    # ✅ 修正4: 使用正确的 API 路径和 where 查询参数
    path = f'/v1/database/instances/(default)/databases/(default)/collections/tasks/documents?where={{"task_id":"{task_id}"}}'
    result = db_request('PATCH', path, {'data': update_data})
    return result is not None and result.get('updated', 0) > 0

def is_code_used(code):
    """检查激活码是否已被使用"""
    path = f'/v1/database/instances/(default)/databases/(default)/collections/used_codes/documents?where={{"code":"{code}"}}'
    result = db_request('GET', path)
    return result and 'data' in result and len(result['data']) > 0

def mark_code_used(code):
    """标记激活码已使用"""
    data = {
        'code': code,
        'used_at': datetime.now().isoformat()
    }
    result = db_request('POST', '/v1/database/instances/(default)/databases/(default)/collections/used_codes/documents', {'data': data})
    return result is not None and 'id' in result

# ============================================================
# 【核心】云存储操作
# ============================================================

def upload_to_storage(local_path, cloud_path):
    """上传文件到云存储"""
    try:
        # 1. 获取上传信息
        url = f'{BASE_URL}/v1/storages/get-objects-upload-info'
        headers = get_headers()
        data = [{'objectId': cloud_path}]
        resp = requests.post(url, headers=headers, json=data, timeout=10)
        upload_info = resp.json()
        if not upload_info or not isinstance(upload_info, list):
            print(f'[云存储] 获取上传信息失败: {upload_info}')
            return None
        info = upload_info[0]
        upload_url = info.get('uploadUrl')

        # 2. 上传文件
        with open(local_path, 'rb') as f:
            file_content = f.read()
        upload_headers = {
            'Authorization': info.get('authorization', ''),
            'X-Cos-Security-Token': info.get('token', ''),
            'X-Cos-Meta-Fileid': info.get('cloudObjectMeta', '')
        }
        resp = requests.put(upload_url, data=file_content, headers=upload_headers, timeout=30)
        if resp.status_code in [200, 201]:
            file_id = f'cloud://{ENV_ID}/{cloud_path}'
            print(f'[云存储] 上传成功: {file_id}')
            return file_id
        else:
            print(f'[云存储] 上传失败: {resp.status_code}')
            return None
    except Exception as e:
        print(f'[云存储] 上传异常: {e}')
        return None

def get_download_url(file_id):
    """获取文件下载链接"""
    try:
        url = f'{BASE_URL}/v1/storages/get-objects-download-info'
        headers = get_headers()
        data = [{'cloudObjectId': file_id}]
        resp = requests.post(url, headers=headers, json=data, timeout=10)
        result = resp.json()
        if result and isinstance(result, list) and len(result) > 0:
            return result[0].get('downloadUrl')
        return None
    except Exception as e:
        print(f'[云存储] 获取下载链接失败: {e}')
        return None

# ============================================================
# 工具函数
# ============================================================

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
        return False, f"文件大小{size/1024/1024:.2f}MB，超过限制{MAX_FILE_SIZE/1024/1024:.0f}MB"
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
    return jsonify({'status': 'ok', 'message': '服务运行中', 'env': ENV_ID})

@app.route('/api/verify', methods=['POST'])
def verify_code():
    data = request.get_json() or {}
    code = data.get('code', '').strip().upper()
    if not code:
        return jsonify({'success': False, 'message': '请输入激活码'}), 400
    # 永久码
    if code in PERMANENT_CODES:
        return jsonify({'success': True, 'message': '验证成功（永久码）', 'permanent': True})
    # 检查是否是有效码
    if code not in VALID_CODES:
        return jsonify({'success': False, 'message': '激活码无效'}), 401
    # 检查是否已被使用
    if is_code_used(code):
        return jsonify({'success': False, 'message': '激活码已被使用'}), 401
    # 标记为已使用
    mark_code_used(code)
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
    # 创建任务到数据库
    create_task(task_id, {
        'status': 'req_uploaded',
        'requirement_type': 'text',
        'requirement_path': req_path
    })
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
    # 创建任务到数据库
    create_task(task_id, {
        'status': 'req_uploaded',
        'requirement_type': 'file',
        'requirement_path': req_path,
        'file_name': file.filename
    })
    return jsonify({'success': True, 'task_id': task_id, 'message': '文件已接收'})

@app.route('/api/generate-rules', methods=['POST'])
def generate_rules():
    data = request.get_json() or {}
    task_id = data.get('taskId') or data.get('task_id')
    if not task_id:
        return jsonify({'success': False, 'message': '缺少任务ID'}), 400
    # 从数据库获取任务
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    req_path = task.get('requirement_path')
    if not req_path or not os.path.exists(req_path):
        return jsonify({'success': False, 'message': '格式要求文件不存在'}), 400
    # 生成规则
    default_rules = {"document_info": {"name": "模拟规则", "generated_date": "2026-07-16"}}
    task_dir = get_task_dir(task_id)
    rules_path = os.path.join(task_dir, 'rules.json')
    with open(rules_path, 'w', encoding='utf-8') as f:
        json.dump(default_rules, f, ensure_ascii=False, indent=2)
    # 更新任务状态
    update_task(task_id, {
        'status': 'rules_generated',
        'rules_path': rules_path
    })
    return jsonify({'success': True, 'message': '配置已生成', 'task_id': task_id})

@app.route('/api/upload-doc', methods=['POST'])
def upload_document():
    task_id = request.form.get('taskId') or request.form.get('task_id')
    if not task_id:
        return jsonify({'success': False, 'message': '缺少任务ID'}), 400
    # 从数据库获取任务
    task = get_task(task_id)
    if not task:
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
        return jsonify({'success': False, 'message': f'文件大小{file_size/1024/1024:.2f}MB，超过限制{MAX_FILE_SIZE/1024/1024:.0f}MB'}), 400
    task_dir = get_task_dir(task_id)
    doc_path = os.path.join(task_dir, 'document.docx')
    file.save(doc_path)
    valid, msg = validate_document(doc_path)
    if not valid:
        os.remove(doc_path)
        return jsonify({'success': False, 'message': msg}), 400
    # 更新任务状态到数据库
    update_task(task_id, {
        'status': 'doc_uploaded',
        'doc_path': doc_path,
        'doc_name': file.filename
    })
    return jsonify({'success': True, 'message': '文件上传成功', 'task_id': task_id})

@app.route('/api/start-check', methods=['POST'])
def start_check():
    data = request.get_json() or {}
    task_id = data.get('taskId') or data.get('task_id')
    if not task_id:
        return jsonify({'success': False, 'message': '缺少任务ID'}), 400
    # 从数据库获取任务
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    if task.get('status') != 'doc_uploaded':
        return jsonify({'success': False, 'message': '请先上传技术文件'}), 400
    doc_path = task.get('doc_path')
    rules_path = task.get('rules_path')
    if not doc_path or not os.path.exists(doc_path):
        return jsonify({'success': False, 'message': '技术文件不存在'}), 400
    if not rules_path or not os.path.exists(rules_path):
        rules_path = os.path.join(PROJECT_ROOT, 'clients', 'config', 'rules.json')
        if not os.path.exists(rules_path):
            return jsonify({'success': False, 'message': '找不到规则配置文件'}), 500
    update_task(task_id, {'status': 'checking'})
    try:
        checker = FormatChecker(rules_path)
        format_issues = checker.check_document(doc_path)
        annotator = DocumentAnnotator()
        out_dir = get_output_dir(task_id)
        annotated_path = annotator.generate_annotated_copy(
            original_path=doc_path,
            format_issues=format_issues,
            output_dir=out_dir,
            suffix='_批注版'
        )
        report_gen = ReportGenerator()
        report_path = report_gen.generate_html_report(
            file_path=doc_path,
            format_issues=format_issues,
            output_dir=out_dir,
            suffix='_检查报告'
        )
        # 上传结果文件到云存储
        report_file_id = upload_to_storage(report_path, f'{task_id}/report.html')
        annotated_file_id = upload_to_storage(annotated_path, f'{task_id}/annotated.docx')
        # 更新任务状态到数据库
        update_task(task_id, {
            'status': 'completed',
            'issue_count': len(format_issues),
            'report_file_id': report_file_id,
            'annotated_file_id': annotated_file_id,
            'completed_at': datetime.now().isoformat()
        })
        return jsonify({
            'success': True,
            'message': '检查完成',
            'task_id': task_id,
            'issue_count': len(format_issues),
            'passed': len(format_issues) == 0
        })
    except Exception as e:
        import traceback
        update_task(task_id, {
            'status': 'failed',
            'error': str(e),
            'traceback': traceback.format_exc()
        })
        return jsonify({'success': False, 'message': f'检查失败: {str(e)}'}), 500

@app.route('/api/status/<task_id>', methods=['GET'])
def get_status(task_id):
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    return jsonify({
        'success': True,
        'task_id': task_id,
        'status': task.get('status'),
        'issue_count': task.get('issue_count', 0)
    })

@app.route('/api/download/<task_id>/<file_type>', methods=['GET'])
def download_result(task_id, file_type):
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    if task.get('status') != 'completed':
        return jsonify({'success': False, 'message': '检查尚未完成'}), 400
    file_id = None
    file_name = ""
    if file_type == 'report':
        file_id = task.get('report_file_id')
        file_name = "格式检查报告.html"
    elif file_type == 'annotated':
        file_id = task.get('annotated_file_id')
        file_name = "批注版.docx"
    else:
        return jsonify({'success': False, 'message': '无效的文件类型'}), 400
    if not file_id:
        return jsonify({'success': False, 'message': '文件不存在'}), 404
    # 获取下载链接
    download_url = get_download_url(file_id)
    if not download_url:
        return jsonify({'success': False, 'message': '获取下载链接失败'}), 500
    return jsonify({
        'success': True,
        'download_url': download_url,
        'file_name': file_name
    })

if __name__ == '__main__':
    print("=" * 50)
    print(" 追标猎手 - 暗标格式检查后端服务")
    print("=" * 50)
    print(f" 项目根目录: {PROJECT_ROOT}")
    print(f" 上传目录: {UPLOAD_DIR}")
    print(f" 输出目录: {OUTPUT_DIR}")
    print(f" 文件大小限制: {MAX_FILE_SIZE/1024/1024:.0f}MB")
    print(f" 云开发环境: {ENV_ID}")
    print("=" * 50)
    app.run(host='0.0.0.0', port=5000, debug=False)

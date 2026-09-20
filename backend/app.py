# -*- coding: utf-8 -*-
"""
backend/app.py — 追标猎手 · 暗标格式检查后端（云函数版）
数据库操作全部通过云函数 dbOperations 中转，云函数内天然有管理员权限，无需 token。
"""
import os
import sys
import json
import time
import uuid
import shutil
import tempfile
import logging
import requests
from datetime import datetime
from flask import Flask, request, jsonify

sys.stdout.reconfigure(line_buffering=True)
sys.stderr.reconfigure(line_buffering=True)
logging.basicConfig(level=logging.INFO, stream=sys.stdout, force=True,
                    format='%(message)s')

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from backend.config import MAX_FILE_SIZE, MAX_PAGES, VALID_CODES, PERMANENT_CODES
from core.format_checker import FormatChecker
from utils.annotator import DocumentAnnotator
from utils.report_gen import ReportGenerator

# ============================================================
# 云开发配置
# ============================================================
ENV_ID = os.environ.get('TCB_ENV', 'darkbid-d8gxpsbued4a2867')
CLOUD_FUNCTION_NAME = 'dbOperations'

logging.info(f'[boot] ENV_ID={ENV_ID}')
logging.info(f'[boot] CLOUD_FUNCTION_NAME={CLOUD_FUNCTION_NAME}')

app = Flask(__name__)
try:
    from flask_cors import CORS
    CORS(app)
except ImportError:
    pass


# ============================================================
# 云函数调用（容器内免鉴权）
# ============================================================
def call_cloud_function(action, collection, query=None, data=None):
    """通过 /tcb/invokecloudfunction 调用云函数 dbOperations"""
    token = request.headers.get('X-WX-CLOUDBASE-ACCESS-TOKEN', '')
    if not token:
        logging.error('[cloudFunc] 未拿到 X-WX-CLOUDBASE-ACCESS-TOKEN，'
                      '请确认请求从微信链路进入（小程序 callContainer），'
                      '且云托管控制台已开启「云调用」')
        return None

    params = {
        'action': action,
        'collection': collection,
        'query': query or {},
        'data': data or {}
    }
    url = 'https://api.weixin.qq.com/tcb/invokecloudfunction'
    payload = {
        'env': ENV_ID,
        'name': CLOUD_FUNCTION_NAME,
        'req': json.dumps(params, ensure_ascii=False)
    }
    try:
        resp = requests.post(url,
                             params={'cloudbase_access_token': token},
                             json=payload, timeout=15)
        result = resp.json()
        logging.info(f'[cloudFunc] {action} {collection} 返回: {str(result)[:300]}')
        if result.get('errcode', 0) != 0:
            logging.error(f'[cloudFunc] 调用失败: {result}')
            return None
        resp_data = result.get('resp_data', '{}')
        fn_result = json.loads(resp_data)
        if not fn_result.get('success'):
            logging.error(f'[cloudFunc] 云函数内部失败: {fn_result}')
            return None
        return fn_result.get('result')
    except Exception:
        import traceback
        logging.error(f'[cloudFunc] 异常: {traceback.format_exc()}')
        return None


# ============================================================
# 数据库操作（全部走云函数）
# ============================================================
def db_add(collection, doc):
    return call_cloud_function('add', collection, data=doc)


def db_query_one(collection, field, value):
    return call_cloud_function('queryOne', collection, query={field: value})


def db_update_where(collection, field, value, data):
    result = call_cloud_function('update', collection,
                                 query={field: value}, data=data)
    return result is not None


# ============================================================
# 云存储（暂时占位，后续接入）
# ============================================================
def storage_upload(local_path, cloud_path):
    logging.error('[storage] 云存储待接入')
    return None


def storage_download(file_id, local_path):
    logging.error('[storage] 云存储待接入')
    return False


# ============================================================
# 任务 / 激活码
# ============================================================
def generate_task_id():
    return 'task_' + str(int(time.time())) + '_' + uuid.uuid4().hex[:8]


def create_task(task_id, data):
    doc = dict(data)
    doc['task_id'] = task_id
    doc['created_at'] = datetime.now().isoformat()
    return db_add('tasks', doc) is not None


def get_task(task_id):
    return db_query_one('tasks', 'task_id', task_id)


def update_task(task_id, data):
    data['updated_at'] = datetime.now().isoformat()
    return db_update_where('tasks', 'task_id', task_id, data)


def is_code_used(code):
    return db_query_one('used_codes', 'code', code) is not None


def mark_code_used(code):
    return db_add('used_codes', {'code': code,
                                 'used_at': datetime.now().isoformat()}) is not None


# ============================================================
# 工具
# ============================================================
def validate_document(file_path):
    size = os.path.getsize(file_path)
    if size > MAX_FILE_SIZE:
        return False, f"文件大小{size / 1024 / 1024:.2f}MB，超过限制{MAX_FILE_SIZE / 1024 / 1024:.0f}MB"
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


def extract_text(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == '.docx':
        from docx import Document
        return '\n'.join(p.text for p in Document(path).paragraphs)
    for enc in ('utf-8', 'gbk'):
        try:
            with open(path, 'r', encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
        return f.read()


# ============================================================
# API 接口
# ============================================================
@app.route('/api/health', methods=['GET'])
def health_check():
    return jsonify({'status': 'ok', 'message': '服务运行中',
                    'env': ENV_ID, 'mode': 'cloud-function'})


@app.route('/api/db-check', methods=['GET'])
def db_check():
    result = call_cloud_function('query', 'tasks', query={})
    return jsonify({
        'success': result is not None,
        'cloud_function': CLOUD_FUNCTION_NAME,
        'sample': result[:1] if result else None
    })


@app.route('/api/verify', methods=['POST'])
def verify_code():
    data = request.get_json(silent=True) or {}
    code = (data.get('code') or '').strip().upper()
    if not code:
        return jsonify({'success': False, 'message': '请输入激活码'}), 400
    if code in PERMANENT_CODES:
        return jsonify({'success': True, 'message': '验证成功（永久码）', 'permanent': True})
    if code not in VALID_CODES:
        return jsonify({'success': False, 'message': '激活码无效'}), 401
    if is_code_used(code):
        return jsonify({'success': False, 'message': '激活码已被使用'}), 401
    if not mark_code_used(code):
        return jsonify({'success': False, 'message': '激活码核销失败，请重试'}), 500
    return jsonify({'success': True, 'message': '验证成功', 'permanent': False})


@app.route('/api/upload-req-text', methods=['POST'])
def upload_requirement_text():
    data = request.get_json(silent=True) or {}
    text = (data.get('text') or '').strip()
    if not text:
        return jsonify({'success': False, 'message': '文本不能为空'}), 400
    if len(text) > 5000:
        return jsonify({'success': False, 'message': '文本超过5000字限制'}), 400
    task_id = generate_task_id()
    if not create_task(task_id, {'status': 'req_uploaded',
                                 'requirement_type': 'text',
                                 'requirement_text': text}):
        logging.error(f'[task] 任务创建失败: {task_id}')
        return jsonify({'success': False, 'message': '任务创建失败'}), 500
    logging.info(f'[task] 文本任务创建成功: {task_id}')
    return jsonify({'success': True, 'task_id': task_id, 'message': '格式要求已接收'})


@app.route('/api/upload-req-file', methods=['POST'])
def upload_requirement_file():
    return jsonify({'success': False, 'message': '云存储待接入'}), 501


@app.route('/api/generate-rules', methods=['POST'])
def generate_rules():
    data = request.get_json(silent=True) or {}
    task_id = data.get('taskId') or data.get('task_id')
    if not task_id:
        return jsonify({'success': False, 'message': '缺少任务ID'}), 400
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    if task.get('status') not in ('req_uploaded', 'rules_generated'):
        return jsonify({'success': False, 'message': '任务状态异常'}), 400
    req_content = task.get('requirement_text') or ''
    rules = {
        'document_info': {
            'name': '模拟规则',
            'generated_date': datetime.now().strftime('%Y-%m-%d'),
        },
        'source_text_length': len(req_content),
    }
    rules_json = json.dumps(rules, ensure_ascii=False, indent=2)
    if not update_task(task_id, {'status': 'rules_generated', 'rules_json': rules_json}):
        return jsonify({'success': False, 'message': '规则保存失败'}), 500
    logging.info(f'[task] 规则生成成功: {task_id}')
    return jsonify({'success': True, 'message': '配置已生成', 'task_id': task_id})


@app.route('/api/upload-doc', methods=['POST'])
def upload_document():
    return jsonify({'success': False, 'message': '云存储待接入'}), 501


@app.route('/api/start-check', methods=['POST'])
def start_check():
    return jsonify({'success': False, 'message': '云存储待接入'}), 501


@app.route('/api/status/<task_id>', methods=['GET'])
def get_status(task_id):
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    return jsonify({'success': True, 'task_id': task_id,
                    'status': task.get('status'),
                    'issue_count': task.get('issue_count', 0)})


@app.route('/api/download/<task_id>/<file_type>', methods=['GET'])
def download_result(task_id, file_type):
    return jsonify({'success': False, 'message': '云存储待接入'}), 501


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 80))
    logging.info('=' * 50)
    logging.info(' 追标猎手 - 后端服务（云函数版）')
    logging.info('=' * 50)
    logging.info(f' 云开发环境: {ENV_ID}')
    logging.info(f' 云函数: {CLOUD_FUNCTION_NAME}')
    logging.info(f' 监听端口: {port}')
    logging.info('=' * 50)
    app.run(host='0.0.0.0', port=port, debug=False)

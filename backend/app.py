# -*- coding: utf-8 -*-
"""
backend/app.py — 追标猎手 · 暗标格式检查后端（云托管直连版）
架构：CloudRun (Container / Python) → NoSQL HTTP API + COS Storage
"""
import os
import sys
import json
import time
import uuid
import tempfile
import shutil
import logging
from datetime import datetime
from flask import Flask, request, jsonify

sys.stdout.reconfigure(line_buffering=True)
sys.stderr.reconfigure(line_buffering=True)
logging.basicConfig(level=logging.INFO, stream=sys.stdout, force=True,
                    format='%(asctime)s [%(levelname)s] %(message)s')

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from backend.config import MAX_FILE_SIZE, MAX_PAGES, VALID_CODES, PERMANENT_CODES
from backend.nosql_client import get_client, NoSQLClient
from backend.storage_client import get_storage_client

from core.format_checker import FormatChecker
from utils.annotator import DocumentAnnotator
from utils.report_gen import ReportGenerator

# ============================================================
# 全局配置
# ============================================================
ENV_ID = os.environ.get('TCB_ENV', 'darkbid-d8gxpsbued4a2867')
logging.info(f'[boot] ENV_ID={ENV_ID}')

app = Flask(__name__)
try:
    from flask_cors import CORS
    CORS(app)
except ImportError:
    pass


# ============================================================
# 数据库客户端封装
# ============================================================
def get_db() -> NoSQLClient:
    return get_client()


def db_add(collection, doc):
    return get_db().add(collection, doc)


def db_query_one(collection, field, value):
    return get_db().query_one(collection, field, value)


def db_query(collection, where=None, limit=100, offset=0, order=None):
    return get_db().query(collection, where=where, limit=limit,
                          offset=offset, order=order)


def db_update_where(collection, field, value, data):
    matched = get_db().update(collection, field, value, data)
    return matched is not None and matched > 0


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


def _make_tmp_dir(prefix):
    return tempfile.mkdtemp(prefix=prefix)


def _cleanup_tmp(d):
    try:
        shutil.rmtree(d, ignore_errors=True)
    except Exception:
        pass


# ============================================================
# 规则解析
# ============================================================
DEFAULT_RULES_PATH = os.path.join(PROJECT_ROOT, 'clients', 'config', 'rules.json')


def build_rules_from_requirement(req_content: str):
    base = {}
    if os.path.exists(DEFAULT_RULES_PATH):
        try:
            with open(DEFAULT_RULES_PATH, 'r', encoding='utf-8') as f:
                base = json.load(f)
        except Exception as e:
            logging.warning(f'[rules] 读取默认规则失败: {e}')

    base.setdefault('document_info', {})
    base['document_info']['name'] = '格式检查规则'
    base['document_info']['generated_date'] = datetime.now().strftime('%Y-%m-%d')
    base['source_text_length'] = len(req_content or '')
    return base


# ============================================================
# API: 健康 / 自检
# ============================================================
@app.route('/api/health', methods=['GET'])
def health_check():
    return jsonify({'status': 'ok', 'message': '服务运行中',
                    'env': ENV_ID, 'mode': 'cloudrun-direct'})


@app.route('/api/db-check', methods=['GET'])
def db_check():
    try:
        docs = db_query('tasks', limit=1)
        return jsonify({'success': docs is not None,
                        'mode': 'nosql-http-api',
                        'sample': docs[:1] if docs else []})
    except Exception as e:
        logging.exception('[db-check] 失败')
        return jsonify({'success': False, 'error': str(e)}), 500


# ============================================================
# API: 激活码
# ============================================================
@app.route('/api/verify', methods=['POST'])
def verify_code():
    data = request.get_json(silent=True) or {}
    code = (data.get('code') or '').strip().upper()
    if not code:
        return jsonify({'success': False, 'message': '请输入激活码'}), 400
    if code in PERMANENT_CODES:
        return jsonify({'success': True, 'message': '验证成功（永久码）',
                        'permanent': True})
    if code not in VALID_CODES:
        return jsonify({'success': False, 'message': '激活码无效'}), 401
    if is_code_used(code):
        return jsonify({'success': False, 'message': '激活码已被使用'}), 401
    if not mark_code_used(code):
        return jsonify({'success': False, 'message': '激活码核销失败'}), 500
    return jsonify({'success': True, 'message': '验证成功', 'permanent': False})


# ============================================================
# API: 上传格式要求（文本）
# ============================================================
@app.route('/api/upload-req-text', methods=['POST'])
def upload_requirement_text():
    data = request.get_json(silent=True) or {}
    text = (data.get('text') or '').strip()
    if not text:
        return jsonify({'success': False, 'message': '文本不能为空'}), 400
    if len(text) > 5000:
        return jsonify({'success': False, 'message': '文本超过5000字限制'}), 400
    task_id = generate_task_id()
    ok = create_task(task_id, {
        'status': 'req_uploaded',
        'requirement_type': 'text',
        'requirement_text': text,
    })
    if not ok:
        return jsonify({'success': False, 'message': '任务创建失败'}), 500
    logging.info(f'[task] 文本任务创建成功: {task_id}')
    return jsonify({'success': True, 'task_id': task_id,
                    'message': '格式要求已接收'})


# ============================================================
# API: 上传格式要求（文件）
# ============================================================
@app.route('/api/upload-req-file', methods=['POST'])
def upload_requirement_file():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': '未选择文件'}), 400
    file = request.files['file']
    if not file.filename:
        return jsonify({'success': False, 'message': '文件名为空'}), 400
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in ('.docx', '.txt'):
        return jsonify({'success': False, 'message': '仅支持 .docx 和 .txt'}), 400

    tmp_dir = _make_tmp_dir('req_')
    tmp_path = os.path.join(tmp_dir, file.filename)
    try:
        file.save(tmp_path)
        if ext == '.docx':
            valid, msg = validate_document(tmp_path)
            if not valid:
                return jsonify({'success': False, 'message': msg}), 400
        content = extract_text(tmp_path)
        if len(content) > 5000:
            return jsonify({'success': False, 'message': '要求文本超过5000字限制'}), 400

        storage = get_storage_client()
        cloud_path = f"requirements/{uuid.uuid4().hex}/{file.filename}"
        fileid = storage.upload(tmp_path, cloud_path)
        if not fileid:
            return jsonify({'success': False, 'message': '文件上传失败'}), 500

        task_id = generate_task_id()
        ok = create_task(task_id, {
            'status': 'req_uploaded',
            'requirement_type': 'file',
            'requirement_text': content,
            'file_name': file.filename,
            'req_file_id': fileid,
        })
        if not ok:
            return jsonify({'success': False, 'message': '任务创建失败'}), 500
        logging.info(f'[task] 文件任务创建成功: {task_id}')
        return jsonify({'success': True, 'task_id': task_id,
                        'message': '格式要求文件已上传',
                        'file_name': file.filename})
    finally:
        _cleanup_tmp(tmp_dir)


# ============================================================
# API: 生成规则
# ============================================================
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
    rules = build_rules_from_requirement(req_content)
    rules_json = json.dumps(rules, ensure_ascii=False, indent=2)

    if not update_task(task_id, {'status': 'rules_generated',
                                 'rules_json': rules_json}):
        return jsonify({'success': False, 'message': '规则保存失败'}), 500
    logging.info(f'[task] 规则生成成功: {task_id}')
    return jsonify({'success': True, 'message': '配置已生成',
                    'task_id': task_id})


# ============================================================
# API: 上传技术文档
# ============================================================
@app.route('/api/upload-doc', methods=['POST'])
def upload_document():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': '未选择文件'}), 400
    file = request.files['file']
    if not file.filename:
        return jsonify({'success': False, 'message': '文件名为空'}), 400
    if os.path.splitext(file.filename)[1].lower() != '.docx':
        return jsonify({'success': False, 'message': '仅支持 .docx 格式'}), 400

    task_id = request.form.get('task_id') or request.form.get('taskId')
    if not task_id:
        return jsonify({'success': False, 'message': '缺少任务ID'}), 400
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404

    tmp_dir = _make_tmp_dir('doc_')
    tmp_path = os.path.join(tmp_dir, file.filename)
    try:
        file.save(tmp_path)
        size = os.path.getsize(tmp_path)
        if size == 0:
            return jsonify({'success': False, 'message': '文件为空'}), 400
        if size > MAX_FILE_SIZE:
            return jsonify({'success': False,
                            'message': f'文件大小{size / 1024 / 1024:.2f}MB，'
                                       f'超过限制{MAX_FILE_SIZE / 1024 / 1024:.0f}MB'}), 400
        valid, msg = validate_document(tmp_path)
        if not valid:
            return jsonify({'success': False, 'message': msg}), 400

        storage = get_storage_client()
        cloud_path = f"documents/{task_id}/{file.filename}"
        doc_file_id = storage.upload(tmp_path, cloud_path)
        if not doc_file_id:
            return jsonify({'success': False, 'message': '文件上传失败'}), 500

        update_task(task_id, {
            'status': 'doc_uploaded',
            'doc_file_name': file.filename,
            'doc_file_id': doc_file_id,
        })
        logging.info(f'[task] 技术文档上传成功: {task_id}')
        return jsonify({'success': True, 'task_id': task_id,
                        'message': '投标文档已上传',
                        'file_name': file.filename})
    finally:
        _cleanup_tmp(tmp_dir)


# ============================================================
# API: 开始检查
# ============================================================
@app.route('/api/start-check', methods=['POST'])
def start_check():
    data = request.get_json(silent=True) or {}
    task_id = data.get('taskId') or data.get('task_id')
    if not task_id:
        return jsonify({'success': False, 'message': '缺少任务ID'}), 400
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    if task.get('status') != 'doc_uploaded':
        return jsonify({'success': False,
                        'message': f'任务状态异常（当前 {task.get("status")}，需 doc_uploaded）'}), 400

    doc_file_id = task.get('doc_file_id')
    if not doc_file_id:
        return jsonify({'success': False, 'message': '技术文件 fileid 缺失'}), 400

    update_task(task_id, {'status': 'checking'})
    tmp_dir = _make_tmp_dir('chk_')
    try:
        storage = get_storage_client()

        # 1. 下载文档
        doc_name = task.get('doc_file_name') or 'document.docx'
        doc_local = os.path.join(tmp_dir, doc_name)
        download_url = storage.get_download_url(doc_file_id, expires=600)
        if not download_url:
            update_task(task_id, {'status': 'doc_uploaded'})
            return jsonify({'success': False, 'message': '获取文档下载链接失败'}), 500
        import requests as _rq
        r = _rq.get(download_url, timeout=120)
        if r.status_code != 200:
            update_task(task_id, {'status': 'doc_uploaded'})
            return jsonify({'success': False, 'message': '文档下载失败'}), 500
        with open(doc_local, 'wb') as f:
            f.write(r.content)

        # 2. 准备规则文件
        rules_path = os.path.join(tmp_dir, 'rules.json')
        rules_json = task.get('rules_json')
        if rules_json:
            with open(rules_path, 'w', encoding='utf-8') as f:
                f.write(rules_json)
        else:
            if not os.path.exists(DEFAULT_RULES_PATH):
                update_task(task_id, {'status': 'doc_uploaded'})
                return jsonify({'success': False, 'message': '找不到规则文件'}), 500
            rules_path = DEFAULT_RULES_PATH

        # 3. 执行检查
        logging.info(f'[check] 开始检查: {task_id}')
        checker = FormatChecker(rules_path)
        format_issues = checker.check_document(doc_local)
        logging.info(f'[check] 检查完成，问题数: {len(format_issues)}')

        # 4. 生成批注版
        annotator = DocumentAnnotator()
        annotated_path = annotator.generate_annotated_copy(
            original_path=doc_local,
            format_issues=format_issues,
            output_dir=tmp_dir,
            suffix='_批注版',
        )

        # 5. 生成 HTML 报告
        report_gen = ReportGenerator()
        report_path = report_gen.generate_html_report(
            file_path=doc_local,
            format_issues=format_issues,
            output_dir=tmp_dir,
            suffix='_检查报告',
        )

        # 6. 上传结果
        report_cloud = f"reports/{task_id}/report.html"
        annotated_cloud = f"reports/{task_id}/annotated.docx"
        report_file_id = storage.upload(report_path, report_cloud)
        annotated_file_id = storage.upload(annotated_path, annotated_cloud)
        if not report_file_id or not annotated_file_id:
            update_task(task_id, {'status': 'failed',
                                  'error': '结果文件上传失败'})
            return jsonify({'success': False, 'message': '结果文件上传失败'}), 500

        # 7. 更新任务
        update_task(task_id, {
            'status': 'completed',
            'issue_count': len(format_issues),
            'report_file_id': report_file_id,
            'annotated_file_id': annotated_file_id,
            'checked_at': datetime.now().isoformat(),
        })
        logging.info(f'[task] 检查完成: {task_id}, 问题数: {len(format_issues)}')
        return jsonify({'success': True, 'task_id': task_id,
                        'issue_count': len(format_issues),
                        'passed': len(format_issues) == 0,
                        'message': '检查完成'})
    except Exception as e:
        logging.exception(f'[check] 检查异常: {task_id}')
        update_task(task_id, {'status': 'failed', 'error': str(e)})
        return jsonify({'success': False, 'message': f'检查失败: {str(e)}'}), 500
    finally:
        _cleanup_tmp(tmp_dir)


# ============================================================
# API: 任务状态
# ============================================================
@app.route('/api/status/<task_id>', methods=['GET'])
def get_status(task_id):
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    return jsonify({
        'success': True,
        'task_id': task_id,
        'status': task.get('status'),
        'issue_count': task.get('issue_count', 0),
        'file_name': task.get('file_name') or task.get('doc_file_name'),
        'created_at': task.get('created_at'),
        'updated_at': task.get('updated_at'),
    })


# ============================================================
# API: 下载结果
# ============================================================
@app.route('/api/download/<task_id>/<file_type>', methods=['GET'])
def download_result(task_id, file_type):
    task = get_task(task_id)
    if not task:
        return jsonify({'success': False, 'message': '任务不存在'}), 404
    if task.get('status') != 'completed':
        return jsonify({'success': False, 'message': '检查尚未完成'}), 400

    storage = get_storage_client()
    if file_type == 'report':
        fileid = task.get('report_file_id')
        file_name = '格式检查报告.html'
    elif file_type == 'annotated':
        fileid = task.get('annotated_file_id')
        file_name = '批注版.docx'
    else:
        return jsonify({'success': False, 'message': f'不支持的类型: {file_type}'}), 400

    if not fileid:
        return jsonify({'success': False, 'message': '结果文件不存在'}), 404
    url = storage.get_download_url(fileid, expires=3600)
    if not url:
        return jsonify({'success': False, 'message': '获取下载链接失败'}), 500
    return jsonify({'success': True, 'download_url': url, 'file_name': file_name})


# ============================================================
# 入口
# ============================================================
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 80))
    logging.info('=' * 50)
    logging.info('  追标猎手 - 后端服务（云托管直连版）')
    logging.info('=' * 50)
    logging.info(f'  环境:{ENV_ID}  端口:{port}')
    logging.info('=' * 50)
    app.run(host='0.0.0.0', port=port, debug=False)

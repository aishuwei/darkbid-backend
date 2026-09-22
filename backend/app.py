# -*- coding: utf-8 -*-
"""
backend/app.py — 追标猎手 · 暗标格式检查后端（云托管直连版 · 配置化 + 异步检查）
架构：CloudRun (Container / Python) → NoSQL HTTP API + COS Storage
配置来源：
  - 内置配置：clients/config/*.json（打包在镜像里）+ clients/config/index.json（索引）
  - 自定义配置：云数据库 configs 集合（永久码用户上传）
"""
import os
import re
import sys
import json
import time
import uuid
import threading
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
# ⚠️ 配置目录指向 clients/config（你实际存放 guizhou.json 和 index.json 的地方）
CONFIGS_DIR = os.path.join(PROJECT_ROOT, 'clients', 'config')
DEFAULT_RULES_PATH = os.path.join(CONFIGS_DIR, 'rules.json')

logging.info(f'[boot] ENV_ID={ENV_ID}')
logging.info(f'[boot] CONFIGS_DIR={CONFIGS_DIR}')

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
# 配置管理
# ============================================================
# 配置里可能出现的元信息字段（不属于规则本身，不传给 format_checker）
META_KEYS = {'config_id', 'config_name', 'version', 'source', 'updated_at'}


def _extract_rules(cfg):
    """从配置 dict 里剥出纯规则，兼容两种格式：
       A) 有 rules 外壳：{config_id, config_name, rules: {...}} → 返回 cfg['rules']
       B) 平铺格式：{config_id, config_name, document_info: {...}, ...} → 剥元字段后返回
       C) 纯规则：{document_info: {...}, page_check: {...}} → 直接返回
    """
    if isinstance(cfg.get('rules'), dict):
        return cfg['rules']
    return {k: v for k, v in cfg.items() if k not in META_KEYS}


def _load_builtin_configs():
    """读取内置配置索引 clients/config/index.json"""
    index_path = os.path.join(CONFIGS_DIR, 'index.json')
    if not os.path.exists(index_path):
        logging.warning(f'[configs] 内置配置索引不存在: {index_path}')
        return []
    try:
        with open(index_path, 'r', encoding='utf-8') as f:
            items = json.load(f)
        return items if isinstance(items, list) else []
    except Exception as e:
        logging.error(f'[configs] 读取内置索引失败: {e}')
        return []


def _find_builtin_name(config_id):
    """从 index.json 里查显示名"""
    for item in _load_builtin_configs():
        if item.get('config_id') == config_id:
            return item.get('config_name')
    return None


def _get_builtin_config(config_id):
    """读取内置配置全文"""
    path = os.path.join(CONFIGS_DIR, f'{config_id}.json')
    if not os.path.exists(path):
        return None
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        logging.error(f'[configs] 读取内置配置 {config_id} 失败: {e}')
        return None


def _list_db_configs():
    """从数据库读取自定义配置"""
    try:
        docs = db_query('configs', limit=100) or []
        return docs
    except Exception as e:
        logging.error(f'[configs] 读取数据库配置失败: {e}')
        return []


def _get_db_config(config_id):
    """从数据库读取指定配置"""
    try:
        return db_query_one('configs', 'config_id', config_id)
    except Exception as e:
        logging.error(f'[configs] 查询数据库配置 {config_id} 失败: {e}')
        return None


def get_config_by_id(config_id):
    """优先查内置，再查数据库。返回 {'config_name': ..., 'rules': {...}} 或 None"""
    # 1. 内置配置
    cfg = _get_builtin_config(config_id)
    if cfg:
        # 显示名：优先配置文件里的 config_name，没有就从 index.json 查
        name = cfg.get('config_name') or _find_builtin_name(config_id) or config_id
        return {
            'config_id': config_id,
            'config_name': name,
            'rules': _extract_rules(cfg),
        }

    # 2. 数据库配置
    cfg = _get_db_config(config_id)
    if cfg:
        return {
            'config_id': config_id,
            'config_name': cfg.get('config_name') or config_id,
            'rules': _extract_rules(cfg),
        }

    return None


def list_all_configs():
    """合并内置 + 数据库配置列表，去重"""
    merged = []
    seen = set()

    # 内置配置列表
    for c in _load_builtin_configs():
        cid = c.get('config_id')
        if not cid or cid in seen:
            continue
        seen.add(cid)
        merged.append({
            'config_id': cid,
            'config_name': c.get('config_name') or cid,
        })

    # 数据库配置列表
    for c in _list_db_configs():
        cid = c.get('config_id')
        if not cid or cid in seen:
            continue
        seen.add(cid)
        merged.append({
            'config_id': cid,
            'config_name': c.get('config_name') or cid,
        })

    return merged


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


def _make_tmp_dir(prefix):
    return tempfile.mkdtemp(prefix=prefix)


def _cleanup_tmp(d):
    try:
        shutil.rmtree(d, ignore_errors=True)
    except Exception:
        pass


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
# API: 配置列表
# ============================================================
@app.route('/api/configs/list', methods=['GET'])
def api_configs_list():
    try:
        configs = list_all_configs()
        return jsonify({'success': True, 'configs': configs})
    except Exception as e:
        logging.exception('[configs/list] 失败')
        return jsonify({'success': False, 'message': str(e)}), 500


# ============================================================
# API: 一次性检查（file + config_id 或 rules_json）
# ============================================================
@app.route('/api/check', methods=['POST'])
def api_check():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': '未选择文件'}), 400
    file = request.files['file']
    if not file.filename or not file.filename.lower().endswith('.docx'):
        return jsonify({'success': False, 'message': '仅支持 .docx 格式'}), 400

    config_id = (request.form.get('config_id') or '').strip()
    rules_json_str = (request.form.get('rules_json') or '').strip()

    # 解析规则
    rules = None
    config_name = '自定义标准'

    if config_id:
        cfg = get_config_by_id(config_id)
        if not cfg:
            return jsonify({'success': False, 'message': f'配置 {config_id} 不存在'}), 404
        rules = cfg['rules']
        config_name = cfg['config_name']
    elif rules_json_str:
        try:
            rules = json.loads(rules_json_str)
        except Exception as e:
            return jsonify({'success': False, 'message': f'rules_json 解析失败: {e}'}), 400
        if not isinstance(rules, dict):
            return jsonify({'success': False, 'message': 'rules_json 必须是 JSON 对象'}), 400
        config_name = '自定义标准'
    else:
        return jsonify({'success': False, 'message': '必须提供 config_id 或 rules_json'}), 400

    # 校验文档
    task_id = generate_task_id()
    tmp_dir = _make_tmp_dir('doc_')
    tmp_path = os.path.join(tmp_dir, file.filename)
    try:
        file.save(tmp_path)
        size = os.path.getsize(tmp_path)
        if size == 0:
            return jsonify({'success': False, 'message': '文件为空'}), 400
        if size > MAX_FILE_SIZE:
            return jsonify({'success': False,
                            'message': f'文件超过 {MAX_FILE_SIZE / 1024 / 1024:.0f}MB 限制'}), 400
        valid, msg = validate_document(tmp_path)
        if not valid:
            return jsonify({'success': False, 'message': msg}), 400

        # 上传云存储
        storage = get_storage_client()
        cloud_path = f"documents/{task_id}/{file.filename}"
        doc_file_id = storage.upload(tmp_path, cloud_path)
        if not doc_file_id:
            return jsonify({'success': False, 'message': '文件上传失败'}), 500

        # 创建任务
        rules_json = json.dumps(rules, ensure_ascii=False)
        ok = create_task(task_id, {
            'status': 'checking',
            'config_id': config_id or '',
            'config_name': config_name,
            'doc_file_name': file.filename,
            'doc_file_id': doc_file_id,
            'rules_json': rules_json,
        })
        if not ok:
            return jsonify({'success': False, 'message': '任务创建失败'}), 500

        # 启动后台检查
        t = threading.Thread(target=_run_check_async, args=(task_id,), daemon=True)
        t.start()
        logging.info(f'[task] 检查已启动: {task_id}, 配置: {config_name}')
        return jsonify({'success': True, 'task_id': task_id, 'status': 'checking'})
    finally:
        _cleanup_tmp(tmp_dir)


# ============================================================
# API: 上传配置（永久码鉴权）
# ============================================================
@app.route('/api/admin/upload-config', methods=['POST'])
def api_admin_upload_config():
    data = request.get_json(silent=True) or {}

    perm_code = (data.get('perm_code') or '').strip().upper()
    if perm_code not in PERMANENT_CODES:
        logging.warning(f'[admin] 无效永久码尝试: {perm_code}')
        return jsonify({'success': False, 'message': '无效的管理员码'}), 403

    config_id = (data.get('config_id') or '').strip()
    config_name = (data.get('config_name') or '').strip()
    rules = data.get('rules')

    if not config_id or not config_name:
        return jsonify({'success': False, 'message': '缺少 config_id 或 config_name'}), 400
    if not isinstance(rules, dict):
        return jsonify({'success': False, 'message': 'rules 必须是 JSON 对象'}), 400
    if not re.match(r'^[a-zA-Z0-9_]{1,40}$', config_id):
        return jsonify({'success': False,
                        'message': 'config_id 只允许字母、数字、下划线，长度 1-40'}), 400

    # 不允许覆盖内置配置
    if _get_builtin_config(config_id):
        return jsonify({'success': False,
                        'message': f'config_id「{config_id}」已被内置配置占用'}), 409

    try:
        existing = _get_db_config(config_id)
        if existing:
            db_update_where('configs', 'config_id', config_id, {
                'config_name': config_name,
                'rules': rules,
                'updated_at': datetime.now().isoformat(),
            })
            logging.info(f'[admin] 更新配置: {config_id} - {config_name}')
        else:
            db_add('configs', {
                'config_id': config_id,
                'config_name': config_name,
                'rules': rules,
                'source': 'admin',
                'created_at': datetime.now().isoformat(),
            })
            logging.info(f'[admin] 新增配置: {config_id} - {config_name}')
    except Exception as e:
        logging.exception('[admin] 保存配置失败')
        return jsonify({'success': False, 'message': f'保存失败: {e}'}), 500

    return jsonify({'success': True, 'config_id': config_id,
                    'message': f'配置「{config_name}」已保存'})


# ============================================================
# 后台线程：执行真正的检查逻辑
# ============================================================
def _run_check_async(task_id):
    task = get_task(task_id)
    if not task:
        logging.error(f'[check-async] 任务不存在: {task_id}')
        return

    doc_file_id = task.get('doc_file_id')
    if not doc_file_id:
        update_task(task_id, {'status': 'failed', 'error': 'fileid 缺失'})
        return

    tmp_dir = _make_tmp_dir('chk_')
    try:
        storage = get_storage_client()

        # 1. 下载文档
        doc_name = task.get('doc_file_name') or 'document.docx'
        doc_local = os.path.join(tmp_dir, doc_name)
        download_url = storage.get_download_url(doc_file_id, expires=600)
        if not download_url:
            update_task(task_id, {'status': 'failed', 'error': '获取下载链接失败'})
            return
        import requests as _rq
        r = _rq.get(download_url, timeout=120)
        if r.status_code != 200:
            update_task(task_id, {'status': 'failed', 'error': '文档下载失败'})
            return
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
                update_task(task_id, {'status': 'failed', 'error': '找不到规则文件'})
                return
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
            update_task(task_id, {'status': 'failed', 'error': '结果文件上传失败'})
            return

        # 7. 更新任务
        update_task(task_id, {
            'status': 'completed',
            'issue_count': len(format_issues),
            'report_file_id': report_file_id,
            'annotated_file_id': annotated_file_id,
            'checked_at': datetime.now().isoformat(),
        })
        logging.info(f'[task] 检查完成: {task_id}, 问题数: {len(format_issues)}')
    except Exception as e:
        logging.exception(f'[check-async] 检查异常: {task_id}')
        update_task(task_id, {'status': 'failed', 'error': str(e)})
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
        'error': task.get('error'),
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
    logging.info('  追标猎手 - 后端服务（配置化 + 异步检查版）')
    logging.info('=' * 50)
    logging.info(f'  环境:{ENV_ID}  端口:{port}')
    logging.info('=' * 50)
    app.run(host='0.0.0.0', port=port, debug=False)

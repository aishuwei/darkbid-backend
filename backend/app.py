	# -*- coding: utf-8 -*-
	"""
	backend/app.py — 追标猎手 · 暗标格式检查后端（微信云托管可用版）
	架构：
	1. 数据库：云开发数据库（tasks / used_codes 集合），走微信开放接口 /tcb/*
	   云托管内调用免 access_token（留空自动注入）；
	   若失效，配置环境变量 WX_APPID / WX_SECRET 后自动改为自行获取 token。
	2. 文件：全部存云存储，数据库只存 file_id；容器内仅用临时目录（用完即删），
	   任意实例可处理任意任务，天然支持多实例。
	3. 端口：默认监听 80，云托管服务端口需一致（或配环境变量 PORT 覆盖）。
	4. 接口：与原版路径、参数、返回结构完全一致，小程序端无需修改。
	"""
	import os
	import sys
	import json
	import time
	import uuid
	import shutil
	import tempfile
	import requests
	from datetime import datetime
	from flask import Flask, request, jsonify
	# --- 项目路径 ---
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
	ENV_ID = os.environ.get('TCB_ENV', 'darkbid-d8gxpsbued4a2867')  # 云托管所在的环境ID
	WX_API = 'https://api.weixin.qq.com'
	APPID = os.environ.get('WX_APPID', '')    # 可选：免鉴权失效时才需要
	SECRET = os.environ.get('WX_SECRET', '')  # 可选：同上
	app = Flask(__name__)
	try:
	    from flask_cors import CORS
	    CORS(app)
	except ImportError:
	    pass  # callContainer 不依赖 CORS，装不装都能跑
	class DbError(Exception):
	    """云数据库访问失败"""
	@app.errorhandler(DbError)
	def handle_db_error(e):
	    print(f'[db] 全局异常: {e}')
	    return jsonify({'success': False, 'message': '云数据库暂不可用，请稍后重试'}), 500
	# ============================================================
	# 微信开放接口调用（access_token 管理）
	# ============================================================
	_token_cache = {'token': '', 'expire_at': 0}
	def get_access_token():
	    """有 AppID/Secret 则自取并缓存；否则返回空串（云托管内由网关自动注入）"""
	    if not APPID or not SECRET:
	        return ''
	    if _token_cache['token'] and time.time() < _token_cache['expire_at']:
	        return _token_cache['token']
	    try:
	        r = requests.get(f'{WX_API}/cgi-bin/token', params={
	            'grant_type': 'client_credential',
	            'appid': APPID, 'secret': SECRET,
	        }, timeout=10).json()
	    except Exception as e:
	        print(f'[token] 获取失败: {e}')
	        return ''
	    if 'access_token' not in r:
	        print(f'[token] 接口返回异常: {r}')
	        return ''
	    _token_cache['token'] = r['access_token']
	    _token_cache['expire_at'] = time.time() + r.get('expires_in', 7200) - 300
	    return _token_cache['token']
	def call_wx_api(path, payload):
	    """POST 微信开放接口（/tcb/* 系列），失败返回 None"""
	    try:
	        resp = requests.post(f'{WX_API}{path}',
	                             params={'access_token': get_access_token()},
	                             json=payload, timeout=15)
	        result = resp.json()
	    except Exception as e:
	        print(f'[wxapi] {path} 请求异常: {e}')
	        return None
	    if result.get('errcode', 0) != 0:
	        print(f'[wxapi] {path} 失败: {result}')
	        return None
	    return result
	# ============================================================
	# 云数据库操作（/tcb/database* 系列）
	# 注意：databasequery 返回的 data 是 JSON 字符串数组
	# ============================================================
	def db_add(collection, doc):
	    query = 'db.collection("%s").add({data: %s})' % (
	        collection, json.dumps(doc, ensure_ascii=False))
	    return call_wx_api('/tcb/databaseadd', {'env': ENV_ID, 'query': query})
	def db_query_one(collection, field, value):
	    query = 'db.collection("%s").where({%s: %s}).limit(1).get()' % (
	        collection, field, json.dumps(str(value), ensure_ascii=False))
	    r = call_wx_api('/tcb/databasequery', {'env': ENV_ID, 'query': query})
	    if r is None:
	        raise DbError(f'查询 {collection} 失败')
	    data = r.get('data') or []
	    return json.loads(data[0]) if data else None
	def db_update_where(collection, field, value, data):
	    query = 'db.collection("%s").where({%s: %s}).update({data: %s})' % (
	        collection, field, json.dumps(str(value), ensure_ascii=False),
	        json.dumps(data, ensure_ascii=False))
	    r = call_wx_api('/tcb/databaseupdate', {'env': ENV_ID, 'query': query})
	    if r is None:
	        raise DbError(f'更新 {collection} 失败')
	    return True
	# ============================================================
	# 云存储操作（/tcb/uploadfile + /tcb/batchdownloadfile）
	# ============================================================
	def storage_upload(local_path, cloud_path):
	    """上传本地文件到云存储，成功返回 file_id（cloud://...）"""
	    r = call_wx_api('/tcb/uploadfile', {'env': ENV_ID, 'path': cloud_path})
	    if not r or 'url' not in r:
	        print(f'[storage] 获取上传凭证失败: {r}')
	        return None
	    try:
	        with open(local_path, 'rb') as f:
	            resp = requests.put(r['url'], data=f, headers={
	                'Authorization': r.get('authorization', ''),
	                'x-cos-security-token': r.get('token', ''),
	                'x-cos-meta-fileid': r.get('cos_file_id', ''),
	            }, timeout=120)
	    except Exception as e:
	        print(f'[storage] 上传异常: {e}')
	        return None
	    if resp.status_code in (200, 204):
	        file_id = r.get('fileid')
	        print(f'[storage] 上传成功: {file_id}')
	        return file_id
	    print(f'[storage] 上传失败 HTTP {resp.status_code}: {resp.text[:200]}')
	    return None
	def _get_download_url(file_id):
	    r = call_wx_api('/tcb/batchdownloadfile',
	                    {'env': ENV_ID, 'fileid_list': [file_id]})
	    if r and r.get('file_list'):
	        item = r['file_list'][0]
	        if item.get('status') == 0:
	            return item.get('download_url')
	    print(f'[storage] 获取下载链接失败: {r}')
	    return None
	def storage_download(file_id, local_path):
	    """从云存储下载文件到本地路径"""
	    url = _get_download_url(file_id)
	    if not url:
	        return False
	    try:
	        resp = requests.get(url, timeout=120)
	        if resp.status_code == 200:
	            with open(local_path, 'wb') as f:
	                f.write(resp.content)
	            return True
	    except Exception as e:
	        print(f'[storage] 下载异常: {e}')
	    return False
	# ============================================================
	# 任务 / 激活码 业务封装（全部落库，不再依赖容器本地路径）
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
	# 工具函数
	# ============================================================
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
	def extract_text(path):
	    """读取要求文件文本（.txt/.md 直接读，.docx 提取段落）"""
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
	    db_ok = False
	    try:
	        r = call_wx_api('/tcb/databasequery',
	                        {'env': ENV_ID, 'query': 'db.collection("tasks").limit(1).get()'})
	        db_ok = r is not None
	    except Exception:
	        pass
	    return jsonify({'status': 'ok', 'message': '服务运行中',
	                    'env': ENV_ID, 'db_ok': db_ok})
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
	    # 文本直接存数据库字段，天然跨实例
	    if not create_task(task_id, {'status': 'req_uploaded',
	                                 'requirement_type': 'text',
	                                 'requirement_text': text}):
	        return jsonify({'success': False, 'message': '任务创建失败（数据库不可用）'}), 500
	    print(f'[task] 文本任务创建成功: {task_id}')
	    return jsonify({'success': True, 'task_id': task_id, 'message': '格式要求已接收'})
	@app.route('/api/upload-req-file', methods=['POST'])
	def upload_requirement_file():
	    if 'file' not in request.files:
	        return jsonify({'success': False, 'message': '没有上传文件'}), 400
	    file = request.files['file']
	    if not file.filename:
	        return jsonify({'success': False, 'message': '文件名为空'}), 400
	    task_id = generate_task_id()
	    tmp_dir = tempfile.mkdtemp(prefix='req_')
	    try:
	        ext = os.path.splitext(file.filename)[1].lower()
	        tmp_path = os.path.join(tmp_dir, 'requirement' + ext)
	        file.save(tmp_path)
	        file_id = storage_upload(tmp_path, f'{task_id}/requirement{ext}')
	        if not file_id:
	            return jsonify({'success': False, 'message': '要求文件保存失败（云存储不可用）'}), 500
	        if not create_task(task_id, {'status': 'req_uploaded',
	                                     'requirement_type': 'file',
	                                     'req_file_id': file_id,
	                                     'file_name': file.filename}):
	            return jsonify({'success': False, 'message': '任务创建失败（数据库不可用）'}), 500
	        print(f'[task] 文件任务创建成功: {task_id}')
	        return jsonify({'success': True, 'task_id': task_id, 'message': '文件已接收'})
	    finally:
	        shutil.rmtree(tmp_dir, ignore_errors=True)
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
	        return jsonify({'success': False, 'message': '任务状态异常，请重新提交'}), 400
	    # --- 读取格式要求内容 ---
	    if task.get('requirement_type') == 'text':
	        req_content = task.get('requirement_text') or ''
	    else:
	        file_id = task.get('req_file_id')
	        if not file_id:
	            return jsonify({'success': False, 'message': '格式要求文件不存在'}), 400
	        tmp_dir = tempfile.mkdtemp(prefix='reqdl_')
	        try:
	            ext = os.path.splitext(task.get('file_name') or '')[1].lower()
	            tmp_path = os.path.join(tmp_dir, 'requirement' + ext)
	            if not storage_download(file_id, tmp_path):
	                return jsonify({'success': False, 'message': '要求文件下载失败'}), 500
	            req_content = extract_text(tmp_path)
	        finally:
	            shutil.rmtree(tmp_dir, ignore_errors=True)
	    # --- 解析要求生成规则 ---
	    # TODO: 在这里接入真正的「要求文本 → 规则」解析逻辑，替换下面的模拟规则
	    rules = {
	        'document_info': {
	            'name': '模拟规则',
	            'generated_date': datetime.now().strftime('%Y-%m-%d'),
	        },
	        'source_text_length': len(req_content),
	    }
	    rules_json = json.dumps(rules, ensure_ascii=False, indent=2)
	    # 规则 JSON 直接存数据库字段（start-check 时再落成临时文件供 FormatChecker 使用）
	    if not update_task(task_id, {'status': 'rules_generated', 'rules_json': rules_json}):
	        return jsonify({'success': False, 'message': '规则保存失败'}), 500
	    print(f'[task] 规则生成成功: {task_id}')
	    return jsonify({'success': True, 'message': '配置已生成', 'task_id': task_id})
	@app.route('/api/upload-doc', methods=['POST'])
	def upload_document():
	    task_id = request.form.get('taskId') or request.form.get('task_id')
	    if not task_id:
	        return jsonify({'success': False, 'message': '缺少任务ID'}), 400
	    task = get_task(task_id)
	    if not task:
	        return jsonify({'success': False, 'message': '任务不存在'}), 404
	    if 'file' not in request.files:
	        return jsonify({'success': False, 'message': '没有上传文件'}), 400
	    file = request.files['file']
	    if not file.filename:
	        return jsonify({'success': False, 'message': '文件名为空'}), 400
	    if not file.filename.lower().endswith('.docx'):
	        return jsonify({'success': False, 'message': '仅支持 .docx 格式'}), 400
	    file.seek(0, os.SEEK_END)
	    file_size = file.tell()
	    file.seek(0)
	    if file_size == 0:
	        return jsonify({'success': False, 'message': '文件为空'}), 400
	    if file_size > MAX_FILE_SIZE:
	        return jsonify({'success': False,
	                        'message': f"文件大小{file_size/1024/1024:.2f}MB，"
	                                   f"超过限制{MAX_FILE_SIZE/1024/1024:.0f}MB"}), 400
	    tmp_dir = tempfile.mkdtemp(prefix='doc_')
	    try:
	        doc_path = os.path.join(tmp_dir, 'document.docx')
	        file.save(doc_path)
	        valid, msg = validate_document(doc_path)
	        if not valid:
	            return jsonify({'success': False, 'message': msg}), 400
	        # 校验通过 → 存云存储，数据库记 file_id
	        file_id = storage_upload(doc_path, f'{task_id}/document.docx')
	        if not file_id:
	            return jsonify({'success': False, 'message': '文件保存失败（云存储不可用）'}), 500
	        if not update_task(task_id, {'status': 'doc_uploaded',
	                                     'doc_file_id': file_id,
	                                     'doc_name': file.filename}):
	            return jsonify({'success': False, 'message': '任务更新失败'}), 500
	        print(f'[task] 文档上传成功: {task_id}')
	        return jsonify({'success': True, 'message': '文件上传成功', 'task_id': task_id})
	    finally:
	        shutil.rmtree(tmp_dir, ignore_errors=True)
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
	        return jsonify({'success': False, 'message': '请先上传技术文件'}), 400
	    doc_file_id = task.get('doc_file_id')
	    if not doc_file_id:
	        return jsonify({'success': False, 'message': '技术文件不存在，请重新上传'}), 400
	    tmp_dir = tempfile.mkdtemp(prefix='chk_')
	    try:
	        update_task(task_id, {'status': 'checking'})
	        # 1. 从云存储取回待检文档
	        doc_path = os.path.join(tmp_dir, 'document.docx')
	        if not storage_download(doc_file_id, doc_path):
	            update_task(task_id, {'status': 'doc_uploaded'})
	            return jsonify({'success': False, 'message': '技术文件下载失败，请重试'}), 500
	        # 2. 规则：优先任务自带 rules_json，否则回退默认规则文件
	        rules_path = os.path.join(tmp_dir, 'rules.json')
	        rules_json = task.get('rules_json')
	        if rules_json:
	            with open(rules_path, 'w', encoding='utf-8') as f:
	                f.write(rules_json)
	        else:
	            fallback = os.path.join(PROJECT_ROOT, 'clients', 'config', 'rules.json')
	            if not os.path.exists(fallback):
	                update_task(task_id, {'status': 'doc_uploaded'})
	                return jsonify({'success': False, 'message': '找不到规则配置文件'}), 500
	            rules_path = fallback
	        # 3. 执行检查（业务模块不变）
	        checker = FormatChecker(rules_path)
	        format_issues = checker.check_document(doc_path)
	        annotator = DocumentAnnotator()
	        annotated_path = annotator.generate_annotated_copy(
	            original_path=doc_path, format_issues=format_issues,
	            output_dir=tmp_dir, suffix='_批注版')
	        report_gen = ReportGenerator()
	        report_path = report_gen.generate_html_report(
	            file_path=doc_path, format_issues=format_issues,
	            output_dir=tmp_dir, suffix='_检查报告')
	        # 4. 结果上传云存储
	        report_file_id = storage_upload(report_path, f'{task_id}/report.html')
	        annotated_file_id = storage_upload(annotated_path, f'{task_id}/annotated.docx')
	        if not report_file_id or not annotated_file_id:
	            update_task(task_id, {'status': 'failed', 'error': '结果文件上传云存储失败'})
	            return jsonify({'success': False, 'message': '结果文件保存失败，请重试'}), 500
	        # 5. 任务收尾
	        update_task(task_id, {
	            'status': 'completed',
	            'issue_count': len(format_issues),
	            'report_file_id': report_file_id,
	            'annotated_file_id': annotated_file_id,
	            'completed_at': datetime.now().isoformat(),
	        })
	        print(f'[task] 检查完成: {task_id}, 问题数: {len(format_issues)}')
	        return jsonify({'success': True, 'message': '检查完成', 'task_id': task_id,
	                        'issue_count': len(format_issues),
	                        'passed': len(format_issues) == 0})
	    except Exception as e:
	        import traceback
	        update_task(task_id, {'status': 'failed', 'error': str(e),
	                              'traceback': traceback.format_exc()[:2000]})
	        return jsonify({'success': False, 'message': f'检查失败: {str(e)}'}), 500
	    finally:
	        shutil.rmtree(tmp_dir, ignore_errors=True)
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
	    task = get_task(task_id)
	    if not task:
	        return jsonify({'success': False, 'message': '任务不存在'}), 404
	    if task.get('status') != 'completed':
	        return jsonify({'success': False, 'message': '检查尚未完成'}), 400
	    if file_type == 'report':
	        file_id, file_name = task.get('report_file_id'), '格式检查报告.html'
	    elif file_type == 'annotated':
	        file_id, file_name = task.get('annotated_file_id'), '批注版.docx'
	    else:
	        return jsonify({'success': False, 'message': '无效的文件类型'}), 400
	    if not file_id:
	        return jsonify({'success': False, 'message': '文件不存在'}), 404
	    url = _get_download_url(file_id)
	    if not url:
	        return jsonify({'success': False, 'message': '获取下载链接失败'}), 500
	    return jsonify({'success': True, 'download_url': url, 'file_name': file_name})
	if __name__ == '__main__':
	    port = int(os.environ.get('PORT', 80))
	    print('=' * 50)
	    print(' 追标猎手 - 暗标格式检查后端服务（云托管版）')
	    print('=' * 50)
	    print(f' 云开发环境: {ENV_ID}')
	    print(f' 监听端口: {port}（云托管服务端口需与此一致）')
	    print(f' access_token 模式: {"自行获取" if APPID and SECRET else "云托管免鉴权"}')
	    print('=' * 50)
	    app.run(host='0.0.0.0', port=port, debug=False)

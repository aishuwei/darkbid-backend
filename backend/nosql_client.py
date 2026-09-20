# -*- coding: utf-8 -*-
"""
backend/nosql_client.py — CloudBase 文档型数据库（NoSQL）HTTP API 客户端
使用 CloudBase NoSQL RESTful HTTP API 直接操作文档数据库。
API Key --> AccessToken (OAuth2) --> NoSQL HTTP API
"""
import os, json, time, logging
import requests

TOKEN_URL = "https://tcb-api.tencentcloudapi.com/auth/oauth/token"
logger = logging.getLogger("nosql_client")

# ---------- EJSON ----------
def _to_ejson(obj):
    if isinstance(obj, dict):
        return {k: _to_ejson(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_to_ejson(v) for v in obj]
    return obj

def _from_ejson(obj):
    if isinstance(obj, dict):
        if "$oid" in obj: return obj["$oid"]
        if "$date" in obj and isinstance(obj["$date"], dict) and "$numberLong" in obj["$date"]:
            return time.strftime("%Y-%m-%dT%H:%M:%S", time.gmtime(int(obj["$date"]["$numberLong"]) / 1000))
        if "$numberInt" in obj: return int(obj["$numberInt"])
        if "$numberLong" in obj: return int(obj["$numberLong"])
        return {k: _from_ejson(v) for k, v in obj.items()}
    if isinstance(obj, list): return [_from_ejson(v) for v in obj]
    return obj

# ---------- Token ----------
class TokenManager:
    def __init__(self, env_id: str, api_key: str):
        self.env_id, self.api_key = env_id, api_key
        self._token: str | None = None
        self._expires_at: float = 0.0

    @property
    def token(self) -> str:
        if time.time() + 300 > self._expires_at:
            self._refresh()
        return self._token

    def _refresh(self):
        logger.info("[token] 获取 AccessToken ...")
        resp = requests.post(TOKEN_URL, json={
            "grant_type": "client_credentials",
            "client_id": self.env_id,
            "client_secret": self.api_key,
        }, timeout=10)
        data = resp.json()
        if "access_token" not in data:
            raise RuntimeError(f"获取 AccessToken 失败: {data}")
        self._token = data["access_token"]
        self._expires_at = time.time() + data.get("expires_in", 7200)
        logger.info(f"[token] 成功，有效期 {data.get('expires_in', 7200)}s")

# ---------- NoSQL Client ----------
class NoSQLClient:
    def __init__(self, env_id: str, api_key: str, instance_id="(default)", database="(default)"):
        self.env_id = env_id
        self.instance_id = instance_id
        self._token_mgr = TokenManager(env_id, api_key)
        self._base_url = (f"https://{env_id}.api.tcloudbasegateway.com"
                          f"/v1/database/instances/{instance_id}/databases/{database}")

    def _headers(self):
        return {"Authorization": f"Bearer {self._token_mgr.token}", "Content-Type": "application/json"}

    def _request(self, method, path, **kwargs):
        url = f"{self._base_url}{path}"
        resp = requests.request(method, url, headers=self._headers(), timeout=15, **kwargs)
        body = resp.json()
        if resp.status_code >= 400:
            logger.error(f"[nosql] 失败 [{resp.status_code}] {body.get('code','')}: {body.get('message','')}")
            return None
        return body

    def query(self, collection, where=None, limit=100, offset=0, order=None, count=False):
        params = {"limit": limit, "offset": offset, "count": "true" if count else "false"}
        if where: params["query"] = json.dumps(_to_ejson(where), ensure_ascii=False)
        if order: params["order"] = json.dumps(order, ensure_ascii=False)
        result = self._request("GET", f"/collections/{collection}/documents", params=params)
        if result is None: return None
        if count: return int(result.get("total", 0))
        return [_from_ejson(d) for d in result.get("list", [])]

    def query_one(self, collection, field, value):
        docs = self.query(collection, {field: value}, limit=1)
        return docs[0] if isinstance(docs, list) and docs else None

    def add(self, collection, data):
        result = self._request("POST", f"/collections/{collection}/documents",
                               json={"data": [_to_ejson(data)]})
        if result is None: return None
        ids = result.get("insertedIds", [])
        return ids[0] if ids else None

    def update(self, collection, query_field, query_value, update_data, multi=False):
        has_op = any(k.startswith("$") for k in update_data)
        payload = {
            "query": _to_ejson({query_field: query_value}),
            "data": _to_ejson(update_data) if has_op else {"$set": _to_ejson(update_data)},
            "multi": multi,
        }
        result = self._request("PATCH", f"/collections/{collection}/documents", json=payload)
        if result is None: return None
        return int(result.get("matched", 0))

    def delete(self, collection, query_field, query_value, multi=False):
        payload = {"query": _to_ejson({query_field: query_value}), "multi": multi}
        result = self._request("POST", f"/collections/{collection}/documents/remove", json=payload)
        if result is None: return None
        return int(result.get("deleted", 0))

# ---------- 单例 ----------
_client = None
def get_client():
    global _client
    if _client is not None: return _client
    env_id = os.environ.get("TCB_ENV")
    api_key = os.environ.get("CLOUDBASE_APIKEY")
    instance_id = os.environ.get("NOSQL_INSTANCE_ID", "(default)")
    if not env_id: raise RuntimeError("缺少 TCB_ENV")
    if not api_key: raise RuntimeError("缺少 CLOUDBASE_APIKEY")
    _client = NoSQLClient(env_id, api_key, instance_id)
    return _client

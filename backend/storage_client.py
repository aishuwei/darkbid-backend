# -*- coding: utf-8 -*-
"""
backend/storage_client.py — CloudBase 云存储 HTTP API 客户端
直接使用 API Key 作为 Bearer Token。
"""
import os
import logging
import requests

logger = logging.getLogger("storage_client")


class StorageClient:
    def __init__(self, env_id: str, api_key: str):
        self.env_id = env_id
        self.api_key = api_key
        self._base_url = f"https://{env_id}.api.tcloudbasegateway.com"

    def _headers(self):
        return {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }

    def upload(self, local_path: str, cloud_path: str):
        """上传文件到云存储，成功返回完整 fileid（cloud://...），失败返回 None"""
        if not os.path.isfile(local_path):
            logger.error(f"[storage] 文件不存在: {local_path}")
            return None
        logger.info(f"[storage] 获取上传信息: {cloud_path}")
        resp = requests.post(
            f"{self._base_url}/v1/storages/get-objects-upload-info",
            headers=self._headers(),
            json=[{"objectId": cloud_path}],
            timeout=15,
        )
        data = resp.json()
        if not isinstance(data, list) or not data:
            logger.error(f"[storage] 获取上传信息失败: {data}")
            return None
        item = data[0]
        if "code" in item:
            logger.error(f"[storage] 返回错误: {item}")
            return None

        with open(local_path, "rb") as f:
            file_data = f.read()

        upload_resp = requests.put(
            item["uploadUrl"],
            data=file_data,
            headers={
                "Authorization": item.get("authorization", ""),
                "X-Cos-Security-Token": item.get("token", ""),
                "X-Cos-Meta-Fileid": item.get("cloudObjectMeta", ""),
            },
            timeout=120,
        )
        if upload_resp.status_code not in (200, 201):
            logger.error(f"[storage] 上传失败: {upload_resp.status_code} {upload_resp.text[:200]}")
            return None

        # 从上传返回里拿 fileid；拿不到就按标准格式拼
        fileid = item.get("fileid") or ""
        if not fileid.startswith("cloud://"):
            # 尝试从 cloudObjectMeta 里解析
            meta = item.get("cloudObjectMeta", "") or ""
            for part in meta.split("&"):
                if part.startswith("x-cos-meta-fileid="):
                    fileid = part.split("=", 1)[1]
                    break
        if not fileid.startswith("cloud://"):
            # 兜底：标准格式（env_id.bucket/path）
            bucket = f"{self.env_id}"
            fileid = f"cloud://{self.env_id}.{bucket}/{cloud_path}"

        logger.info(f"[storage] 上传成功: {cloud_path} ({len(file_data)} bytes)")
        logger.info(f"[storage] fileid={fileid}")
        return fileid

    def get_download_url(self, fileid: str, expires=3600):
        """传入完整 fileid，返回下载链接"""
        logger.info(f"[storage] 请求下载链接: {fileid}")
        resp = requests.post(
            f"{self._base_url}/v1/storages/get-objects-download-info",
            headers=self._headers(),
            json=[{"cloudObjectId": fileid}],
            timeout=15,
        )
        try:
            data = resp.json()
        except Exception:
            data = {"raw": resp.text[:500]}
        logger.info(f"[storage] 下载信息 HTTP {resp.status_code} 返回: {str(data)[:500]}")
        if not isinstance(data, list) or not data:
            logger.error(f"[storage] 返回不是列表或为空: {data}")
            return None
        item = data[0]
        if "code" in item:
            logger.error(f"[storage] 接口返回错误: {item}")
            return None
        url = item.get("downloadUrl") or item.get("download_url") or item.get("url")
        if not url:
            logger.error(f"[storage] 未找到下载链接字段, keys={list(item.keys())}")
            return None
        return url

    def delete(self, fileid: str) -> bool:
        resp = requests.post(
            f"{self._base_url}/v1/storages/delete-objects",
            headers=self._headers(),
            json=[{"cloudObjectId": fileid}],
            timeout=15,
        )
        data = resp.json()
        return isinstance(data, list) and "code" not in data[0] if data else False


_client = None


def get_storage_client():
    global _client
    if _client is not None:
        return _client
    env_id = os.environ.get("TCB_ENV")
    api_key = os.environ.get("CLOUDBASE_APIKEY")
    if not env_id:
        raise RuntimeError("缺少 TCB_ENV")
    if not api_key:
        raise RuntimeError("缺少 CLOUDBASE_APIKEY")
    _client = StorageClient(env_id, api_key)
    return _client

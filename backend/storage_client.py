# -*- coding: utf-8 -*-
"""
backend/storage_client.py — CloudBase 云存储 HTTP API 客户端
直接使用 API Key 作为 Bearer Token，无需 TokenManager。
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

    def upload(self, local_path: str, cloud_path: str) -> bool:
        if not os.path.isfile(local_path):
            logger.error(f"文件不存在: {local_path}")
            return False
        logger.info(f"获取上传信息: {cloud_path}")
        resp = requests.post(
            f"{self._base_url}/v1/storages/get-objects-upload-info",
            headers=self._headers(),
            json=[{"objectId": cloud_path}],
            timeout=15,
        )
        data = resp.json()
        if not isinstance(data, list) or not data:
            logger.error(f"获取上传信息失败: {data}")
            return False
        item = data[0]
        if "code" in item:
            logger.error(f"返回错误: {item}")
            return False
        with open(local_path, "rb") as f:
            file_data = f.read()
        upload_resp = requests.put(
            item["uploadUrl"],
            data=file_data,
            headers={
                "Authorization": item["authorization"],
                "X-Cos-Security-Token": item["token"],
                "X-Cos-Meta-Fileid": item["cloudObjectMeta"],
            },
            timeout=120,
        )
        if upload_resp.status_code not in (200, 201):
            logger.error(f"上传失败: {upload_resp.status_code} {upload_resp.text[:200]}")
            return False
        logger.info(f"上传成功: {cloud_path} ({len(file_data)} bytes)")
        return True

    def get_download_url(self, cloud_path: str, expires=3600):
        resp = requests.post(
            f"{self._base_url}/v1/storages/get-objects-download-info",
            headers=self._headers(),
            json=[{"cloudObjectId": cloud_path}],
            timeout=15,
        )
        data = resp.json()
        if not isinstance(data, list) or not data:
            return None
        item = data[0]
        return None if "code" in item else item.get("downloadUrl")

    def delete(self, cloud_path: str) -> bool:
        resp = requests.post(
            f"{self._base_url}/v1/storages/delete-objects",
            headers=self._headers(),
            json=[{"cloudObjectId": cloud_path}],
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

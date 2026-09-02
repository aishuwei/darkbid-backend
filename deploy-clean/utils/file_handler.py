"""
utils/file_handler.py - 文件输入输出处理
职责：管理项目目录、文件路径、输出目录创建等
"""
import os
import shutil
from pathlib import Path


class FileHandler:
    """文件处理器：统一管理输入输出路径"""

    def __init__(self, project_root=None):
        """
        :param project_root: 项目根目录，默认当前工作目录
        """
        self.project_root = Path(project_root) if project_root else Path.cwd()
        self.config_dir = self.project_root / "config"
        self.output_dir = self.project_root / "output"
        self._ensure_dirs()

    def _ensure_dirs(self):
        """确保必要目录存在"""
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def get_rules_path(self, filename="rules.json"):
        """获取规则文件路径"""
        return str(self.config_dir / filename)

    def get_output_path(self, filename):
        """获取输出文件路径"""
        self.output_dir.mkdir(parents=True, exist_ok=True)
        return str(self.output_dir / filename)

    def resolve_input_path(self, file_path):
        """
        解析输入文件路径
        支持相对路径（相对于项目根目录）和绝对路径
        """
        path = Path(file_path)
        if not path.is_absolute():
            path = self.project_root / path
        return str(path.resolve())

    def clean_output(self):
        """清空输出目录"""
        if self.output_dir.exists():
            shutil.rmtree(self.output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

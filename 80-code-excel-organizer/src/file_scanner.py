# src/file_scanner.py
"""
文件扫描模块 - 扫描文件夹中的所有Excel文件
"""

from pathlib import Path
from typing import List
import hashlib


class FileScanner:
    """文件扫描器"""
    
    def __init__(self, logger):
        """初始化扫描器"""
        self.logger = logger
        self.excel_extensions = ('.xlsx', '.xlsm', '.xls')
    
    def scan_folder(self, folder_path: Path) -> List[Path]:
        """
        扫描文件夹中的所有Excel文件
        
        参数：folder_path - 文件夹路径
        返回：Excel文件路径列表（已排序）
        """
        excel_files = []
        temp_files = 0
        non_excel_files = 0
        
        try:
            for file_path in folder_path.iterdir():
                if not file_path.is_file():
                    continue
                
                file_name = file_path.name
                file_lower = file_name.lower()
                
                # 过滤临时文件
                if file_name.startswith('~$'):
                    temp_files += 1
                    continue
                
                # 检查Excel文件
                if file_lower.endswith(self.excel_extensions):
                    if self._is_file_accessible(file_path):
                        excel_files.append(file_path)
                    else:
                        print(f"警告：无法访问文件 '{file_name}'")
                        self.logger.add_error(file_name, "文件无法访问")
                else:
                    if file_path.suffix:
                        non_excel_files += 1
            
            excel_files.sort(key=lambda x: x.name)
            
            if temp_files > 0:
                print(f"发现 {temp_files} 个临时文件（已忽略）")
            
            return excel_files
            
        except Exception as e:
            raise Exception(f"扫描文件夹失败：{str(e)}")
    
    def _is_file_accessible(self, file_path: Path) -> bool:
        """检查文件是否可访问"""
        try:
            with open(file_path, 'rb') as f:
                f.read(4)
            return True
        except:
            return False
    
    def calculate_file_hash(self, file_path: Path) -> str:
        """计算文件的MD5哈希值，用于检测重复文件"""
        try:
            hash_md5 = hashlib.md5()
            with open(file_path, "rb") as f:
                for chunk in iter(lambda: f.read(4096), b""):
                    hash_md5.update(chunk)
            return hash_md5.hexdigest()
        except:
            return None

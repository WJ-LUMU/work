# src/data_processor.py
"""
数据处理核心模块

负责：
1. 合并重复80码
2. 按标准顺序排列
3. 识别新增80码
4. 核对总数
"""

from typing import List, Dict
from collections import defaultdict


class DataProcessor:
    """数据处理类"""
    
    def __init__(self):
        """初始化处理器"""
        pass
    
    def process(self, standard_codes: List[str], report_data: List[Dict]) -> Dict:
        """
        处理数据的主方法
        
        参数：
            standard_codes: 标准80码清单
            report_data: 生产日报数据
        
        返回：
            包含处理结果的字典
        """
        # 第1步：合并重复的80码（求和）
        merged_data = self._merge_duplicates(report_data)
        
        # 第2步：计算原始总数
        original_total = sum(item['quantity'] for item in report_data)
        
        # 第3步：按标准顺序整理数据
        standard_result = self._arrange_by_standard(standard_codes, merged_data)
        
        # 第4步：识别新增80码
        new_codes = self._find_new_codes(standard_codes, merged_data)
        
        # 第5步：核对总数
        standard_total = sum(item['quantity'] for item in standard_result)
        new_total = sum(item['quantity'] for item in new_codes)
        final_total = standard_total + new_total
        
        match_status = "核对一致" if original_total == final_total else "核对不一致"
        difference = original_total - final_total
        
        # 组织结果
        result = {
            'standard_result': self._format_with_sequence(standard_result, "是"),
            'new_codes': self._format_with_sequence(new_codes, "否"),
            'summary': {
                'original_total': original_total,
                'standard_total': standard_total,
                'new_total': new_total,
                'final_total': final_total,
                'match_status': match_status,
                'difference': difference
            }
        }
        
        return result
    
    def _merge_duplicates(self, data: List[Dict]) -> Dict[str, int]:
        """
        合并重复的80码，对数量求和
        
        参数：
            data: 原始数据列表
        
        返回：
            {80码: 数量} 的字典
        """
        merged = defaultdict(int)
        
        for item in data:
            code = item['code']
            quantity = item['quantity']
            merged[code] += quantity
        
        return dict(merged)
    
    def _arrange_by_standard(self, standard_codes: List[str], 
                            merged_data: Dict[str, int]) -> List[Dict]:
        """
        按标准顺序排列数据
        标准清单中有的80码保留，没有的填0，不存在的不加入
        
        参数：
            standard_codes: 标准80码清单
            merged_data: 合并后的数据
        
        返回：
            按顺序排列的结果列表
        """
        result = []
        
        for code in standard_codes:
            quantity = merged_data.get(code, 0)
            exists = code in merged_data
            
            result.append({
                'code': code,
                'quantity': quantity,
                'exists': exists
            })
        
        return result
    
    def _find_new_codes(self, standard_codes: List[str], 
                       merged_data: Dict[str, int]) -> List[Dict]:
        """
        找出新增的80码（标准清单中没有，但报表中存在的）
        
        参数：
            standard_codes: 标准80码清单
            merged_data: 合并后的数据
        
        返回：
            新增80码列表
        """
        standard_set = set(standard_codes)
        new_codes = []
        
        for code, quantity in merged_data.items():
            if code not in standard_set:
                new_codes.append({
                    'code': code,
                    'quantity': quantity,
                    'remark': '标准80码清单中不存在'
                })
        
        # 按80码数值排序新增80码（便于查看）
        try:
            new_codes.sort(key=lambda x: int(x['code']))
        except ValueError:
            # 如果无法转换为数字，则保持原顺序
            pass
        
        return new_codes
    
    def _format_with_sequence(self, data: List[Dict], exists_flag: str = None) -> List[Dict]:
        """
        添加序号，格式化结果
        
        参数：
            data: 数据列表
            exists_flag: 用于标准结果中"是否存在"字段
        
        返回：
            添加序号的结果列表
        """
        result = []
        
        for idx, item in enumerate(data, 1):
            formatted_item = {
                'sequence': idx,
                'code': item['code'],
                'quantity': item['quantity']
            }
            
            # 标准结果中添加"是否存在"
            if exists_flag is not None:
                formatted_item['exists'] = '是' if item.get('exists', False) else '否'
            
            # 新增80码中添加"备注"
            if 'remark' in item:
                formatted_item['remark'] = item['remark']
            
            result.append(formatted_item)
        
        return result

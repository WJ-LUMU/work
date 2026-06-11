# src/excel_reader.py
"""
Excel读取模块

负责：
1. 读取标准80码清单
2. 读取生产日报"汇总"工作表
3. 处理各种Excel格式问题
"""

import pandas as pd
from pathlib import Path
from typing import List, Dict, Tuple


class ExcelReader:
    """Excel文件读取类"""
    
    def __init__(self):
        """初始化读取器"""
        pass
    
    def read_standard_codes(self, file_path: Path) -> List[str]:
        """
        读取标准80码清单
        
        参数：
            file_path: 标准清单文件路径
        
        返回：
            按顺序排列的80码列表
        
        抛出：
            Exception: 当文件无法打开或格式不正确时
        """
        try:
            # 读取Excel文件，第一列作为80码
            df = pd.read_excel(file_path, header=None)
            
            if df.empty:
                raise Exception("标准80码清单文件为空")
            
            # 获取第一列数据
            codes = df.iloc[:, 0].tolist()
            
            # 清理数据：转换成文本、去空格、处理浮点数
            cleaned_codes = []
            for code in codes:
                # 跳过NaN值
                if pd.isna(code):
                    continue
                
                # 转换成字符串
                code_str = str(code).strip()
                
                # 处理浮点数格式（例如"80962510809.0"变成"80962510809"）
                if code_str.endswith('.0'):
                    code_str = code_str[:-2]
                
                # 去除末尾的浮点部分
                if '.' in code_str:
                    parts = code_str.split('.')
                    if len(parts[1]) == 0 or (len(parts[1]) <= 5 and parts[1].replace('0', '') == ''):
                        code_str = parts[0]
                
                # 再次去空格
                code_str = code_str.strip()
                
                if code_str:  # 非空则添加
                    cleaned_codes.append(code_str)
            
            if not cleaned_codes:
                raise Exception("找不到有效的80码数据")
            
            # 检查重复的80码
            if len(cleaned_codes) != len(set(cleaned_codes)):
                duplicates = [code for code in cleaned_codes if cleaned_codes.count(code) > 1]
                unique_duplicates = list(set(duplicates))
                print(f"\n⚠️  警告：标准清单中发现重复的80码：{unique_duplicates}")
                print(f"    将使用第一次出现的位置")
                # 保留第一次出现的位置，移除后续重复
                seen = set()
                cleaned_codes = [x for x in cleaned_codes if not (x in seen or seen.add(x))]
            
            return cleaned_codes
            
        except FileNotFoundError:
            raise Exception(f"找不到文件：{file_path}")
        except Exception as e:
            raise Exception(f"读取标准80码清单时出错：{str(e)}")
    
    def read_report_sheet(self, file_path: Path) -> List[Dict]:
        """
        读取生产日报"汇总"工作表
        
        参数：
            file_path: 生产日报文件路径
        
        返回：
            包含80码和数量的字典列表
        
        抛出：
            Exception: 当找不到"汇总"工作表或必要列时
        """
        try:
            # 首先获取所有工作表名称
            xls = pd.ExcelFile(file_path)
            sheet_names = xls.sheet_names
            
            # 查找"汇总"工作表
            summary_sheet = None
            for sheet in sheet_names:
                if sheet.strip() == "汇总":
                    summary_sheet = sheet
                    break
            
            if summary_sheet is None:
                available_sheets = ", ".join(sheet_names)
                raise Exception(
                    f"找不到名称为'汇总'的工作表\n"
                    f"可用的工作表有：{available_sheets}"
                )
            
            # 读取汇总工作表
            df = pd.read_excel(file_path, sheet_name=summary_sheet)
            
            if df.empty:
                raise Exception("'汇总'工作表中没有数据")
            
            # 查找80码列
            code_column = self._find_code_column(df)
            if code_column is None:
                columns = ", ".join(df.columns.tolist())
                raise Exception(
                    f"找不到80码列\n"
                    f"可用的列有：{columns}\n"
                    f"请确保包含'80码'、'编码'或'物料编码'列"
                )
            
            # 查找数量列
            quantity_column = self._find_quantity_column(df)
            if quantity_column is None:
                columns = ", ".join(df.columns.tolist())
                raise Exception(
                    f"找不到数量列\n"
                    f"可用的列有：{columns}\n"
                    f"请确保包含'合计完成'或'合计'列"
                )
            
            # 提取数据
            result = []
            invalid_rows = []
            
            for idx, row in df.iterrows():
                code = row[code_column]
                quantity = row[quantity_column]
                
                # 跳过NaN行
                if pd.isna(code):
                    continue
                
                # 清理80码
                code_str = str(code).strip()
                
                # 处理浮点数格式
                if code_str.endswith('.0'):
                    code_str = code_str[:-2]
                
                if '.' in code_str:
                    parts = code_str.split('.')
                    if len(parts[1]) == 0 or (len(parts[1]) <= 5 and parts[1].replace('0', '') == ''):
                        code_str = parts[0]
                
                code_str = code_str.strip()
                
                if not code_str:
                    continue
                
                # 清理数量
                try:
                    if pd.isna(quantity):
                        quantity_value = 0
                    else:
                        quantity_value = float(quantity)
                        # 转换为整数
                        quantity_value = int(quantity_value)
                except (ValueError, TypeError):
                    invalid_rows.append({
                        'code': code_str,
                        'quantity': quantity,
                        'row': idx + 2  # 加2是因为行号从0开始，还要加上表头
                    })
                    continue
                
                result.append({
                    'code': code_str,
                    'quantity': quantity_value
                })
            
            # 提示无效行
            if invalid_rows:
                print(f"\n⚠️  警告：发现 {len(invalid_rows)} 行数据无法转换为数字，已忽略：")
                for item in invalid_rows[:5]:  # 只显示前5个
                    print(f"   第{item['row']}行：80码={item['code']}, 数量={item['quantity']}")
                if len(invalid_rows) > 5:
                    print(f"   ...以及其他 {len(invalid_rows) - 5} 行")
            
            if not result:
                raise Exception("'汇总'工作表中没有找到有效的数据")
            
            return result
            
        except FileNotFoundError:
            raise Exception(f"找不到文件：{file_path}")
        except Exception as e:
            raise Exception(f"读取生产日报时出错：{str(e)}")
    
    def _find_code_column(self, df: pd.DataFrame) -> str:
        """
        查找80码列
        
        参数：
            df: 数据框
        
        返回：
            列名或None
        """
        keywords = ['80码', '编码', '物料编码', '物料', 'code', 'Code']
        
        for col in df.columns:
            col_lower = str(col).lower().strip()
            for keyword in keywords:
                if keyword.lower() in col_lower:
                    return col
        
        return None
    
    def _find_quantity_column(self, df: pd.DataFrame) -> str:
        """
        查找数量列
        优先查找"合计完成"，其次查找"合计"
        
        参数：
            df: 数据框
        
        返回：
            列名或None
        """
        # 优先查找的关键词（按优先级）
        keywords_priority = [
            ['合计完成', '完成数'],
            ['合计', 'total'],
            ['数量', 'qty', 'quantity']
        ]
        
        for priority_keywords in keywords_priority:
            for col in df.columns:
                col_str = str(col).strip()
                for keyword in priority_keywords:
                    if keyword in col_str:
                        return col
        
        return None

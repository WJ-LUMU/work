# main.py
"""
80码生产数据自动整理工具 - 主程序入口

这是程序的主入口，负责：
1. 获取用户输入的文件路径
2. 调用各模块处理数据
3. 显示处理结果
"""

import os
import sys
from pathlib import Path

# 导入自定义模块
from src.excel_reader import ExcelReader
from src.data_processor import DataProcessor
from src.excel_exporter import ExcelExporter


def print_welcome():
    """打印欢迎信息"""
    print("\n" + "="*60)
    print("           80码生产数据自动整理工具 v1.0")
    print("="*60 + "\n")


def get_file_path(prompt, file_type="Excel文件"):
    """
    获取用户输入的文件路径
    
    参数：
        prompt: 提示信息
        file_type: 文件类型说明
    
    返回：
        Path对象或None
    """
    while True:
        print(f"\n请输入{file_type}的完整路径（或直接拖入文件）：")
        user_input = input(f"> ").strip()
        
        # 处理拖入文件时的引号
        if user_input.startswith('"') and user_input.endswith('"'):
            user_input = user_input[1:-1]
        
        # 如果为空则提示
        if not user_input:
            print(f"❌ 错误：文件路径不能为空")
            continue
        
        # 转换为Path对象
        file_path = Path(user_input)
        
        # 检查文件是否存在
        if not file_path.exists():
            print(f"❌ 错误：找不到文件 '{file_path}'")
            print(f"   请检查路径是否正确")
            continue
        
        # 检查是否是Excel文件
        if not str(file_path).lower().endswith(('.xlsx', '.xls')):
            print(f"❌ 错误：请选择Excel文件（.xlsx 或 .xls）")
            continue
        
        print(f"✓ 已选择：{file_path}")
        return file_path


def get_save_path():
    """
    获取保存位置
    
    返回：
        Path对象
    """
    while True:
        print("\n请输入结果保存位置（文��夹路径）：")
        print("（如果直接回车，将保存到当前目录）")
        user_input = input("> ").strip()
        
        # 处理拖入文件夹时的引号
        if user_input.startswith('"') and user_input.endswith('"'):
            user_input = user_input[1:-1]
        
        # 如果为空，使用当前目录
        if not user_input:
            save_dir = Path.cwd()
            print(f"✓ 将保存到当前目录：{save_dir}")
            return save_dir
        
        save_dir = Path(user_input)
        
        # 如果不存在则创建
        if not save_dir.exists():
            try:
                save_dir.mkdir(parents=True, exist_ok=True)
                print(f"✓ 已创建保存目录：{save_dir}")
                return save_dir
            except Exception as e:
                print(f"❌ 错误：无法创建目录 '{save_dir}'")
                print(f"   错误原因：{str(e)}")
                continue
        
        if save_dir.is_dir():
            print(f"✓ 已选择保存目录：{save_dir}")
            return save_dir
        else:
            print(f"❌ 错误：'{save_dir}' 不是一个有效的文件夹")
            continue


def main():
    """主程序"""
    try:
        print_welcome()
        
        # 第1步：获取标准80码清单文件
        print("【步骤1】选择标准80码清单文件")
        print("-" * 60)
        standard_file = get_file_path("标准80码清单文件路径")
        if not standard_file:
            print("❌ 已取消操作")
            return
        
        # 第2步：获取生产日报文件
        print("\n【步骤2】选择生产日报Excel文件")
        print("-" * 60)
        report_file = get_file_path("生产日报文件路径")
        if not report_file:
            print("❌ 已取消操作")
            return
        
        # 第3步：获取保存位置
        print("\n【步骤3】选择结果保存位置")
        print("-" * 60)
        save_dir = get_save_path()
        
        # 第4步：开始处理
        print("\n【步骤4】开始处理数据")
        print("-" * 60)
        print("正在读取标准80码清单...")
        
        # 读取标准80码清单
        reader = ExcelReader()
        standard_codes = reader.read_standard_codes(standard_file)
        print(f"✓ 已读取标准80码 {len(standard_codes)} 个")
        
        # 读取生产日报
        print("正在读取生产日报...")
        report_data = reader.read_report_sheet(report_file)
        print(f"✓ 已读取生产日报数据 {len(report_data)} 行")
        
        # 处理数据
        print("正在处理数据...")
        processor = DataProcessor()
        result = processor.process(standard_codes, report_data)
        
        # 导出结果
        print("正在导出结果...")
        exporter = ExcelExporter()
        output_file = exporter.export(result, save_dir)
        
        # 显示结果
        print("\n" + "="*60)
        print("                    ✓ 处理完成")
        print("="*60)
        print(f"\n📊 数据统计：")
        print(f"   标准清单内总计：{result['summary']['standard_total']} 个")
        print(f"   新增80码总计：{result['summary']['new_total']} 个")
        print(f"   整理后总计：{result['summary']['final_total']} 个")
        print(f"   原始汇总表总计：{result['summary']['original_total']} 个")
        
        print(f"\n📋 核对结果：")
        if result['summary']['match_status'] == '核对一致':
            print(f"   ✓ {result['summary']['match_status']}")
        else:
            print(f"   ❌ {result['summary']['match_status']}")
            print(f"      数量差异：{result['summary']['difference']} 个")
        
        print(f"\n💾 结果已保存到：")
        print(f"   {output_file}")
        print("\n" + "="*60 + "\n")
        
    except Exception as e:
        print(f"\n❌ 发生错误：{str(e)}")
        print(f"\n请检查：")
        print(f"   1. 文件是否已被Excel或其他程序打开")
        print(f"   2. 文件格式是否正确")
        print(f"   3. 磁盘空间是否充足")
        sys.exit(1)


if __name__ == "__main__":
    main()

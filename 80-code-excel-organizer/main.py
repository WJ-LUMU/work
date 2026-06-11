# main.py
"""
全年生产日报自动汇总工具 - 主程序入口
"""

import sys
from pathlib import Path
from src.file_scanner import FileScanner
from src.excel_reader import ExcelReader
from src.data_processor import DataProcessor
from src.excel_exporter import ExcelExporter
from src.logger import Logger

def print_welcome():
    print("\n" + "="*70)
    print("             全年生产日报自动汇总工具 v2.0")
    print("="*70 + "\n")

def get_folder_path():
    while True:
        print("请输入存放全年日报的文件夹路径：")
        user_input = input("> ").strip()
        if user_input.startswith('"') and user_input.endswith('"'):
            user_input = user_input[1:-1]
        if not user_input:
            print("错误：路径不能为空\n")
            continue
        folder_path = Path(user_input)
        if not folder_path.exists():
            print("错误：找不到文件夹\n")
            continue
        if not folder_path.is_dir():
            print("错误：不是文件夹\n")
            continue
        print(f"已选择：{folder_path}\n")
        return folder_path

def get_standard_codes_file():
    while True:
        print("是否要使用固定的SKU清单文件？(y/n，默认n)")
        choice = input("> ").strip().lower()
        if choice in ['', 'n', 'no']:
            return None
        if choice in ['y', 'yes']:
            while True:
                print("请输入SKU清单文件路径：")
                user_input = input("> ").strip()
                if user_input.startswith('"') and user_input.endswith('"'):
                    user_input = user_input[1:-1]
                file_path = Path(user_input)
                if not file_path.exists():
                    print("找不到文件\n")
                    continue
                if not str(file_path).lower().endswith(('.xlsx', '.xls')):
                    print("请选择Excel文件\n")
                    continue
                print(f"已选择：{file_path}\n")
                return file_path

def get_save_path():
    print("请输入结果保存位置(默认当前目录)：")
    user_input = input("> ").strip()
    if user_input.startswith('"') and user_input.endswith('"'):
        user_input = user_input[1:-1]
    if not user_input:
        return Path.cwd()
    save_dir = Path(user_input)
    if not save_dir.exists():
        save_dir.mkdir(parents=True, exist_ok=True)
    return save_dir

def main():
    try:
        print_welcome()
        print("【步骤1】选择输入和输出位置\n" + "-" * 70)
        report_folder = get_folder_path()
        standard_codes_file = get_standard_codes_file()
        save_dir = get_save_path()
        
        print("【步骤2】初始化系统\n" + "-" * 70)
        logger = Logger()
        print("日志系统已初始化\n")
        
        print("【步骤3】扫描Excel文件\n" + "-" * 70)
        scanner = FileScanner(logger)
        excel_files = scanner.scan_folder(report_folder)
        print(f"找到 {len(excel_files)} 个Excel文件\n")
        
        if not excel_files:
            print("错误：文件夹中没有找到Excel文件\n")
            return
        
        print("【步骤4】读取标准SKU清单\n" + "-" * 70)
        standard_codes = []
        if standard_codes_file:
            reader = ExcelReader(logger)
            try:
                standard_codes = reader.read_standard_codes(standard_codes_file)
                print(f"已读取标准SKU清单，共 {len(standard_codes)} 个\n")
            except Exception as e:
                print(f"警告：无法读取SKU清单\n")
        
        print("【步骤5】处理Excel文件\n" + "-" * 70)
        reader = ExcelReader(logger)
        processor = DataProcessor(logger, standard_codes)
        
        for idx, file_path in enumerate(excel_files, 1):
            print(f"[{idx}/{len(excel_files)}] {file_path.name}")
            try:
                report_data = reader.read_report_file(file_path)
                if report_data is None:
                    continue
                processor.add_daily_report(file_path, report_data)
                print(f"  成功处理")
            except Exception as e:
                print(f"  错误：{str(e)}")
                logger.add_error(file_path.name, str(e))
        
        print("\n【步骤6】检测重复文件\n" + "-" * 70)
        processor.detect_duplicates()
        duplicates_count = len(processor.duplicates)
        print(f"检测完成，发现 {duplicates_count} 个疑似重复\n")
        
        print("【步骤7】生成汇总结果\n" + "-" * 70)
        result = processor.get_result()
        print(f"汇总完成\n")
        
        print("【步骤8】导出Excel文件\n" + "-" * 70)
        exporter = ExcelExporter(logger)
        output_file = exporter.export(result, save_dir, standard_codes, processor)
        print(f"文件已保存\n")
        
        print("\n" + "="*70 + "\n处理完成\n" + "="*70)
        print(f"\n数据统计：")
        print(f"  总文件数：{len(excel_files)}")
        print(f"  成功读取：{result['summary']['files_processed']}")
        print(f"  异常文件：{len(logger.errors)}")
        print(f"  疑似重复：{duplicates_count}")
        print(f"  汇总SKU数：{result['summary']['sku_count']}")
        print(f"  全年总数量：{result['summary']['total_quantity']}")
        print(f"  新增SKU：{result['summary']['new_sku_count']}")
        print(f"\n核对结果：{result['summary']['match_status']}")
        if result['summary']['match_status'] != '核对一致':
            print(f"  差异：{result['summary']['difference']}")
        print(f"\n输出文件：{output_file}\n")
        
    except KeyboardInterrupt:
        print("\n程序已被中断\n")
        sys.exit(0)
    except Exception as e:
        print(f"\n错误：{str(e)}\n")
        sys.exit(1)

if __name__ == "__main__":
    main()

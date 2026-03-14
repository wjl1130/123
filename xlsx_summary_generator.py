# -*- coding: utf-8 -*-
"""
Excel文件金额汇总脚本

功能：自动遍历当前目录下所有 .xlsx 文件，读取每个文件的'Sheet1'工作表，
      提取'金额'列的总和，并将结果汇总写入 summary_report.xlsx 文件。

依赖：pandas, openpyxl
安装：pip install pandas openpyxl

兼容性：Python 3.8+
"""

import os
import sys
from pathlib import Path
from typing import List, Tuple

import pandas as pd


def find_xlsx_files(directory: str = ".") -> List[str]:
    """
    查找指定目录下所有的 .xlsx 文件（排除汇总结果文件）

    Args:
        directory: 要搜索的目录路径，默认为当前目录

    Returns:
        找到的 .xlsx 文件路径列表
    """
    xlsx_files = []
    try:
        for file in os.listdir(directory):
            if file.endswith(".xlsx") and file != "summary_report.xlsx":
                full_path = os.path.join(directory, file)
                if os.path.isfile(full_path):
                    xlsx_files.append(full_path)
    except Exception as e:
        print(f"❌ 扫描目录时出错: {str(e)}")
        sys.exit(1)
    
    return xlsx_files


def read_amount_sum(file_path: str, sheet_name: str = "Sheet1", column_name: str = "金额") -> Tuple[str, float]:
    """
    读取单个Excel文件，计算'金额'列的总和

    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称，默认为'Sheet1'
        column_name: 要计算的列名，默认为'金额'

    Returns:
        (文件名, 金额总和) 元组

    Raises:
        FileNotFoundError: 文件不存在
        ValueError: 工作表或列不存在
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        if column_name not in df.columns:
            raise ValueError(f"未找到'{column_name}'列")
        
        amount_sum = df[column_name].sum()
        filename = os.path.basename(file_path)
        
        return (filename, amount_sum)
    
    except FileNotFoundError:
        print(f"❌ 文件不存在: {file_path}")
        raise
    except ValueError as ve:
        print(f"❌ 读取文件 '{os.path.basename(file_path)}' 时出错: {str(ve)}")
        raise
    except Exception as e:
        print(f"❌ 处理文件 '{os.path.basename(file_path)}' 时发生未知错误: {str(e)}")
        raise


def generate_summary_report(results: List[Tuple[str, float]], output_file: str = "summary_report.xlsx") -> None:
    """
    生成汇总报告Excel文件

    Args:
        results: 包含(文件名, 金额总和)的列表
        output_file: 输出文件路径，默认为'summary_report.xlsx'
    """
    try:
        df_summary = pd.DataFrame(results, columns=["文件名", "总金额"])
        
        df_summary.to_excel(output_file, index=False, sheet_name="汇总报告")
        print(f"\n✅ 汇总报告已生成: {output_file}")
        print(f"   共处理了 {len(results)} 个Excel文件")
        
    except PermissionError:
        print(f"❌ 无法写入文件 '{output_file}'，文件可能被其他程序占用")
        raise
    except Exception as e:
        print(f"❌ 生成汇总报告时出错: {str(e)}")
        raise


def print_summary(results: List[Tuple[str, float]]) -> None:
    """
    打印汇总结果到控制台

    Args:
        results: 包含(文件名, 金额总和)的列表
    """
    print("\n" + "="*50)
    print("💰 Excel文件金额汇总结果")
    print("="*50)
    
    if not results:
        print("   未找到任何 .xlsx 文件或处理结果为空")
        return
    
    total_sum = 0.0
    for filename, amount in results:
        print(f"   {filename}: ¥{amount:,.2f}")
        total_sum += amount
    
    print("-"*50)
    print(f"   所有文件总金额: ¥{total_sum:,.2f}")
    print("="*50)


def main(directory: str = ".") -> None:
    """
    主函数：执行完整的汇总流程

    Args:
        directory: 要处理的目录路径，默认为当前目录
    """
    print(f"📂 正在扫描目录: {os.path.abspath(directory)}")
    
    xlsx_files = find_xlsx_files(directory)
    
    if not xlsx_files:
        print("❌ 未找到任何 .xlsx 文件")
        sys.exit(0)
    
    print(f"✅ 找到 {len(xlsx_files)} 个Excel文件，开始处理...")
    
    results = []
    for file_path in xlsx_files:
        try:
            filename, amount_sum = read_amount_sum(file_path)
            results.append((filename, amount_sum))
            print(f"   ✓ 已处理: {filename}")
        except Exception:
            continue
    
    if results:
        generate_summary_report(results)
        print_summary(results)
    else:
        print("\n❌ 没有成功处理任何文件")
        sys.exit(1)


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Excel文件金额汇总工具 - 遍历目录下所有.xlsx文件，提取'金额'列总和并生成汇总报告"
    )
    parser.add_argument(
        "-d", "--directory",
        type=str,
        default=".",
        help="要扫描的目录路径，默认为当前目录"
    )
    
    args = parser.parse_args()
    
    main(args.directory)

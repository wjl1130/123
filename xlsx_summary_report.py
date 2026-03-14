# -*- coding: utf-8 -*-
"""
Excel 文件金额汇总脚本
功能：自动遍历当前目录下所有 .xlsx 文件，读取每个文件的 'Sheet1' 工作表，
      提取 '金额' 列的总和，并将文件名和总金额汇总写入 summary_report.xlsx
"""

import os
import pandas as pd
from pathlib import Path


def find_xlsx_files(directory):
    """
    查找指定目录下的所有 .xlsx 文件
    
    参数:
        directory: 要搜索的目录路径
    
    返回:
        xlsx 文件路径列表
    """
    xlsx_files = []
    # 遍历目录中的所有文件
    for filename in os.listdir(directory):
        # 检查文件是否以 .xlsx 结尾（排除临时文件 ~$ 开头）
        if filename.endswith('.xlsx') and not filename.startswith('~$'):
            xlsx_files.append(os.path.join(directory, filename))
    return xlsx_files


def extract_amount_sum(file_path):
    """
    从 Excel 文件中提取 '金额' 列的总和
    
    参数:
        file_path: Excel 文件的完整路径
    
    返回:
        (文件名, 金额总和) 的元组，如果出错则返回 (文件名, None)
    """
    try:
        # 读取 Excel 文件的 Sheet1 工作表
        df = pd.read_excel(file_path, sheet_name='Sheet1')
        
        # 检查是否存在 '金额' 列
        if '金额' not in df.columns:
            print(f"警告: 文件 '{os.path.basename(file_path)}' 中没有找到 '金额' 列")
            return os.path.basename(file_path), None
        
        # 提取金额列并计算总和
        # 先删除空值，再求和
        amount_sum = df['金额'].dropna().sum()
        
        return os.path.basename(file_path), amount_sum
        
    except Exception as e:
        print(f"错误: 处理文件 '{os.path.basename(file_path)}' 时出错: {str(e)}")
        return os.path.basename(file_path), None


def create_summary_report(results, output_path):
    """
    创建汇总报告 Excel 文件
    
    参数:
        results: 包含 (文件名, 金额总和) 元组的列表
        output_path: 输出文件的完整路径
    """
    # 创建 DataFrame
    df = pd.DataFrame(results, columns=['文件名', '总金额'])
    
    # 添加统计信息
    total_files = len(results)
    valid_files = sum(1 for _, amount in results if amount is not None)
    total_amount = sum(amount for _, amount in results if amount is not None)
    
    # 使用 ExcelWriter 创建更复杂的 Excel 文件
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 写入主数据表
        df.to_excel(writer, sheet_name='汇总数据', index=False)
        
        # 创建统计信息表
        stats_data = {
            '统计项': ['文件总数', '成功处理文件数', '所有文件金额总和'],
            '数值': [total_files, valid_files, total_amount]
        }
        stats_df = pd.DataFrame(stats_data)
        stats_df.to_excel(writer, sheet_name='统计信息', index=False)
    
    print(f"\n汇总报告已生成: {output_path}")
    print(f"统计信息:")
    print(f"  - 文件总数: {total_files}")
    print(f"  - 成功处理: {valid_files}")
    print(f"  - 金额总和: {total_amount:,.2f}")


def main():
    """
    主函数：协调整个处理流程
    """
    # 获取当前脚本所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    print("=" * 60)
    print("Excel 文件金额汇总工具")
    print("=" * 60)
    print(f"工作目录: {current_dir}")
    print()
    
    # 查找所有 .xlsx 文件
    print("正在查找 .xlsx 文件...")
    xlsx_files = find_xlsx_files(current_dir)
    
    # 排除汇总报告文件本身
    xlsx_files = [f for f in xlsx_files if 'summary_report' not in os.path.basename(f).lower()]
    
    if not xlsx_files:
        print("未找到任何 .xlsx 文件！")
        return
    
    print(f"找到 {len(xlsx_files)} 个 .xlsx 文件:")
    for f in xlsx_files:
        print(f"  - {os.path.basename(f)}")
    print()
    
    # 处理每个文件
    print("正在处理文件...")
    results = []
    for file_path in xlsx_files:
        filename, amount_sum = extract_amount_sum(file_path)
        if amount_sum is not None:
            print(f"  ✓ {filename}: {amount_sum:,.2f}")
            results.append((filename, amount_sum))
        else:
            print(f"  ✗ {filename}: 处理失败")
            results.append((filename, 0))  # 失败的文件金额记为 0
    
    print()
    
    # 生成汇总报告
    output_file = os.path.join(current_dir, 'summary_report.xlsx')
    create_summary_report(results, output_file)
    
    print("\n处理完成!")
    print("=" * 60)


if __name__ == '__main__':
    main()

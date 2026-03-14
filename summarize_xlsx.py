import os
import pandas as pd
from pathlib import Path

def summarize_xlsx_files():
    """
    遍历当前目录下所有.xlsx文件，提取'金额'列总和并生成汇总报告
    
    功能说明：
    1. 自动扫描当前工作目录下的所有.xlsx文件
    2. 读取每个文件的'Sheet1'工作表
    3. 计算'金额'列的总和
    4. 将结果汇总写入summary_report.xlsx文件
    """
    
    # 获取当前工作目录
    current_dir = Path.cwd()
    
    # 查找当前目录下所有的.xlsx文件（排除输出文件本身）
    xlsx_files = [f for f in current_dir.glob("*.xlsx") 
                  if f.name != "summary_report.xlsx"]
    
    # 如果没有找到任何xlsx文件，提示用户并退出
    if not xlsx_files:
        print("当前目录下没有找到任何.xlsx文件（除summary_report.xlsx外）")
        return
    
    # 用于存储汇总结果的列表
    summary_data = []
    
    # 遍历每个xlsx文件
    for file_path in xlsx_files:
        try:
            # 读取Excel文件的Sheet1工作表
            df = pd.read_excel(file_path, sheet_name="Sheet1")
            
            # 检查是否存在'金额'列
            if "金额" not in df.columns:
                print(f"警告：文件 '{file_path.name}' 中没有找到'金额'列，已跳过")
                continue
            
            # 计算'金额'列的总和
            # 使用pd.to_numeric确保数据为数值类型，errors='coerce'会将非数值转为NaN
            amount_series = pd.to_numeric(df["金额"], errors="coerce")
            total_amount = amount_series.sum()
            
            # 将文件名和总金额添加到汇总列表
            summary_data.append({
                "文件名": file_path.name,
                "金额总和": total_amount
            })
            
            print(f"已处理：{file_path.name}，金额总和：{total_amount}")
            
        except Exception as e:
            # 捕获并报告读取文件时的错误，继续处理下一个文件
            print(f"错误：读取文件 '{file_path.name}' 时出错：{str(e)}")
            continue
    
    # 如果没有成功处理任何文件，提示用户并退出
    if not summary_data:
        print("没有成功处理任何文件，无法生成汇总报告")
        return
    
    # 创建汇总DataFrame
    summary_df = pd.DataFrame(summary_data)
    
    # 添加总计行
    total_row = pd.DataFrame([{
        "文件名": "总计",
        "金额总和": summary_df["金额总和"].sum()
    }])
    summary_df = pd.concat([summary_df, total_row], ignore_index=True)
    
    # 定义输出文件路径
    output_path = current_dir / "summary_report.xlsx"
    
    # 将汇总结果写入Excel文件
    # index=False 表示不写入行索引
    summary_df.to_excel(output_path, index=False, sheet_name="汇总报告")
    
    print(f"\n汇总报告已生成：{output_path}")
    print(f"共处理 {len(summary_data)} 个文件")


# 程序入口
if __name__ == "__main__":
    summarize_xlsx_files()

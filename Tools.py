import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

def find_optimal_cutoff(data_df):
    # 提取需要的列
    df = data_df[['Gene_Mfg_ID', 'file_reads', 'flagstat']]

    # 初始化变量
    max_below_90_count = 0
    optimal_cutoff = None

    # 设置file_reads的cutoff范围，从100到500，并以20为单位增加
    for cutoff in range(100, 501, 20):
        below_90_count = df[df['file_reads'] >= cutoff]['flagstat'].lt(90).sum()

        # 更新最大数量和对应的cutoff
        if below_90_count > max_below_90_count:
            max_below_90_count = below_90_count
            optimal_cutoff = cutoff
        elif below_90_count == max_below_90_count and (optimal_cutoff is None or cutoff < optimal_cutoff):
            optimal_cutoff = cutoff

    return optimal_cutoff, max_below_90_count

# 示例用法
def main():
    # 使用tkinter选择文件
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
    
    if not file_path:
        print("未选择文件")
        return

    # 读取完整数据文件
    data_df = pd.read_excel(file_path, engine='xlrd')

    # 找到最优cutoff
    optimal_cutoff, max_below_90_count = find_optimal_cutoff(data_df)
    print(f"Optimal cutoff: {optimal_cutoff}, Max below 90 count: {max_below_90_count}")

if __name__ == "__main__":
    main()
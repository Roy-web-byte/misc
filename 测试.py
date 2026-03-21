import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def select_file(title):
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx *.xls")])
    root.destroy()
    return file_path

def main():
    print("--- 启动表格自动填充脚本 ---")

    # 1. 选择文件
    target_path = select_file("第一步：选择【空表格】（模板）")
    if not target_path:
        return

    source_path = select_file("第二步：选择【数据来源表格】")
    if not source_path:
        return

    try:
        # 2. 读取数据
        # 假设第一列是行标题 (index_col=0)
        target_df = pd.read_excel(target_path, index_col=0)
        source_df = pd.read_excel(source_path, index_col=0)

        print("\n正在匹配数据...")

        # 3. 核心逻辑：对齐填充
        # reindex 会根据 target_df 的行名和列名，从 source_df 中提取对应值
        # 如果 source 中不存在对应位置，会填入 NaN (空)
        filled_df = source_df.reindex(index=target_df.index, columns=target_df.columns)

        # 4. 保存结果
        # 我们直接覆盖原来的空表格，或者你可以改名保存
        filled_df.to_excel(target_path)
        
        messagebox.showinfo("成功", f"数据已填充完毕！\n保存路径：{target_path}")
        print(f"成功！已更新: {target_path}")

    except Exception as e:
        messagebox.showerror("错误", f"处理过程中出现问题：\n{str(e)}")
        print(f"错误详情: {e}")

if __name__ == "__main__":
    main()

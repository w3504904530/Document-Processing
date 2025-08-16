#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件对比工具 - GUI界面
支持CSV、Excel文件的智能对比和高亮功能
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
import threading
from datetime import datetime

# 导入我们的文件处理模块
import module.files
from TableComparison import merge_and_reorder, highlight_differences, save_to_excel


class FileCompareGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("文件对比工具 v1.0")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # 文件路径变量
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.comparison_column = tk.StringVar(value="A")
        
        # 策略变量
        self.preserve_order_by = tk.StringVar(value="None")
        self.column_sort_strategy = tk.StringVar(value="alternating")
        
        # 创建界面
        self.create_widgets()
        
        # 设置默认输出路径
        self.output_path.set("./data/对比结果.xlsx")
        
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="文件对比工具", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        # 文件1
        ttk.Label(file_frame, text="文件1:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        ttk.Entry(file_frame, textvariable=self.file1_path, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(file_frame, text="浏览", command=self.browse_file1).grid(row=0, column=2)
        
        # 文件2
        ttk.Label(file_frame, text="文件2:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        ttk.Entry(file_frame, textvariable=self.file2_path, width=50).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 5), pady=(5, 0))
        ttk.Button(file_frame, text="浏览", command=self.browse_file2).grid(row=1, column=2, pady=(5, 0))
        
        # 输出文件
        ttk.Label(file_frame, text="输出文件:").grid(row=2, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        ttk.Entry(file_frame, textvariable=self.output_path, width=50).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(0, 5), pady=(5, 0))
        ttk.Button(file_frame, text="浏览", command=self.browse_output).grid(row=2, column=2, pady=(5, 0))
        
        # 参数设置区域
        param_frame = ttk.LabelFrame(main_frame, text="对比参数", padding="10")
        param_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 配置参数区域的列权重
        param_frame.columnconfigure(1, weight=1)
        param_frame.columnconfigure(3, weight=1)
        param_frame.columnconfigure(5, weight=1)
        
        # 第一行：比较列名和行顺序策略
        ttk.Label(param_frame, text="比较列名:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        ttk.Entry(param_frame, textvariable=self.comparison_column, width=15).grid(row=0, column=1, sticky=tk.W, padx=(0, 20))
        
        ttk.Label(param_frame, text="行顺序策略:").grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
        order_combo = ttk.Combobox(param_frame, textvariable=self.preserve_order_by, 
                                  values=["None", "df1", "df2"], state="readonly", width=12)
        order_combo.grid(row=0, column=3, sticky=tk.W)
        
        # 第二行：列排序策略
        ttk.Label(param_frame, text="列排序策略:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(10, 0))
        sort_combo = ttk.Combobox(param_frame, textvariable=self.column_sort_strategy,
                                 values=["alternating", "grouped", "alphabetical"], state="readonly", width=15)
        sort_combo.grid(row=1, column=1, sticky=tk.W, pady=(10, 0))
        
        # 第三行：说明文字
        order_info = ttk.Label(param_frame, text="行顺序: None=按比较列排序 | df1=保留文件1顺序 | df2=保留文件2顺序", 
                               font=("Arial", 9), foreground="gray")
        order_info.grid(row=2, column=0, columnspan=4, sticky=tk.W, pady=(5, 0))
        
        strategy_info = ttk.Label(param_frame, text="列排序: alternating=交替排列 | grouped=分组排列 | alphabetical=字母顺序", 
                                 font=("Arial", 9), foreground="gray")
        strategy_info.grid(row=3, column=0, columnspan=4, sticky=tk.W, pady=(2, 0))
        
        # 操作按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=(0, 10))
        
        # 开始对比按钮
        self.compare_button = ttk.Button(button_frame, text="开始对比", command=self.start_comparison, 
                                        style="Accent.TButton")
        self.compare_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # 清空日志按钮
        ttk.Button(button_frame, text="清空日志", command=self.clear_log).pack(side=tk.LEFT, padx=(0, 10))
        
        # 打开输出文件夹按钮
        ttk.Button(button_frame, text="打开输出文件夹", command=self.open_output_folder).pack(side=tk.LEFT)
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="处理日志", padding="5")
        log_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # 日志文本框
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, width=80)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 进度条
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E))
        
    def browse_file1(self):
        filename = filedialog.askopenfilename(
            title="选择第一个文件",
            filetypes=[("所有支持的文件", "*.csv;*.xlsx;*.xls"), 
                      ("CSV文件", "*.csv"), 
                      ("Excel文件", "*.xlsx;*.xls")]
        )
        if filename:
            self.file1_path.set(filename)
            self.log_message(f"已选择文件1: {filename}")
            
    def browse_file2(self):
        filename = filedialog.askopenfilename(
            title="选择第二个文件",
            filetypes=[("所有支持的文件", "*.csv;*.xlsx;*.xls"), 
                      ("CSV文件", "*.csv"), 
                      ("Excel文件", "*.xlsx;*.xls")]
        )
        if filename:
            self.file2_path.set(filename)
            self.log_message(f"已选择文件2: {filename}")
            
    def browse_output(self):
        filename = filedialog.asksaveasfilename(
            title="选择输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("CSV文件", "*.csv")]
        )
        if filename:
            self.output_path.set(filename)
            self.log_message(f"输出文件设置为: {filename}")
            
    def log_message(self, message):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def clear_log(self):
        """清空日志"""
        self.log_text.delete(1.0, tk.END)
        
    def open_output_folder(self):
        """打开输出文件夹"""
        output_path = self.output_path.get()
        if not output_path:
            messagebox.showwarning("警告", "请先设置输出文件路径")
            return
            
        output_dir = os.path.dirname(output_path)
        if not output_dir:  # 如果是相对路径或空路径
            output_dir = os.getcwd()  # 使用当前工作目录
        
        # 转换为绝对路径，解决Windows相对路径问题
        output_dir = os.path.abspath(output_dir)
        
        # 确保输出目录存在
        try:
            os.makedirs(output_dir, exist_ok=True)
        except Exception as e:
            messagebox.showwarning("警告", f"无法创建输出目录: {str(e)}")
            return
        
        # 尝试打开文件夹
        try:
            if os.name == 'nt':  # Windows
                os.startfile(output_dir)
            else:  # Linux/Mac
                import subprocess
                subprocess.run(['xdg-open', output_dir])
        except Exception as e:
            messagebox.showwarning("警告", f"无法打开文件夹: {str(e)}\n文件夹路径: {output_dir}")
            
    def validate_inputs(self):
        """验证输入参数"""
        if not self.file1_path.get():
            messagebox.showerror("错误", "请选择第一个文件")
            return False
            
        if not self.file2_path.get():
            messagebox.showerror("错误", "请选择第二个文件")
            return False
            
        if not self.comparison_column.get():
            messagebox.showerror("错误", "请输入比较列名")
            return False
            
        if not self.output_path.get():
            messagebox.showerror("错误", "请设置输出文件路径")
            return False
            
        return True
        
    def start_comparison(self):
        """开始对比处理"""
        if not self.validate_inputs():
            return
            
        # 禁用按钮，显示进度条
        self.compare_button.config(state="disabled")
        self.progress.start()
        self.status_var.set("正在处理...")
        
        # 在新线程中执行对比，避免界面冻结
        thread = threading.Thread(target=self.run_comparison)
        thread.daemon = True
        thread.start()
        
    def run_comparison(self):
        """执行对比操作"""
        try:
            self.log_message("开始文件对比处理...")
            
            # 读取文件
            self.log_message("正在读取文件1...")
            df1 = module.files.read_file(self.file1_path.get())
            if df1 is None:
                raise Exception("文件1读取失败")
                
            self.log_message("正在读取文件2...")
            df2 = module.files.read_file(self.file2_path.get())
            if df2 is None:
                raise Exception("文件2读取失败")
                
            self.log_message(f"文件1列名: {list(df1.columns)}")
            self.log_message(f"文件2列名: {list(df2.columns)}")
            
            # 检查比较列是否存在
            comparison_col = self.comparison_column.get()
            if comparison_col not in df1.columns:
                raise Exception(f"比较列 '{comparison_col}' 在文件1中不存在")
            if comparison_col not in df2.columns:
                raise Exception(f"比较列 '{comparison_col}' 在文件2中不存在")
                
            # 执行合并和重排序
            preserve_order = self.preserve_order_by.get() if self.preserve_order_by.get() != "None" else None
            self.log_message(f"使用列排序策略: {self.column_sort_strategy.get()}")
            self.log_message(f"行顺序保留策略: {preserve_order if preserve_order else 'None (按比较列排序)'}")
            merged_df, column_pairs = merge_and_reorder(
                df1, df2, comparison_col, 
                preserve_order,
                self.column_sort_strategy.get()
            )
            
            # 保存结果
            self.log_message("正在保存结果...")
            output_path = self.output_path.get()
            
            # 转换为绝对路径，解决Windows相对路径问题
            output_path = os.path.abspath(output_path)
            
            # 确保输出目录存在
            output_dir = os.path.dirname(output_path)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
                self.log_message(f"确保输出目录存在: {output_dir}")
            
            save_to_excel(merged_df, output_path)
            
            # 执行高亮
            if column_pairs:
                self.log_message("正在执行高亮处理...")
                highlight_differences(output_path, column_pairs)  # 使用转换后的绝对路径
                self.log_message(f"找到 {len(column_pairs)} 对可对比的列")
            else:
                self.log_message("未找到可高亮的成对列")
            
            # 添加索引列说明
            if preserve_order == 'df1':
                self.log_message("已添加文件1的原始行索引列 (_df1_original_index)")
            elif preserve_order == 'df2':
                self.log_message("已添加文件2的原始行索引列 (_df2_original_index)")
            elif preserve_order is None:
                self.log_message("已添加两个文件的原始行索引列 (_df1_original_index, _df2_original_index)")
                
            self.log_message("处理完成！")
            self.status_var.set("处理完成")
            
            # 显示成功消息
            self.root.after(0, lambda: messagebox.showinfo("成功", "文件对比处理完成！"))
            
        except Exception as e:
            error_msg = f"处理过程中出现错误: {str(e)}"
            self.log_message(error_msg)
            self.status_var.set("处理失败")
            self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
            
        finally:
            # 恢复界面状态
            self.root.after(0, self.restore_ui)
            
    def restore_ui(self):
        """恢复界面状态"""
        self.compare_button.config(state="normal")
        self.progress.stop()
        self.status_var.set("就绪")


def main():
    root = tk.Tk()
    app = FileCompareGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

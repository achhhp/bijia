#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
供应商比价软件
功能：
1. 上传至少3份报价单（Excel或CSV格式）
2. 自动找出每个物料的最低价和对应供应商
3. 统计每个供应商中了哪些低价物料
4. 展示结果并导出分析报告
"""

import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import numpy as np

class VendorPriceComparison:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("供应商比价软件")
        self.root.geometry("1000x700")
        self.root.resizable(True, True)
        
        # 设置界面风格
        style = ttk.Style()
        style.theme_use('clam')
        
        # 初始化变量
        self.files = []
        self.data = {}
        self.analysis_result = None
        self.vendor_stats = None
        
        # 创建主界面
        self.create_main_window()
    
    def create_main_window(self):
        # 创建顶部工具栏
        toolbar = ttk.Frame(self.root, padding="10")
        toolbar.pack(fill=tk.X, side=tk.TOP)
        
        # 上传文件按钮
        upload_btn = ttk.Button(toolbar, text="上传报价单", command=self.upload_files)
        upload_btn.pack(side=tk.LEFT, padx=5)
        
        # 分析按钮
        analyze_btn = ttk.Button(toolbar, text="开始分析", command=self.analyze_prices)
        analyze_btn.pack(side=tk.LEFT, padx=5)
        
        # 导出按钮
        export_btn = ttk.Button(toolbar, text="导出报告", command=self.export_report)
        export_btn.pack(side=tk.LEFT, padx=5)
        
        # 清空按钮
        clear_btn = ttk.Button(toolbar, text="清空", command=self.clear_all)
        clear_btn.pack(side=tk.LEFT, padx=5)
        
        # 创建文件列表框
        file_frame = ttk.LabelFrame(self.root, text="已上传的报价单", padding="10")
        file_frame.pack(fill=tk.BOTH, expand=True, side=tk.TOP, padx=10, pady=5)
        
        self.file_tree = ttk.Treeview(file_frame, columns=("file",), show="headings")
        self.file_tree.heading("file", text="文件名")
        self.file_tree.column("file", width=400)
        self.file_tree.pack(fill=tk.BOTH, expand=True)
        
        # 创建结果展示区
        result_frame = ttk.LabelFrame(self.root, text="分析结果", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, side=tk.BOTTOM, padx=10, pady=5)
        
        # 创建结果标签页
        notebook = ttk.Notebook(result_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # 最低价分析标签页
        price_tab = ttk.Frame(notebook)
        notebook.add(price_tab, text="最低价分析")
        
        self.price_tree = ttk.Treeview(price_tab, columns=("item", "price", "vendor"), show="headings")
        self.price_tree.heading("item", text="物料名称")
        self.price_tree.heading("price", text="最低价")
        self.price_tree.heading("vendor", text="供应商")
        self.price_tree.column("item", width=300)
        self.price_tree.column("price", width=100)
        self.price_tree.column("vendor", width=200)
        
        price_scroll = ttk.Scrollbar(price_tab, orient=tk.VERTICAL, command=self.price_tree.yview)
        self.price_tree.configure(yscroll=price_scroll.set)
        price_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.price_tree.pack(fill=tk.BOTH, expand=True)
        
        # 供应商中标统计标签页
        vendor_tab = ttk.Frame(notebook)
        notebook.add(vendor_tab, text="供应商中标统计")
        
        self.vendor_tree = ttk.Treeview(vendor_tab, columns=("vendor", "count", "items"), show="headings")
        self.vendor_tree.heading("vendor", text="供应商")
        self.vendor_tree.heading("count", text="中标数量")
        self.vendor_tree.heading("items", text="中标物料")
        self.vendor_tree.column("vendor", width=200)
        self.vendor_tree.column("count", width=100)
        self.vendor_tree.column("items", width=500)
        
        vendor_scroll = ttk.Scrollbar(vendor_tab, orient=tk.VERTICAL, command=self.vendor_tree.yview)
        self.vendor_tree.configure(yscroll=vendor_scroll.set)
        vendor_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.vendor_tree.pack(fill=tk.BOTH, expand=True)
    
    def upload_files(self):
        """上传报价单文件"""
        filetypes = [
            ("Excel文件", "*.xlsx *.xls"),
            ("CSV文件", "*.csv"),
            ("所有文件", "*.*")
        ]
        
        new_files = filedialog.askopenfilenames(
            title="选择报价单文件",
            filetypes=filetypes,
            initialdir=os.getcwd()
        )
        
        if new_files:
            for file_path in new_files:
                if file_path not in self.files:
                    self.files.append(file_path)
                    filename = os.path.basename(file_path)
                    self.file_tree.insert("", tk.END, values=(filename,))
            
            if len(self.files) >= 3:
                messagebox.showinfo("提示", f"已上传 {len(self.files)} 份报价单，可以开始分析")
            else:
                messagebox.showinfo("提示", f"已上传 {len(self.files)} 份报价单，请至少上传3份")
    
    def parse_file(self, file_path):
        """解析报价单文件"""
        try:
            filename = os.path.basename(file_path)
            vendor_name = filename.split('.')[0]  # 用文件名作为供应商名称
            
            # 尝试读取Excel文件的所有工作表
            if file_path.endswith('.csv'):
                # CSV文件直接读取
                df = pd.read_csv(file_path)
                result, error = self.process_dataframe(df, filename)
                if error:
                    messagebox.showerror("错误", error)
                    return None
                return vendor_name, result
            else:
                # Excel文件，尝试读取所有工作表
                try:
                    # 首先尝试读取第一个工作表
                    df = pd.read_excel(file_path)
                    result, error = self.process_dataframe(df, filename)
                    if not error:
                        return vendor_name, result
                    
                    # 如果第一个工作表失败，尝试读取所有工作表
                    xl = pd.ExcelFile(file_path)
                    for sheet_name in xl.sheet_names:
                        df = pd.read_excel(xl, sheet_name)
                        result, error = self.process_dataframe(df, filename)
                        if not error:
                            return vendor_name, result
                    
                    # 所有工作表都失败
                    messagebox.showerror("错误", f"文件 {filename} 的所有工作表都缺少必要的列：品名/名称 或 单项报价/价格")
                    return None
                except Exception as e:
                    messagebox.showerror("错误", f"读取Excel文件 {file_path} 时出错：{str(e)}")
                    return None
        except Exception as e:
            messagebox.showerror("错误", f"解析文件 {file_path} 时出错：{str(e)}")
            return None
    
    def process_dataframe(self, df, filename):
        """处理数据框，查找必要的列"""
        # 标准化列名（去除空格、特殊字符，转换为小写）
        normalized_columns = {}
        for col in df.columns:
            # 标准化列名：去除空格、特殊字符，转换为小写
            normalized_col = ''.join(e for e in str(col) if e.isalnum() or e == '_').lower()
            normalized_columns[normalized_col] = col
        
        # 查找物料列（支持多种可能的列名）
        item_col = None
        item_col_options = ['品名', '名称', '物料', '货品', '商品']
        for option in item_col_options:
            normalized_option = ''.join(e for e in option if e.isalnum() or e == '_').lower()
            if normalized_option in normalized_columns:
                item_col = normalized_columns[normalized_option]
                break
        
        # 查找价格列（支持多种可能的列名）
        price_col = None
        price_col_options = ['单项报价', '报价', '价格', '单价', '单位价格']
        for option in price_col_options:
            normalized_option = ''.join(e for e in option if e.isalnum() or e == '_').lower()
            if normalized_option in normalized_columns:
                price_col = normalized_columns[normalized_option]
                break
        
        # 检查是否找到必要的列
        if not item_col or not price_col:
            return None, f"文件 {filename} 缺少必要的列：品名/名称 或 单项报价/价格"
        
        # 清理数据
        df = df.dropna(subset=[item_col, price_col])
        df['物料'] = df[item_col].astype(str).str.strip()  # 重命名为'物料'，与后续代码保持一致
        df['价格'] = pd.to_numeric(df[price_col], errors='coerce')  # 重命名为'价格'，与后续代码保持一致
        df = df.dropna(subset=['价格'])
        
        return df, None
    
    def analyze_prices(self):
        """分析价格，找出最低价"""
        if len(self.files) < 3:
            messagebox.showwarning("警告", "请至少上传3份报价单")
            return
        
        # 解析所有文件
        self.data = {}
        for file_path in self.files:
            result = self.parse_file(file_path)
            if result:
                vendor_name, df = result
                self.data[vendor_name] = df
        
        if not self.data:
            messagebox.showerror("错误", "没有成功解析的报价单")
            return
        
        # 收集所有物料
        all_items = set()
        for vendor, df in self.data.items():
            all_items.update(df['物料'].tolist())
        
        # 分析每个物料的最低价
        analysis_data = []
        for item in all_items:
            min_price = float('inf')
            min_vendor = ""
            
            for vendor, df in self.data.items():
                item_prices = df[df['物料'] == item]['价格']
                if not item_prices.empty:
                    price = item_prices.iloc[0]
                    if price < min_price:
                        min_price = price
                        min_vendor = vendor
            
            if min_vendor:
                analysis_data.append({
                    '物料名称': item,
                    '最低价': min_price,
                    '供应商': min_vendor
                })
        
        # 生成分析结果
        self.analysis_result = pd.DataFrame(analysis_data)
        
        # 统计每个供应商的中标情况
        self.vendor_stats = {}
        for _, row in self.analysis_result.iterrows():
            vendor = row['供应商']
            item = row['物料名称']
            
            if vendor not in self.vendor_stats:
                self.vendor_stats[vendor] = []
            self.vendor_stats[vendor].append(item)
        
        # 显示结果
        self.display_results()
        messagebox.showinfo("完成", "价格分析完成")
    
    def display_results(self):
        """显示分析结果"""
        # 清空现有结果
        for item in self.price_tree.get_children():
            self.price_tree.delete(item)
        
        for item in self.vendor_tree.get_children():
            self.vendor_tree.delete(item)
        
        # 显示最低价分析结果
        if self.analysis_result is not None:
            for _, row in self.analysis_result.iterrows():
                self.price_tree.insert("", tk.END, values=(
                    row['物料名称'],
                    f"{row['最低价']:.2f}",
                    row['供应商']
                ))
        
        # 显示供应商中标统计
        if self.vendor_stats is not None:
            for vendor, items in self.vendor_stats.items():
                self.vendor_tree.insert("", tk.END, values=(
                    vendor,
                    len(items),
                    ", ".join(items)
                ))
    
    def export_report(self):
        """导出分析报告"""
        if self.analysis_result is None:
            messagebox.showwarning("警告", "请先进行价格分析")
            return
        
        # 选择导出路径
        file_path = filedialog.asksaveasfilename(
            title="导出分析报告",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        
        if file_path:
            try:
                # 创建Excel写入器
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    # 写入最低价分析结果
                    self.analysis_result.to_excel(writer, sheet_name='最低价分析', index=False)
                    
                    # 写入供应商中标统计
                    vendor_stats_data = []
                    for vendor, items in self.vendor_stats.items():
                        vendor_stats_data.append({
                            '供应商': vendor,
                            '中标数量': len(items),
                            '中标物料': ", ".join(items)
                        })
                    vendor_stats_df = pd.DataFrame(vendor_stats_data)
                    vendor_stats_df.to_excel(writer, sheet_name='供应商中标统计', index=False)
                
                messagebox.showinfo("成功", f"分析报告已导出到 {file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"导出报告时出错：{str(e)}")
    
    def clear_all(self):
        """清空所有数据"""
        # 清空文件列表
        self.files = []
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        
        # 清空数据
        self.data = {}
        self.analysis_result = None
        self.vendor_stats = None
        
        # 清空结果
        for item in self.price_tree.get_children():
            self.price_tree.delete(item)
        
        for item in self.vendor_tree.get_children():
            self.vendor_tree.delete(item)
        
        messagebox.showinfo("提示", "已清空所有数据")
    
    def run(self):
        """运行应用程序"""
        self.root.mainloop()

if __name__ == "__main__":
    app = VendorPriceComparison()
    app.run()

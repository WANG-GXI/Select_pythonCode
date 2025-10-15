import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
import threading

class JobFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("岗位筛选系统 - 多条件同时匹配")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # 设置中文字体支持
        self.style = ttk.Style()
        self.style.configure("TLabel", font=("SimHei", 10))
        self.style.configure("TButton", font=("SimHei", 10))
        self.style.configure("TEntry", font=("SimHei", 10))
        
        # 文件路径变量
        self.file_path = tk.StringVar()
        
        # 创建界面组件
        self.create_widgets()
        
        # 绑定事件
        self.bind_events()
    
    def create_widgets(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 第一行：文件选择
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(file_frame, text="Excel文件路径:").pack(side=tk.LEFT, padx=(0, 10))
        
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_path, width=50)
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        self.browse_btn = ttk.Button(file_frame, text="浏览...", command=self.browse_file)
        self.browse_btn.pack(side=tk.LEFT)
        
        # 第二行：条件输入
        condition_frame = ttk.Frame(main_frame)
        condition_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(condition_frame, text="筛选条件:").pack(side=tk.TOP, anchor=tk.W, pady=(0, 5))
        
        self.condition_text = scrolledtext.ScrolledText(condition_frame, height=4, wrap=tk.WORD)
        self.condition_text.pack(fill=tk.X, expand=True)
        self.condition_text.insert(tk.END, "1. 电气工程\n2. 江苏省\n3. 工程师")
        
        # 第三行：操作按钮
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.search_btn = ttk.Button(btn_frame, text="开始搜索", command=self.start_search)
        self.search_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.save_btn = ttk.Button(btn_frame, text="保存结果", command=self.save_results, state=tk.DISABLED)
        self.save_btn.pack(side=tk.LEFT)
        
        # 第四行：状态和进度
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(status_frame, textvariable=self.status_var).pack(side=tk.LEFT)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, mode="indeterminate")
        self.progress_bar.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(10, 0))
        
        # 第五行：结果显示
        result_frame = ttk.Frame(main_frame)
        result_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(result_frame, text="搜索结果:").pack(side=tk.TOP, anchor=tk.W, pady=(0, 5))
        
        # 结果树状视图
        columns = ("序号", "工作表", "行号", "匹配详情")
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings")
        
        # 设置列宽和标题
        self.result_tree.column("序号", width=50, anchor=tk.CENTER)
        self.result_tree.column("工作表", width=100, anchor=tk.W)
        self.result_tree.column("行号", width=60, anchor=tk.CENTER)
        self.result_tree.column("匹配详情", width=500, anchor=tk.W)
        
        for col in columns:
            self.result_tree.heading(col, text=col)
        
        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_tree.configure(yscrollcommand=scrollbar.set)
        
        # 绑定双击事件查看详情
        self.result_tree.bind("<Double-1>", self.show_result_details)
        
        # 存储搜索结果
        self.results = []
    
    def bind_events(self):
        # 可以在这里绑定其他事件
        pass
    
    def browse_file(self):
        """浏览并选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if file_path:
            self.file_path.set(file_path)
    
    def parse_user_input(self, user_input):
        """解析用户输入，提取所有条件"""
        # 使用正则表达式匹配所有数字和内容
        pattern = r'\d\.\s*([^\d]+?)(?=\s*\d\.|$)'
        matches = re.findall(pattern, user_input)
        
        # 去除空白并返回所有条件
        conditions = [match.strip() for match in matches if match.strip()]
        return conditions
    
    def start_search(self):
        """开始搜索，在新线程中执行以避免界面冻结"""
        # 清空之前的结果
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        self.results = []
        
        # 获取文件路径和条件
        file_path = self.file_path.get()
        user_input = self.condition_text.get("1.0", tk.END)
        
        # 验证输入
        if not file_path:
            messagebox.showerror("错误", "请选择Excel文件")
            return
        
        if not os.path.exists(file_path):
            messagebox.showerror("错误", "文件不存在，请检查路径是否正确")
            return
        
        conditions = self.parse_user_input(user_input)
        if not conditions:
            messagebox.showerror("错误", "输入格式错误，请按照示例格式输入")
            return
        
        # 禁用按钮，显示进度
        self.search_btn.config(state=tk.DISABLED)
        self.save_btn.config(state=tk.DISABLED)
        self.status_var.set("正在搜索...")
        self.progress_bar.start()
        
        # 在新线程中执行搜索
        threading.Thread(target=self.perform_search, args=(file_path, conditions), daemon=True).start()
    
    def perform_search(self, file_path, conditions):
        """执行搜索操作"""
        try:
            results = self.search_jobs_in_excel(file_path, conditions)
            self.results = results
            
            # 在主线程中更新UI
            self.root.after(0, self.update_search_results, results, conditions)
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"搜索过程中出错: {str(e)}"))
            self.root.after(0, self.reset_ui_state)
    
    def search_jobs_in_excel(self, file_path, conditions):
        """在Excel文件中搜索同时满足所有条件的岗位"""
        try:
            # 读取Excel文件的所有工作表
            excel_file = pd.ExcelFile(file_path)
            all_results = []
            
            # 遍历每个工作表
            for sheet_name in excel_file.sheet_names:
                # 读取工作表数据
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                
                # 确保数据框不为空
                if df.empty:
                    continue
                
                # 遍历每一行
                for index, row in df.iterrows():
                    # 检查该行是否包含所有条件
                    row_contains_all_conditions = True
                    match_details = []
                    
                    # 将整行内容合并为一个字符串用于搜索
                    row_text = ' '.join([str(cell_value) for cell_value in row if pd.notna(cell_value)])
                    
                    # 检查每个条件是否都在该行中
                    for condition in conditions:
                        if condition not in row_text:
                            row_contains_all_conditions = False
                            break
                        else:
                            # 记录匹配的详细信息
                            for col_name, cell_value in row.items():
                                cell_str = str(cell_value) if pd.notna(cell_value) else ""
                                if condition in cell_str:
                                    match_details.append(f"'{condition}' 在列 '{col_name}': {cell_str}")
                                    break  # 每个条件只记录一次匹配
                    
                    # 如果行满足所有条件，记录结果
                    if row_contains_all_conditions:
                        result = {
                            '工作表': sheet_name,
                            '行号': index + 2,  # Excel行号从1开始，加上标题行
                            '匹配详情': match_details,
                            '完整数据': row.to_dict()
                        }
                        all_results.append(result)
            
            return all_results
            
        except Exception as e:
            raise e
    
    def update_search_results(self, results, conditions):
        """更新搜索结果到界面"""
        # 重置UI状态
        self.reset_ui_state()
        
        if not results:
            self.status_var.set(f"未找到同时满足条件的岗位信息")
            messagebox.showinfo("提示", f"未找到同时满足条件 {conditions} 的岗位信息。")
            return
        
        # 显示结果
        for i, result in enumerate(results, 1):
            details = "; ".join(result['匹配详情'])
            self.result_tree.insert("", tk.END, values=(
                i,
                result['工作表'],
                result['行号'],
                details
            ))
        
        self.status_var.set(f"搜索完成！共找到 {len(results)} 个匹配的岗位")
        self.save_btn.config(state=tk.NORMAL)
    
    def reset_ui_state(self):
        """重置UI状态"""
        self.search_btn.config(state=tk.NORMAL)
        self.progress_bar.stop()
        self.status_var.set("就绪")
    
    def show_result_details(self, event):
        """显示结果详情"""
        selected_item = self.result_tree.selection()
        if not selected_item:
            return
        
        # 获取选中项的索引
        index = int(self.result_tree.item(selected_item, "values")[0]) - 1
        if 0 <= index < len(self.results):
            result = self.results[index]
            
            # 创建详情窗口
            detail_window = tk.Toplevel(self.root)
            detail_window.title(f"详情 - {result['工作表']} 第{result['行号']}行")
            detail_window.geometry("600x500")
            detail_window.transient(self.root)
            detail_window.grab_set()
            
            # 创建详情框架
            frame = ttk.Frame(detail_window, padding="10")
            frame.pack(fill=tk.BOTH, expand=True)
            
            # 显示匹配详情
            ttk.Label(frame, text="匹配详情:", font=("SimHei", 12, "bold")).pack(anchor=tk.W, pady=(0, 5))
            for detail in result['匹配详情']:
                ttk.Label(frame, text=f"• {detail}").pack(anchor=tk.W)
            
            ttk.Separator(frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
            
            # 显示完整数据
            ttk.Label(frame, text="完整数据:", font=("SimHei", 12, "bold")).pack(anchor=tk.W, pady=(0, 5))
            text_widget = scrolledtext.ScrolledText(frame, wrap=tk.WORD)
            text_widget.pack(fill=tk.BOTH, expand=True)
            
            for key, value in result['完整数据'].items():
                if pd.notna(value):
                    text_widget.insert(tk.END, f"{key}: {value}\n")
            
            text_widget.config(state=tk.DISABLED)
    
    def save_results(self):
        """保存结果到CSV文件"""
        if not self.results:
            messagebox.showinfo("提示", "没有结果可保存")
            return
        
        file_path = self.file_path.get()
        try:
            # 创建保存数据的列表
            save_data = []
            for result in self.results:
                row_data = result['完整数据'].copy()
                row_data['匹配工作表'] = result['工作表']
                row_data['匹配行号'] = result['行号']
                row_data['匹配详情'] = '; '.join(result['匹配详情'])
                save_data.append(row_data)
            
            # 转换为DataFrame
            df_save = pd.DataFrame(save_data)
            
            # 生成输出文件名
            base_name = os.path.splitext(file_path)[0]
            output_file = f"{base_name}_筛选结果.csv"
            
            # 询问是否覆盖现有文件
            if os.path.exists(output_file):
                if not messagebox.askyesno("确认", f"文件 {output_file} 已存在，是否覆盖？"):
                    # 让用户选择新的文件名
                    output_file = filedialog.asksaveasfilename(
                        defaultextension=".csv",
                        filetypes=[("CSV文件", "*.csv")],
                        initialfile=os.path.basename(output_file)
                    )
                    if not output_file:
                        return
            
            # 保存到CSV
            df_save.to_csv(output_file, index=False, encoding='utf-8-sig')
            messagebox.showinfo("成功", f"结果已保存到:\n{output_file}")
            
        except Exception as e:
            messagebox.showerror("错误", f"保存文件时出错: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = JobFilterApp(root)
    root.mainloop()

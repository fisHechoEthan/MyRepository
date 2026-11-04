#可处理word与excel
#适应行高换行确保比例正确（可选择同名文件覆盖方式）

import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import pythoncom
import win32com.client as win32

class OfficeToPDFWithWPS:
    def __init__(self, root):
        self.root = root
        self.root.title("Office 批量转 PDF (WPS 终极排版版)")
        self.root.geometry("750x680") # 稍微加高窗口以容纳新选项
        self.root.resizable(False, False)

        self.input_dir = tk.StringVar()
        self.keyword = tk.StringVar()
        self.output_mode = tk.StringVar(value="keep")
        self.custom_output_dir = tk.StringVar()

        self.convert_word = tk.BooleanVar(value=True)
        self.convert_excel = tk.BooleanVar(value=True)

        # === 新增：同名文件处理方式 ===
        self.exist_file_action = tk.StringVar(value="suffix") # 默认加后缀
        self.custom_str = tk.StringVar(value="_new") # 默认的字符串

        self.status_text = None
        self.progress_var = tk.DoubleVar()
        self.total_files = 0
        self.processed_files = 0

        self.create_widgets()

    def create_widgets(self):
        # 1. 输入文件夹
        tk.Label(self.root, text="1. 选择输入文件夹（自动搜索所有子目录）").pack(anchor='w', padx=10, pady=3)
        frame_input = tk.Frame(self.root)
        frame_input.pack(fill='x', padx=10, pady=3)
        tk.Entry(frame_input, textvariable=self.input_dir, width=65).pack(side='left', padx=2)
        tk.Button(frame_input, text="浏览...", command=self.select_input_dir).pack(side='left', padx=4)

        # 2. 筛选
        filter_frame = tk.LabelFrame(self.root, text="2. 筛选条件", padx=10, pady=5)
        filter_frame.pack(fill='x', padx=10, pady=5)
        tk.Label(filter_frame, text="关键词:").grid(row=0, column=0, sticky='w', pady=2)
        tk.Entry(filter_frame, textvariable=self.keyword, width=40).grid(row=0, column=1, padx=5, pady=2, sticky='w')
        tk.Label(filter_frame, text="（留空则不过滤）").grid(row=0, column=2, sticky='w', padx=5)

        tk.Label(filter_frame, text="文件类型:").grid(row=1, column=0, sticky='w', pady=2)
        tk.Checkbutton(filter_frame, text="Word (.doc/.docx)", variable=self.convert_word).grid(row=1, column=1, sticky='w', padx=5, pady=2)
        tk.Checkbutton(filter_frame, text="Excel (.xls/.xlsx)", variable=self.convert_excel).grid(row=1, column=2, sticky='w', padx=5, pady=2)

        # 3. 输出模式
        tk.Label(self.root, text="3. 输出模式").pack(anchor='w', padx=10, pady=3)
        mode_frame = tk.Frame(self.root)
        mode_frame.pack(anchor='w', padx=15)
        tk.Radiobutton(mode_frame, text="保留原目录结构", variable=self.output_mode, value="keep").pack(anchor='w')
        tk.Radiobutton(mode_frame, text="统一输出到指定文件夹", variable=self.output_mode, value="flat", command=self.toggle_custom_output).pack(anchor='w')

        self.custom_frame = tk.Frame(self.root)
        tk.Label(self.custom_frame, text="输出文件夹:").pack(side='left', padx=5)
        tk.Entry(self.custom_frame, textvariable=self.custom_output_dir, width=50).pack(side='left', padx=5)
        tk.Button(self.custom_frame, text="浏览...", command=self.select_custom_output_dir).pack(side='left', padx=4)
        self.custom_frame.pack_forget()

        # === 4. 新增：同名PDF处理策略 ===
        duplicate_frame = tk.LabelFrame(self.root, text="4. 同名 PDF 文件处理", padx=10, pady=5)
        duplicate_frame.pack(fill='x', padx=10, pady=5)
        
        # 选项行
        radio_row = tk.Frame(duplicate_frame)
        radio_row.pack(fill='x', pady=2)
        tk.Radiobutton(radio_row, text="添加自定义后缀", variable=self.exist_file_action, value="suffix").pack(side='left', padx=5)
        tk.Radiobutton(radio_row, text="添加自定义前缀", variable=self.exist_file_action, value="prefix").pack(side='left', padx=15)
        tk.Radiobutton(radio_row, text="直接覆盖原PDF", variable=self.exist_file_action, value="overwrite").pack(side='left', padx=5)
        
        # 自定义字符串输入框
        str_row = tk.Frame(duplicate_frame)
        str_row.pack(fill='x', pady=2)
        tk.Label(str_row, text="自定义前缀/后缀内容:").pack(side='left', padx=5)
        tk.Entry(str_row, textvariable=self.custom_str, width=20).pack(side='left', padx=5)
        tk.Label(str_row, text="(如填 _备份，则变为 文件名_备份.pdf)", fg="gray").pack(side='left', padx=5)

        # 5. 状态
        tk.Label(self.root, text="5. 处理状态").pack(anchor='w', padx=10, pady=3)
        self.status_text = tk.Text(self.root, height=10, state='disabled')
        self.status_text.pack(fill='both', expand=True, padx=10, pady=3)

        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill='x', padx=10, pady=5)

        self.start_btn = tk.Button(self.root, text="开始转换", command=self.start_conversion, bg="#4CAF50", fg="white", font=("Arial", 12))
        self.start_btn.pack(pady=10)

    def select_input_dir(self):
        path = filedialog.askdirectory(title="选择输入文件夹")
        if path: self.input_dir.set(path)

    def select_custom_output_dir(self):
        path = filedialog.askdirectory(title="选择PDF输出文件夹")
        if path: self.custom_output_dir.set(path)

    def toggle_custom_output(self):
        if self.output_mode.get() == "flat":
            self.custom_frame.pack(anchor='w', padx=15, pady=5)
        else:
            self.custom_frame.pack_forget()

    def log(self, msg):
        self.status_text.config(state='normal')
        self.status_text.insert('end', msg + "\n")
        self.status_text.see('end')
        self.status_text.config(state='disabled')
        self.root.update()

    def update_progress(self):
        if self.total_files > 0:
            self.progress_var.set((self.processed_files / self.total_files) * 100)
            self.root.update()

    # ================ 核心转换逻辑 ================
    def convert_worker(self, input_dir):
        try:
            pythoncom.CoInitialize()
            
            supported_extensions = []
            if self.convert_word.get(): supported_extensions.extend(['.doc', '.docx'])
            if self.convert_excel.get(): supported_extensions.extend(['.xls', '.xlsx'])

            if not supported_extensions:
                self.log("！未选择任何文件类型。")
                return

            office_files = []
            keyword = self.keyword.get().strip()
            ext_tuple = tuple(supported_extensions)

            for root, _, files in os.walk(input_dir):
                for f in files:
                    if f.lower().endswith(ext_tuple):
                        full_path = os.path.join(root, f)
                        if keyword:
                            if re.search(keyword, f, re.IGNORECASE): office_files.append(full_path)
                        else:
                            office_files.append(full_path)

            self.total_files = len(office_files)
            if self.total_files == 0:
                self.log("！未找到符合条件的文件")
                return

            self.log(f"共找到 {self.total_files} 个文件，开始转换...")
            self.processed_files = 0

            for idx, file_path in enumerate(office_files, 1):
                file_ext = os.path.splitext(file_path)[1].lower()
                base_name = os.path.basename(file_path)
                name_only = os.path.splitext(base_name)[0]
                
                # === 确定 PDF 输出路径（处理同名冲突） ===
                if self.output_mode.get() == "keep":
                    out_dir = os.path.dirname(file_path)
                else:
                    out_dir = self.custom_output_dir.get()

                # 基础文件名
                pdf_name = name_only + ".pdf"
                pdf_path = os.path.join(out_dir, pdf_name)

                # 判断同名文件是否存在，并执行策略
                if os.path.exists(pdf_path):
                    action = self.exist_file_action.get()
                    custom_str = self.custom_str.get()
                    if action == "overwrite":
                        pass # 直接覆盖，路径不变
                    elif action == "suffix":
                        pdf_name = f"{name_only}{custom_str}.pdf"
                        pdf_path = os.path.join(out_dir, pdf_name)
                    elif action == "prefix":
                        pdf_name = f"{custom_str}{name_only}.pdf"
                        pdf_path = os.path.join(out_dir, pdf_name)
                    
                    self.log(f"ℹ 发现同名PDF，采用规则: {action} -> {pdf_name}")

                try:
                    if file_ext in ('.doc', '.docx'):
                        wps_word = win32.Dispatch("Kwps.Application")
                        try: wps_word.Visible = False
                        except: pass
                        wps_word.DisplayAlerts = False
                        
                        doc = wps_word.Documents.Open(file_path)
                        doc.ExportAsFixedFormat(pdf_path, 17)
                        doc.Close(SaveChanges=False)
                        wps_word.Quit()
                        
                    elif file_ext in ('.xls', '.xlsx'):
                        wps_excel = win32.Dispatch("Ket.Application")
                        try: wps_excel.Visible = False
                        except: pass
                        wps_excel.DisplayAlerts = False 
                        
                        wb = wps_excel.Workbooks.Open(file_path)
                        
                        for sheet in wb.Sheets:
                            # 1. 开启自动换行
                            sheet.Cells.WrapText = True
                            
                            # 2. 自动调整列宽适应内容
                            sheet.Columns.AutoFit()
                            
                            # 3. 缩放到一页宽
                            sheet.PageSetup.Zoom = False
                            sheet.PageSetup.FitToPagesWide = 1
                            sheet.PageSetup.FitToPagesTall = False 
                            
                            # 4. 触发 WPS 底层重新计算真实需要的行高
                            sheet.Rows.AutoFit()
                            
                            # 5. 【核心修复】加大行高，保证文字不遮挡
                            try:
                                used_range = sheet.UsedRange
                                start_row = used_range.Row
                                row_count = used_range.Rows.Count
                                
                                for r in range(start_row, start_row + row_count):
                                    current_height = sheet.Rows(r).RowHeight
                                    if current_height > 0:
                                        # 如果行高已经大于40（说明包含大量换行文本），放大 40%
                                        if current_height > 40:
                                            sheet.Rows(r).RowHeight = current_height * 1.40
                                        # 如果行高在15-40之间（普通几行文字），放大 25%
                                        elif current_height > 15:
                                            sheet.Rows(r).RowHeight = current_height * 1.25
                                        # 单行文字，放大 10%
                                        else:
                                            sheet.Rows(r).RowHeight = current_height * 1.10
                            except Exception:
                                pass 
                            
                        # 导出
                        wb.ExportAsFixedFormat(0, pdf_path)
                        
                        # 绝对不保存修改
                        wb.Close(SaveChanges=False)
                        wps_excel.Quit()

                    self.processed_files += 1
                    self.log(f"[{idx}/{self.total_files}] 转换成功: {base_name} -> {pdf_name}")
                    self.update_progress()

                except Exception as e:
                    self.log(f"！[{idx}/{self.total_files}] 转换失败: {base_name} → {str(e)}")
                    try: wps_word.Quit()
                    except: pass
                    try: wps_excel.Quit()
                    except: pass

            self.log("\n所有文件转换完成！")
            messagebox.showinfo("完成", f"转换完成！共处理 {self.processed_files} 个文件。")

        except Exception as e:
            self.log(f"❌ 转换过程中发生异常: {str(e)}")
            messagebox.showerror("错误", f"转换失败: {str(e)}")
        finally:
            try: pythoncom.CoUninitialize()
            except: pass
            self.start_btn.config(state='normal')

    def start_conversion(self):
        input_dir = self.input_dir.get().strip()
        if not input_dir or not os.path.exists(input_dir):
            messagebox.showerror("错误", "请选择有效的输入文件夹")
            return

        if not self.convert_word.get() and not self.convert_excel.get():
            messagebox.showerror("错误", "请至少勾选一种文件类型")
            return

        if self.output_mode.get() == "flat":
            output_dir = self.custom_output_dir.get().strip()
            if not output_dir:
                messagebox.showerror("错误", "平铺模式下请选择输出文件夹")
                return
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

        self.start_btn.config(state='disabled')
        self.status_text.config(state='normal')
        self.status_text.delete('1.0', 'end')
        self.status_text.config(state='disabled')
        self.progress_var.set(0)

        thread = threading.Thread(target=self.convert_worker, args=(input_dir,))
        thread.daemon = True
        thread.start()

if __name__ == "__main__":
    root = tk.Tk()
    app = OfficeToPDFWithWPS(root)
    root.mainloop()
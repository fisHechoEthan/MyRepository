import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import os
import threading
from stamper import PDFStamper

class ElectronicSealApp:
    def __init__(self, root):
        self.root = root
        self.root.title("电子印章工具 - 关键词定位")
        self.root.geometry("900x850")
        self.style = ttk.Style("cosmo")

        # 原有变量
        self.input_directory = tk.StringVar()
        self.stamp_file_path = tk.StringVar()
        
        # 模式
        self.mode_var = tk.StringVar(value="first")
        self.cross_page_var = tk.BooleanVar(value=False)

        # 多页设置
        self.page_limit = tk.StringVar(value="")
        self.random_jitter = tk.BooleanVar(value=True)
        self.jitter_amplitude = tk.IntVar(value=20)

        # 骑缝章设置
        self.cross_page_extra_ratio = tk.DoubleVar(value=0.15)

        # ===== 新增：关键词定位参数 =====
        self.use_keyword_position = tk.BooleanVar(value=False)
        self.keyword_text = tk.StringVar(value="报告编号")  # 默认关键词
        self.offset_up_cm = tk.DoubleVar(value=0.1)  # 向上偏移（厘米）

        # 原有参数（作为备选）
        self.pos_x = tk.DoubleVar(value=400.0)
        self.pos_y = tk.DoubleVar(value=500.0)
        self.scale = tk.DoubleVar(value=1.0)
        
        self.is_processing = False
        self.processed_count = 0
        self.total_count = 0

        self._create_widgets()

    def _create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=BOTH, expand=YES)

        ttk.Label(main_frame, text="电子印章工具 - 关键词定位版", font=("Helvetica", 24, "bold"), bootstyle="primary").pack(pady=(0, 20))

        # 1. 文件选择
        file_frame = ttk.Labelframe(main_frame, text="文件选择", padding="15")
        file_frame.pack(fill=X, pady=10)
        self._create_dir_input(file_frame, "输入目录:", self.input_directory, self._select_input_directory)
        self._create_file_input(file_frame, "电子印章:", self.stamp_file_path, self._select_stamp_file)

        # 2. 盖章选项
        options_frame = ttk.Labelframe(main_frame, text="盖章选项", padding="15")
        options_frame.pack(fill=X, pady=10)

        # --- 模式选择 ---
        mode_frame = ttk.Frame(options_frame)
        mode_frame.pack(fill=X, pady=5)
        ttk.Label(mode_frame, text="盖章模式:", width=15, font=("Helvetica", 10, "bold")).pack(side=LEFT)
        ttk.Radiobutton(mode_frame, text="首页盖章", variable=self.mode_var, value="first").pack(side=LEFT, padx=10)
        ttk.Radiobutton(mode_frame, text="多页盖章", variable=self.mode_var, value="multi").pack(side=LEFT, padx=10)
        ttk.Separator(mode_frame, orient=VERTICAL).pack(side=LEFT, fill=Y, padx=20)
        ttk.Checkbutton(mode_frame, text="叠加骑缝章", variable=self.cross_page_var, bootstyle="round-toggle").pack(side=LEFT, padx=10)

        # --- 多页设置 ---
        multi_frame = ttk.Frame(options_frame)
        multi_frame.pack(fill=X, pady=3, padx=(20, 0))
        ttk.Label(multi_frame, text="[多页] 页数 (空=全部):", bootstyle="secondary").pack(side=LEFT)
        ttk.Entry(multi_frame, textvariable=self.page_limit, width=5).pack(side=LEFT, padx=5)
        ttk.Checkbutton(multi_frame, text="随机微调", variable=self.random_jitter, bootstyle="square-toggle").pack(side=LEFT, padx=10)
        ttk.Label(multi_frame, text="幅度:", bootstyle="secondary").pack(side=LEFT)
        ttk.Entry(multi_frame, textvariable=self.jitter_amplitude, width=4).pack(side=LEFT, padx=2)

        # --- 关键词定位设置（新增） ---
        keyword_frame = ttk.Labelframe(options_frame, text="首页关键词定位（仅首页盖章）", padding="10")
        keyword_frame.pack(fill=X, pady=8)

        # 启用开关
        enable_frame = ttk.Frame(keyword_frame)
        enable_frame.pack(fill=X, pady=3)
        ttk.Checkbutton(enable_frame, text="启用关键词定位", variable=self.use_keyword_position, 
                       bootstyle="round-toggle").pack(side=LEFT, padx=5)
        ttk.Label(enable_frame, text="（勾选后，X/Y坐标将被忽略，按关键词位置自动定位）", 
                 bootstyle="secondary", font=("Arial", 9)).pack(side=LEFT, padx=10)

        # 关键词输入
        kw_frame = ttk.Frame(keyword_frame)
        kw_frame.pack(fill=X, pady=3, padx=(10, 0))
        ttk.Label(kw_frame, text="关键词:", bootstyle="secondary").pack(side=LEFT)
        ttk.Entry(kw_frame, textvariable=self.keyword_text, width=30).pack(side=LEFT, padx=5)
        ttk.Label(kw_frame, text="（支持正则表达式）", bootstyle="secondary", font=("Arial", 9)).pack(side=LEFT, padx=5)

        # 偏移量设置
        offset_frame = ttk.Frame(keyword_frame)
        offset_frame.pack(fill=X, pady=3, padx=(10, 0))
        ttk.Label(offset_frame, text="向上偏移:", bootstyle="secondary").pack(side=LEFT)
        ttk.Spinbox(offset_frame, from_=0.0, to=2.0, increment=0.05, 
                   textvariable=self.offset_up_cm, width=6).pack(side=LEFT, padx=5)
        ttk.Label(offset_frame, text="厘米", bootstyle="secondary").pack(side=LEFT)
        ttk.Label(offset_frame, text="（从关键词中心向上偏移）", bootstyle="secondary", font=("Arial", 9)).pack(side=LEFT, padx=10)

        # --- 原有位置参数（作为备选） ---
        pos_label = ttk.Label(options_frame, text="手动定位（关键词定位未启用时使用）", bootstyle="secondary", font=("Arial", 9))
        pos_label.pack(anchor='w', pady=(10, 3))
        
        self._create_slider_input(options_frame, "X 坐标 (点):", self.pos_x, 0, 800)
        self._create_slider_input(options_frame, "Y 坐标 (点):", self.pos_y, 0, 1100)
        self._create_slider_input(options_frame, "缩放比例:", self.scale, 0.1, 5.0)

        # 3. 操作按钮
        action_frame = ttk.Frame(main_frame, padding="0 20")
        action_frame.pack(fill=X)
        self.stamp_button = ttk.Button(action_frame, text="开始批量盖章", command=self._on_stamp_click, 
                                      bootstyle="success-lg", width=20)
        self.stamp_button.pack()
        
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate', bootstyle="success-striped")
        self.status_label = ttk.Label(main_frame, text="准备就绪", bootstyle="secondary")
        self.status_label.pack(pady=10)

    def _create_dir_input(self, parent, label_text, variable, command):
        frame = ttk.Frame(parent)
        frame.pack(fill=X, pady=5)
        ttk.Label(frame, text=label_text, width=15).pack(side=LEFT)
        ttk.Entry(frame, textvariable=variable, state="readonly").pack(side=LEFT, fill=X, expand=YES, padx=5)
        ttk.Button(frame, text="浏览...", command=command, bootstyle="secondary-outline").pack(side=LEFT)

    def _create_file_input(self, parent, label_text, variable, command):
        frame = ttk.Frame(parent)
        frame.pack(fill=X, pady=5)
        ttk.Label(frame, text=label_text, width=15).pack(side=LEFT)
        ttk.Entry(frame, textvariable=variable, state="readonly").pack(side=LEFT, fill=X, expand=YES, padx=5)
        ttk.Button(frame, text="浏览...", command=command, bootstyle="secondary-outline").pack(side=LEFT)

    def _create_slider_input(self, parent, label_text, variable, from_, to_):
        frame = ttk.Frame(parent)
        frame.pack(fill=X, pady=5)
        ttk.Label(frame, text=label_text, width=15).pack(side=LEFT)
        ttk.Scale(frame, from_=from_, to=to_, variable=variable, orient=HORIZONTAL).pack(side=LEFT, fill=X, expand=YES, padx=5)
        ttk.Entry(frame, textvariable=variable, width=8).pack(side=LEFT)

    def _select_input_directory(self):
        directory = filedialog.askdirectory(title="选择包含PDF文件的目录")
        if directory:
            self.input_directory.set(directory)

    def _select_stamp_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Image files", "*.png *.jpg *.jpeg")])
        if filepath:
            self.stamp_file_path.set(filepath)

    def _on_stamp_click(self):
        if self.is_processing:
            return
        
        input_dir = self.input_directory.get()
        stamp_path = self.stamp_file_path.get()
        
        if not input_dir or not os.path.exists(input_dir):
            messagebox.showwarning("输入错误", "请选择有效的输入目录")
            return
        
        if not stamp_path or not os.path.exists(stamp_path):
            messagebox.showwarning("输入错误", "请选择有效的印章图片")
            return

        # 如果启用了关键词定位，检查关键词是否为空
        if self.use_keyword_position.get():
            keyword = self.keyword_text.get().strip()
            if not keyword:
                messagebox.showwarning("输入错误", "请填写关键词")
                return

        self.is_processing = True
        self.stamp_button.configure(state="disabled", text="处理中...")
        self.progress.pack(fill=X, pady=5)
        self.progress.start(10)
        self.status_label.config(text="正在处理...", bootstyle="info")

        thread = threading.Thread(target=self._run_stamping_process, args=(input_dir, stamp_path))
        thread.daemon = True
        thread.start()

    def _run_stamping_process(self, input_dir, stamp_path):
        try:
            pdf_files = self._find_all_pdfs(input_dir)
            total = len(pdf_files)
            
            if total == 0:
                self.root.after(0, lambda: messagebox.showwarning("提示", "未找到PDF文件"))
                self._reset_ui()
                return

            self.root.after(0, lambda: self.status_label.config(text=f"共找到 {total} 个PDF文件，开始处理...", bootstyle="info"))

            success_count = 0
            fail_count = 0

            for idx, pdf_path in enumerate(pdf_files, 1):
                try:
                    self.root.after(0, lambda i=idx, t=total: 
                                   self.status_label.config(text=f"正在处理 [{i}/{t}]: {os.path.basename(pdf_path)}", bootstyle="info"))
                    
                    self._stamp_single_pdf(pdf_path, stamp_path)
                    success_count += 1
                    
                except Exception as e:
                    fail_count += 1
                    print(f"处理失败: {pdf_path} - {str(e)}")

            self.root.after(0, lambda: self._on_finish(success_count, fail_count))

        except Exception as e:
            self.root.after(0, lambda: self._on_error(str(e)))

    def _stamp_single_pdf(self, pdf_path, stamp_path):
        """处理单个PDF文件"""
        stamper = PDFStamper(pdf_path, stamp_path)
        
        scale = self.scale.get()
        mode = self.mode_var.get()
        
        # ===== 首页盖章 =====
        if mode == "first":
            # 判断是否启用关键词定位
            if self.use_keyword_position.get():
                keyword = self.keyword_text.get().strip()
                offset = self.offset_up_cm.get()
                # 使用关键词定位盖章
                stamper.stamp_first_page_by_keyword(keyword, scale, offset_up_cm=offset)
            else:
                # 使用传统坐标定位
                x = self.pos_x.get()
                y = self.pos_y.get()
                stamper.stamp_first_page(x, y, scale)
        
        # ===== 多页盖章 =====
        elif mode == "multi":
            x = self.pos_x.get()
            y = self.pos_y.get()
            limit_str = self.page_limit.get()
            limit = int(limit_str) if limit_str.isdigit() else None
            stamper.stamp_multi_pages(
                x, y, scale,
                page_limit=limit,
                random_offset=self.random_jitter.get(),
                jitter_amplitude=self.jitter_amplitude.get()
            )

        # ===== 骑缝章（叠加） =====
        if self.cross_page_var.get():
            stamper.stamp_cross_page(
                scale,
                first_page_extra_ratio=self.cross_page_extra_ratio.get()
            )

        # 保存
        dir_name = os.path.dirname(pdf_path)
        base_name = os.path.basename(pdf_path)
        name, ext = os.path.splitext(base_name)
        
        suffix_parts = [mode]
        if self.cross_page_var.get():
            suffix_parts.append("cross")
        if self.use_keyword_position.get() and mode == "first":
            suffix_parts.append("kw")
        output_filename = f"{name}_stamped_{'_'.join(suffix_parts)}.pdf"
        output_path = os.path.join(dir_name, output_filename)
        
        stamper.save_pdf(output_path)

    def _find_all_pdfs(self, directory):
        pdf_files = []
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.lower().endswith('.pdf'):
                    pdf_files.append(os.path.join(root, file))
        return pdf_files

    def _on_finish(self, success, fail):
        self._reset_ui()
        self.status_label.config(text=f"处理完成！成功: {success}，失败: {fail}", 
                                bootstyle="success" if fail == 0 else "warning")
        if fail > 0:
            messagebox.showwarning("部分失败", f"处理完成！\n成功: {success}\n失败: {fail}")
        else:
            messagebox.showinfo("完成", f"所有文件处理完成！\n共处理 {success} 个文件。")

    def _on_error(self, error_msg):
        self._reset_ui()
        self.status_label.config(text="处理失败", bootstyle="danger")
        messagebox.showerror("错误", f"处理失败：\n{error_msg}")

    def _reset_ui(self):
        self.is_processing = False
        self.progress.stop()
        self.progress.pack_forget()
        self.stamp_button.configure(state="normal", text="开始批量盖章")

def main():
    root = ttk.Window(themename="cosmo")
    app = ElectronicSealApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
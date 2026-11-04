#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
多功能文件处理工具 (GUI版)
功能1：自定义解压引擎，递归解压（排除Office文件）
功能2：递归统计文件，支持模糊匹配
"""

import os
import sys
import subprocess
import logging
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from collections import Counter

# ================= 基础配置 =================
PSEUDO_ARCHIVE_EXTS = {
    ".docx", ".docm", ".dotx", ".dotm",
    ".xlsx", ".xlsm", ".xltx", ".xltm",
    ".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm",
    ".odt", ".ods", ".odp", ".odg",
    ".jar", ".apk", ".epub", ".whl", ".nupkg", ".vsix", ".ipa",
    ".xpi", ".crx",
}
ARCHIVE_EXTS = [".rar", ".zip", ".7z", ".tar", ".gz", ".bz2"]
MAX_DEPTH = 30

# 自动探测常见解压软件路径
COMMON_UNZIP_PATHS = [
    r"C:\Program Files\WinRAR\WinRAR.exe",
    r"C:\Program Files (x86)\WinRAR\WinRAR.exe",
    r"D:\Program Files\WinRAR\WinRAR.exe",
    r"C:\Program Files\7-Zip\7z.exe",
    r"C:\Program Files (x86)\7-Zip\7z.exe",
    r"C:\Program Files\Bandizip\Bandizip.exe",
    r"C:\Program Files\Bandizip\bz.exe",
]

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("文件批处理工具箱")
        self.root.geometry("750x650")
        self.root.resizable(True, True)
        
        # 创建选项卡
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Tab 1: 递归解压
        self.tab_unzip = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_unzip, text="📦 递归解压")
        self.create_unzip_ui()
        
        # Tab 2: 文件统计
        self.tab_stat = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_stat, text="📊 文件统计")
        self.create_stat_ui()

    # ================= 功能一：递归解压 =================
    def create_unzip_ui(self):
        frame = tk.Frame(self.tab_unzip, padx=10, pady=10)
        frame.pack(fill='both', expand=True)
        
        # 1. 选择目录
        tk.Label(frame, text="1. 选择待解压文件夹 (或单个压缩包):").pack(anchor='w')
        row1 = tk.Frame(frame)
        row1.pack(fill='x', pady=5)
        self.unzip_dir = tk.StringVar()
        tk.Entry(row1, textvariable=self.unzip_dir, width=60).pack(side='left', padx=5)
        tk.Button(row1, text="浏览...", command=self.select_unzip_dir).pack(side='left')
        
        # 2. 解压引擎选择
        tk.Label(frame, text="2. 解压引擎 (默认自动探测):").pack(anchor='w', pady=(10,0))
        row2 = tk.Frame(frame)
        row2.pack(fill='x', pady=5)
        self.unzip_tool = tk.StringVar()
        # 自动探测默认路径
        default_tool = ""
        for p in COMMON_UNZIP_PATHS:
            if os.path.exists(p):
                default_tool = p
                break
        self.unzip_tool.set(default_tool)
        tk.Entry(row2, textvariable=self.unzip_tool, width=60).pack(side='left', padx=5)
        tk.Button(row2, text="浏览...", command=self.select_unzip_tool).pack(side='left')
        
        # 3. 操作按钮与日志
        tk.Button(frame, text="开始递归解压", bg="#4CAF50", fg="white", 
                  font=("Arial", 11), command=self.start_unzip).pack(pady=15)
        
        tk.Label(frame, text="处理日志:").pack(anchor='w')
        self.unzip_log = tk.Text(frame, height=15, state='disabled', bg="#f0f0f0")
        self.unzip_log.pack(fill='both', expand=True, pady=5)

    def select_unzip_dir(self):
        path = filedialog.askdirectory(title="选择包含压缩包的文件夹")
        if path: self.unzip_dir.set(path)
        
    def select_unzip_tool(self):
        path = filedialog.askopenfilename(title="选择解压软件", filetypes=[("Executable", "*.exe")])
        if path: self.unzip_tool.set(path)

    def log_unzip(self, msg):
        self.unzip_log.config(state='normal')
        self.unzip_log.insert('end', msg + "\n")
        self.unzip_log.see('end')
        self.unzip_log.config(state='disabled')
        self.root.update()

    def is_archive(self, path: Path) -> bool:
        suffix = path.suffix.lower()
        if suffix in PSEUDO_ARCHIVE_EXTS: return False
        if suffix in ARCHIVE_EXTS: return True
        # 魔数判断
        try:
            with open(path, "rb") as f: head = f.read(8)
            if head[:4] in (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08"): return True
            if head[:7] == b"Rar!\x1a\x07\x00": return True
            if head[:8] == b"Rar!\x1a\x07\x01\x00": return True
            if head[:6] == b"7z\xbc\xaf\x27\x1c": return True
        except Exception:
            return False
        return False

    def find_archives(self, root: Path):
        result = []
        for dirpath, _, filenames in os.walk(root):
            for fn in filenames:
                p = Path(dirpath) / fn
                if self.is_archive(p): result.append(p)
        return result

    def extract_archive(self, archive: Path, tool_path: str) -> bool:
        target = archive.parent / (archive.stem if archive.suffix else archive.name + "_out")
        target.mkdir(parents=True, exist_ok=True)
        
        # 自动适配不同解压软件的命令行参数
        tool_name = os.path.basename(tool_path).lower()
        if "7z" in tool_name:
            cmd = [tool_path, "x", str(archive), f"-o{str(target)}", "-y"]
        elif "bz" in tool_name or "bandizip" in tool_name:
            cmd = [tool_path, "x", "-o:" + str(target), "-y", "-target:auto", str(archive)]
        else:
            # 默认按 WinRAR 处理
            cmd = [tool_path, "x", "-y", "-ibck", "-o+", "-p-", str(archive), str(target) + os.sep]

        self.log_unzip(f"  解压：{archive.name} -> {target.name}/")
        try:
            r = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=7200)
            # 不同的解压软件返回码不同，0 通常是成功
            if r.returncode == 0:
                return True
            self.log_unzip(f"  ❌ 解压失败 code={r.returncode}")
            return False
        except Exception as e:
            self.log_unzip(f"  ❌ 异常: {e}")
            return False

    def unzip_worker(self, target_path: Path, tool_path: str):
        processed = set()
        round_no = 0
        
        while round_no < MAX_DEPTH:
            round_no += 1
            self.log_unzip(f"========== 第 {round_no} 轮扫描 ==========")
            
            if target_path.is_file():
                all_archives = [target_path] if self.is_archive(target_path) else []
            else:
                all_archives = self.find_archives(target_path)
                
            pending = [a for a in all_archives if str(a.resolve()) not in processed]
            if not pending:
                self.log_unzip("✅ 所有压缩包已处理完毕。")
                break
                
            self.log_unzip(f"发现 {len(pending)} 个待处理压缩包")
            for arc in pending:
                self.extract_archive(arc, tool_path)
                processed.add(str(arc.resolve()))
        else:
            self.log_unzip(f"⚠️ 达到最大递归深度 {MAX_DEPTH}")

    def start_unzip(self):
        path = self.unzip_dir.get().strip()
        tool = self.unzip_tool.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("错误", "请选择有效的文件夹或文件")
            return
        if not tool or not os.path.exists(tool):
            messagebox.showerror("错误", "请选择有效的解压软件 exe")
            return
        
        self.unzip_log.config(state='normal')
        self.unzip_log.delete('1.0', 'end')
        self.unzip_log.config(state='disabled')
        
        threading.Thread(target=self.unzip_worker, args=(Path(path), tool), daemon=True).start()

    # ================= 功能二：文件统计 =================
    def create_stat_ui(self):
        frame = tk.Frame(self.tab_stat, padx=10, pady=10)
        frame.pack(fill='both', expand=True)
        
        # 1. 选择目录
        tk.Label(frame, text="1. 选择要统计的文件夹:").pack(anchor='w')
        row1 = tk.Frame(frame)
        row1.pack(fill='x', pady=5)
        self.stat_dir = tk.StringVar()
        tk.Entry(row1, textvariable=self.stat_dir, width=60).pack(side='left', padx=5)
        tk.Button(row1, text="浏览...", command=self.select_stat_dir).pack(side='left')
        
        # 2. 过滤条件
        filter_frame = tk.LabelFrame(frame, text="2. 过滤条件 (留空则统计全部)", padx=5, pady=5)
        filter_frame.pack(fill='x', pady=10)
        
        # 类型过滤
        row_type = tk.Frame(filter_frame)
        row_type.pack(fill='x', pady=2)
        tk.Label(row_type, text="文件后缀类型:").pack(side='left')
        self.stat_types = tk.StringVar()
        tk.Entry(row_type, textvariable=self.stat_types, width=30).pack(side='left', padx=5)
        tk.Label(row_type, text="(如: .doc,.docx,.pdf 以逗号分隔)", fg="gray").pack(side='left')
        
        # 模糊字符过滤
        row_kw = tk.Frame(filter_frame)
        row_kw.pack(fill='x', pady=2)
        tk.Label(row_kw, text="文件名包含字符:").pack(side='left')
        self.stat_keyword = tk.StringVar()
        tk.Entry(row_kw, textvariable=self.stat_keyword, width=30).pack(side='left', padx=5)
        tk.Label(row_kw, text="(支持模糊匹配)", fg="gray").pack(side='left')

        # 3. 按钮与结果
        tk.Button(frame, text="开始统计", bg="#2196F3", fg="white", 
                  font=("Arial", 11), command=self.start_stat).pack(pady=10)
        
        # 结果表格
        res_frame = tk.Frame(frame)
        res_frame.pack(fill='both', expand=True, pady=5)
        
        self.tree = ttk.Treeview(res_frame, columns=("type", "count"), show="headings")
        self.tree.heading("type", text="文件类型 / 分类")
        self.tree.heading("count", text="数量")
        self.tree.column("type", width=300)
        self.tree.column("count", width=100, anchor='center')
        
        scrollbar = ttk.Scrollbar(res_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

    def select_stat_dir(self):
        path = filedialog.askdirectory(title="选择要统计的文件夹")
        if path: self.stat_dir.set(path)

    def stat_worker(self, root_dir: Path, type_str: str, keyword: str):
        # 解析允许的类型
        allowed_types = []
        if type_str:
            allowed_types = [t.strip().lower() for t in type_str.split(',') if t.strip()]
            allowed_types = [t if t.startswith('.') else '.' + t for t in allowed_types]
            
        type_counter = Counter()
        total_count = 0
        
        for dirpath, _, filenames in os.walk(root_dir):
            for fn in filenames:
                f_lower = fn.lower()
                suffix = Path(fn).suffix.lower()
                
                # 1. 类型过滤
                if allowed_types and suffix not in allowed_types:
                    continue
                # 2. 关键字过滤 (忽略大小写)
                if keyword and keyword.lower() not in f_lower:
                    continue
                
                # 统计
                type_key = suffix if suffix else "无扩展名"
                type_counter[type_key] += 1
                total_count += 1
                
        # 更新UI
        self.root.after(0, lambda: self.show_stat_results(type_counter, total_count, allowed_types, keyword))

    def show_stat_results(self, counter, total, allowed_types, keyword):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        if total == 0:
            self.tree.insert("", "end", values=("未找到匹配的文件", 0))
            return
            
        # 插入分类统计
        for t, c in counter.most_common():
            self.tree.insert("", "end", values=(t, c))
            
        # 插入总计
        self.tree.insert("", "end", values=("【总计】", total))
        
        # 如果是特定搜索，提供弹窗提示
        msg = f"统计完成！\n共找到 {total} 个文件。"
        if allowed_types: msg += f"\n限定类型: {', '.join(allowed_types)}"
        if keyword: msg += f"\n包含关键字: '{keyword}'"
        messagebox.showinfo("统计结果", msg)

    def start_stat(self):
        path = self.stat_dir.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("错误", "请选择有效的文件夹")
            return
            
        type_str = self.stat_types.get().strip()
        keyword = self.stat_keyword.get().strip()
        
        self.tree.delete(*self.tree.get_children())
        self.tree.insert("", "end", values=("正在扫描，请稍候...", "..."))
        
        threading.Thread(target=self.stat_worker, args=(Path(path), type_str, keyword), daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
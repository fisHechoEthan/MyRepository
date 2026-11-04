import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from collections import defaultdict

try:
    import docx
    import openpyxl
except ImportError:
    messagebox.showerror("缺少依赖", "请先在命令行运行：\npip install python-docx openpyxl")
    exit()

SUPPORTED_EXTS = {
    "纯文本文件 (txt/md等)": ['.txt', '.md', '.csv'],
    "Word 文档 (.docx)": ['.docx'],
    "Excel 表格 (.xlsx)": ['.xlsx']
}

class AdvancedBatchProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("高级批量处理工具 (详尽日志版)")
        self.root.geometry("950x750")
        
        self.folder_path = tk.StringVar()
        self.file_type = tk.StringVar(value="Word 文档 (.docx)")
        self.is_running = False
        self.rules = [] # 存放所有的规则

        self.create_widgets()

    def create_widgets(self):
        # --- 1. 目录与文件类型 ---
        frame_top = ttk.Frame(self.root, padding=10)
        frame_top.pack(fill="x")
        
        ttk.Label(frame_top, text="处理目录:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_top, textvariable=self.folder_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(frame_top, text="浏览...", command=self.browse_folder).grid(row=0, column=2)
        
        ttk.Label(frame_top, text="文件类型:").grid(row=1, column=0, sticky="w", pady=5)
        combo = ttk.Combobox(frame_top, textvariable=self.file_type, values=list(SUPPORTED_EXTS.keys()), state="readonly", width=25)
        combo.grid(row=1, column=1, sticky="w", padx=5)

        # --- 2. 规则配置区 ---
        frame_rule = ttk.LabelFrame(self.root, text="配置处理规则", padding=10)
        frame_rule.pack(fill="x", padx=10, pady=5)

        # 操作栏
        frame_input = ttk.Frame(frame_rule)
        frame_input.pack(fill="x", pady=5)
        
        ttk.Label(frame_input, text="操作:").pack(side="left")
        self.op_combo = ttk.Combobox(frame_input, values=["替换", "前方插入", "后方插入", "查找"], state="readonly", width=10)
        self.op_combo.set("查找")
        self.op_combo.pack(side="left", padx=5)

        ttk.Label(frame_input, text="目标文字:").pack(side="left")
        self.entry_target = ttk.Entry(frame_input, width=25)
        self.entry_target.pack(side="left", padx=5)

        ttk.Label(frame_input, text="替换/插入文字:").pack(side="left")
        self.entry_replace = ttk.Entry(frame_input, width=30)
        self.entry_replace.pack(side="left", padx=5)

        ttk.Button(frame_input, text="➕ 添加规则", command=self.add_rule).pack(side="left", padx=10)
        ttk.Button(frame_input, text="❌ 删除选中", command=self.del_rule).pack(side="left")

        # 规则列表
        self.tree = ttk.Treeview(frame_rule, columns=("op", "target", "replace"), show="headings", height=5)
        self.tree.heading("op", text="操作类型")
        self.tree.heading("target", text="目标文字")
        self.tree.heading("replace", text="替换/插入文字")
        self.tree.column("op", width=100)
        self.tree.column("target", width=350)
        self.tree.column("replace", width=350)
        self.tree.pack(fill="x", pady=5)

        # --- 3. 执行控制与日志 ---
        frame_ctrl = ttk.Frame(self.root)
        frame_ctrl.pack(fill="x", padx=10, pady=10)
        self.btn_run = ttk.Button(frame_ctrl, text="▶ 开始执行", command=self.start_processing)
        self.btn_run.pack(side="left", padx=10)
        ttk.Button(frame_ctrl, text="清空日志", command=self.clear_log).pack(side="left")

        frame_log = ttk.LabelFrame(self.root, text="详细执行日志", padding=10)
        frame_log.pack(fill="both", expand=True, padx=10, pady=5)
        self.log_area = scrolledtext.ScrolledText(frame_log, wrap="word", state="disabled", font=("Consolas", 10))
        self.log_area.pack(fill="both", expand=True)

    def browse_folder(self):
        path = filedialog.askdirectory()
        if path: self.folder_path.set(path)

    def add_rule(self):
        op = self.op_combo.get()
        target = self.entry_target.get()
        replace = self.entry_replace.get()
        
        if not target:
            messagebox.showwarning("提示", "目标文字不能为空！")
            return
        if op != "查找" and not replace:
            messagebox.showwarning("提示", "请输入替换/插入的文字！")
            return

        self.rules.append({"op": op, "target": target, "replace": replace})
        self.tree.insert("", "end", values=(op, target, replace))
        
        self.entry_target.delete(0, "end")
        self.entry_replace.delete(0, "end")

    def del_rule(self):
        selected = self.tree.selection()
        if not selected: return
        index = self.tree.index(selected[0])
        self.rules.pop(index)
        self.tree.delete(selected[0])

    def log(self, msg, tag="info"):
        self.log_area.config(state="normal")
        self.log_area.insert("end", msg + "\n", tag)
        self.log_area.tag_config("red", foreground="red")
        self.log_area.tag_config("green", foreground="#27AE60")
        self.log_area.tag_config("blue", foreground="#2980B9")
        self.log_area.tag_config("orange", foreground="#E67E22")
        self.log_area.see("end")
        self.log_area.config(state="disabled")

    def clear_log(self):
        self.log_area.config(state="normal")
        self.log_area.delete(1.0, "end")
        self.log_area.config(state="disabled")

    def start_processing(self):
        if self.is_running or not self.rules: 
            messagebox.showinfo("提示", "请先添加至少一条规则！")
            return
        folder = self.folder_path.get()
        if not folder or not os.path.exists(folder):
            messagebox.showerror("错误", "请选择有效的文件夹！")
            return

        self.is_running = True
        self.btn_run.config(text="执行中...", state="disabled")
        self.clear_log()
        self.log(f"========== 开始批量处理 ==========", "blue")
        self.log(f"共有 {len(self.rules)} 条规则将依次生效...\n")

        threading.Thread(target=self.process_files, args=(folder,), daemon=True).start()

    # ================= 核心文本处理逻辑 =================
    
    def process_string(self, text, rules):
        """
        处理单段文字，返回 (新文本, 是否修改, 修改明细字典)
        """
        if not text: return text, False, {}
        
        modified = False
        changes = defaultdict(int) # 记录每条规定执行了几次
        sorted_rules = sorted(rules, key=lambda x: len(x["target"]), reverse=True)
        placeholders = {}
        current_text = text
        
        for i, rule in enumerate(sorted_rules):
            target = rule["target"]
            replace = rule["replace"]
            op = rule["op"]
            rule_key = f"[{op}] {target}"
            
            if op == "查找":
                count = current_text.count(target)
                if count > 0:
                    changes[rule_key] += count
                    modified = True
                continue
                
            if target not in current_text:
                continue
                
            placeholder = f"___PROTECTED_{i}___"
            
            if op == "替换":
                count = current_text.count(target)
                if count > 0:
                    current_text = current_text.replace(target, placeholder)
                    placeholders[placeholder] = replace
                    changes[rule_key] += count
                    modified = True
                
            elif op == "前方插入":
                prefix = replace
                result = []
                last_end = 0
                count = 0
                while True:
                    idx = current_text.find(target, last_end)
                    if idx == -1:
                        result.append(current_text[last_end:])
                        break
                    
                    preceding_text = current_text[:idx]
                    if preceding_text.endswith(prefix) or preceding_text.endswith(placeholder):
                        result.append(current_text[last_end:idx + len(target)])
                    else:
                        result.append(current_text[last_end:idx] + prefix + target)
                        count += 1
                    last_end = idx + len(target)
                
                temp_text = "".join(result)
                if temp_text != current_text:
                    current_text = temp_text
                    changes[rule_key] += count
                    modified = True
                current_text = current_text.replace(target, placeholder)
                placeholders[placeholder] = target
                
            elif op == "后方插入":
                suffix = replace
                result = []
                last_end = 0
                count = 0
                while True:
                    idx = current_text.find(target, last_end)
                    if idx == -1:
                        result.append(current_text[last_end:])
                        break
                    
                    after_idx = idx + len(target)
                    following_text = current_text[after_idx:]
                    if following_text.startswith(suffix) or following_text.startswith(placeholder):
                        result.append(current_text[last_end:after_idx])
                    else:
                        result.append(current_text[last_end:after_idx] + suffix)
                        count += 1
                    last_end = after_idx
                
                temp_text = "".join(result)
                if temp_text != current_text:
                    current_text = temp_text
                    changes[rule_key] += count
                    modified = True
                current_text = current_text.replace(target, placeholder)
                placeholders[placeholder] = target

        for ph, rep in placeholders.items():
            current_text = current_text.replace(ph, rep)

        # 如果执行了替换或插入，但计数是 0（可能是防重复保护拦截了），也视为未修改
        if not any(changes.values()):
            modified = False

        return current_text, modified, dict(changes)

    # ================= 文件遍历与调度 =================
    def process_files(self, folder):
        ext_list = SUPPORTED_EXTS[self.file_type.get()]
        total_files_scanned = 0
        modified_count = 0
        found_files_count = 0
        error_count = 0
        global_changes = defaultdict(int)
        error_details = []

        for dirpath, _, filenames in os.walk(folder):
            for filename in filenames:
                ext = os.path.splitext(filename)[1].lower()
                if ext not in ext_list: continue
                
                total_files_scanned += 1
                filepath = os.path.join(dirpath, filename)
                rel_path = os.path.relpath(filepath, folder)
                
                try:
                    changed = False
                    file_changes = {}
                    
                    if ext == '.docx':
                        changed, file_changes = self.process_docx(filepath)
                    elif ext == '.xlsx':
                        changed, file_changes = self.process_xlsx(filepath)
                    else:
                        changed, file_changes = self.process_txt(filepath)
                    
                    # 判断是纯查找模式还是修改模式
                    is_find_only = all(r["op"] == "查找" for r in self.rules)
                    
                    if file_changes:
                        if is_find_only:
                            found_files_count += 1
                            self.log(f"🔍 [命中] {rel_path}", "blue")
                        elif changed:
                            modified_count += 1
                            self.log(f"✅ [已修正] {rel_path}", "green")
                        
                        # 打印详细的修改明细
                        detail_str = " | ".join([f"{k}: {v}处" for k, v in file_changes.items()])
                        self.log(f"   └─ 明细: {detail_str}", "info")
                        
                        for k, v in file_changes.items():
                            global_changes[k] += v
                            
                except PermissionError:
                    error_count += 1
                    error_details.append((rel_path, "文件被占用"))
                    self.log(f"❌ [跳过] {rel_path} -> 文件被占用", "red")
                except Exception as e:
                    error_count += 1
                    error_details.append((rel_path, str(e)))
                    self.log(f"❌ [失败] {rel_path} -> {str(e)}", "red")

        # ================= 最终总结报表 =================
        self.log("\n" + "="*60, "blue")
        self.log(" 📊 批处理最终执行报告", "blue")
        self.log("="*60, "blue")
        self.log(f"  共检查文件数 : {total_files_scanned} 个")
        
        is_find_only = all(r["op"] == "查找" for r in self.rules)
        if is_find_only:
            self.log(f"  命中目标文件 : {found_files_count} 个", "blue")
        else:
            self.log(f"  成功修改文件 : {modified_count} 个", "green")
            
        self.log(f"  异常/失败数  : {error_count} 个", "red" if error_count > 0 else "info")
        
        self.log("\n  [操作执行汇总]")
        if global_changes:
            for rule_name, count in global_changes.items():
                self.log(f"   - {rule_name}: 共 {count} 处")
        else:
            self.log("   - 未触发任何规则")

        if error_details:
            self.log("\n  [失败文件列表]")
            for path, err in error_details:
                self.log(f"   - {path} : {err}", "red")
                
        self.log("="*60 + "\n", "blue")

        self.is_running = False
        self.btn_run.config(text="▶ 开始执行", state="normal")

    # ================= 文件类型处理 =================
    def process_txt(self, filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f: content = f.read()
        except UnicodeDecodeError:
            with open(filepath, 'r', encoding='gbk') as f: content = f.read()

        new_content, changed, changes = self.process_string(content, self.rules)

        if changed:
            with open(filepath, 'w', encoding='utf-8') as f: f.write(new_content)
        return changed, changes

    def process_docx(self, filepath):
        doc = docx.Document(filepath)
        is_changed = False
        file_changes = defaultdict(int)

        def process_paragraph(p):
            nonlocal is_changed
            orig_text = p.text
            new_text, changed, changes = self.process_string(orig_text, self.rules)
            
            if changes:
                for k, v in changes.items():
                    file_changes[k] += v
                    
            if changed:
                # 尝试修改 runs 以保留局部字体格式
                for rule in self.rules:
                    if rule["op"] in ("替换", "前方插入", "后方插入"):
                        for run in p.runs:
                            if rule["target"] in run.text:
                                run.text, _, _ = self.process_string(run.text, [rule])
                
                if "".join(r.text for r in p.runs) != new_text:
                    if len(p.runs) > 0:
                        p.runs[0].text = new_text
                        for run in p.runs[1:]:
                            run.text = ""
                    else:
                        p.text = new_text
                is_changed = True

        for p in doc.paragraphs: process_paragraph(p)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs: process_paragraph(p)

        if is_changed:
            doc.save(filepath)
            return True, dict(file_changes)
        return False, dict(file_changes)

    def process_xlsx(self, filepath):
        wb = openpyxl.load_workbook(filepath)
        is_changed = False
        file_changes = defaultdict(int)
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        new_val, changed, changes = self.process_string(cell.value, self.rules)
                        if changes:
                            for k, v in changes.items():
                                file_changes[k] += v
                        if changed:
                            cell.value = new_val
                            is_changed = True
        if is_changed:
            wb.save(filepath)
            return True, dict(file_changes)
        return False, dict(file_changes)

if __name__ == "__main__":
    root = tk.Tk()
    app = AdvancedBatchProcessor(root)
    root.mainloop()
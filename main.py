#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图片OCR识别 → 提取字段 → 填入Excel
支持Windows 10，中文识别优化
"""

import os
import re
import sys
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
from threading import Thread
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from PIL import Image, ImageTk

# ── PaddleOCR 延迟导入（避免首次启动慢） ──
ocr_engine = None
def get_ocr():
    global ocr_engine
    if ocr_engine is None:
        from paddleocr import PaddleOCR
        ocr_engine = PaddleOCR(use_angle_cls=True, lang='ch', show_log=False)
    return ocr_engine


# =═══════════════════════════════════════════════════════════
#  工具函数
# =═══════════════════════════════════════════════════════════

SUPPORTED_EXT = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp'}

def is_image_file(path):
    return Path(path).suffix.lower() in SUPPORTED_EXT


def run_ocr_on_image(image_path):
    """返回 OCR 识别出来的文本行列表"""
    ocr = get_ocr()
    result = ocr.ocr(str(image_path), cls=True)
    lines = []
    if result and result[0]:
        for line_info in result[0]:
            text = line_info[1][0]  # (text, confidence)
            confidence = line_info[1][1]
            lines.append((text.strip(), confidence))
    return lines


def parse_image_fileds(lines, field_names):
    """
    从 OCR 文本行中提取指定字段的值。
    支持格式： "姓名: 张三"  "姓名：张三"  "姓名 张三"  "姓名\t张三"
    返回 {field_name: value, ...}
    """
    # 先收集所有文本
    all_sentences = []
    for text, conf in lines:
        if conf > 0.3:  # 低置信度跳过
            all_sentences.append(text)

    full_text = " ".join(all_sentences)
    result = {}

    for field in field_names:
        # 尝试多种分隔符匹配
        # 模式1: 字段名 + : 或 ： + 空格 + 值
        pattern1 = re.compile(
            re.escape(field) + r'[\s]*[:：][\s]*(.+?)(?=[\s]*[：:\n]|\s{2,}|$)'
        )
        match = pattern1.search(full_text)
        if match:
            val = match.group(1).strip()
            if val and len(val) < 100:
                result[field] = val
                continue

        # 模式2: 字段名 + 空格/制表符 + 值（值不含空格或短）
        pattern2 = re.compile(
            re.escape(field) + r'[\s]{1,3}([^\s:：]{1,60})'
        )
        match = pattern2.search(full_text)
        if match:
            val = match.group(1).strip().rstrip('。，,.;:：')
            if val and not val.isascii() or (val.isascii() and len(val) > 2):
                result[field] = val
                continue

    return result


def extract_all_text(lines):
    """提取所有文本（用于预览）"""
    return [t for t, c in lines if c > 0.3]


# =═══════════════════════════════════════════════════════════
#  Excel 处理
# =═══════════════════════════════════════════════════════════

def create_excel_template(output_path, headers):
    """创建空的 Excel 模板（只有表头）"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "识别数据"
    # 写表头
    header_font = Font(bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
    ws.column_dimensions['A'].width = 12
    wb.save(output_path)
    return output_path


def append_row_to_excel(excel_path, field_to_column_map, row_data, total_processed):
    """
    field_to_column_map: { field_name: column_letter (如 'A'), ... }
    row_data: { field_name: value, ... }
    total_processed: 已处理图片数（决定行号）
    """
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    # 表头在第1行，数据从第2行开始
    data_row = total_processed + 2

    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    for field, col_letter in field_to_column_map.items():
        value = row_data.get(field, "（未识别）")
        col_idx = openpyxl.utils.column_index_from_string(col_letter.upper())
        cell = ws.cell(row=data_row, column=col_idx, value=value)
        cell.border = thin_border

    wb.save(excel_path)


def get_excel_headers(excel_path):
    """读取 Excel 表头（第一行）"""
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    headers = []
    for cell in ws[1]:
        if cell.value:
            headers.append(str(cell.value))
    return headers


# =═══════════════════════════════════════════════════════════
#  GUI 主程序
# =═══════════════════════════════════════════════════════════

class OCRExcelApp:
    def __init__(self, root):
        self.root = root
        root.title("图片文字识别 → Excel 自动填表工具")
        root.geometry("880x780")
        try:
            root.iconbitmap(default='')
        except:
            pass

        # 数据
        self.image_folder = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.fields = []           # [(field_name, column_letter)]
        self.image_files = []
        self.total_processed = 0
        self.is_processing = False

        self._build_ui()
        # 加载配置
        self.config_file = Path.home() / ".ocr_excel_tool_config.json"
        self._load_config()

    def _build_ui(self):
        # ── 风格 ──
        style = ttk.Style()
        style.theme_use('vista' if 'vista' in style.theme_names() else 'clam')
        style.configure('Accent.TButton', font=('微软雅黑', 11, 'bold'))

        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ── 标题 ──
        title_label = ttk.Label(
            main_frame,
            text="📄 图片识别 → Excel 自动填表",
            font=('微软雅黑', 16, 'bold')
        )
        title_label.pack(pady=(0, 15))

        # ── 第一步：选择图片文件夹 ──
        step1 = ttk.LabelFrame(main_frame, text="📁 第一步：图片文件夹", padding=10)
        step1.pack(fill=tk.X, pady=5)

        f1 = ttk.Frame(step1)
        f1.pack(fill=tk.X)
        ttk.Entry(f1, textvariable=self.image_folder, width=60).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(f1, text="浏览...", command=self._browse_folder, width=10).pack(side=tk.LEFT)
        ttk.Button(f1, text="扫描图片", command=self._scan_images, width=10).pack(side=tk.LEFT, padx=5)

        self.folder_status = ttk.Label(step1, text="（未选择文件夹）", foreground="gray")
        self.folder_status.pack(anchor=tk.W, pady=(5, 0))

        # ── 第二步：Excel 文件 ──
        step2 = ttk.LabelFrame(main_frame, text="📊 第二步：Excel 文件", padding=10)
        step2.pack(fill=tk.X, pady=5)

        f2 = ttk.Frame(step2)
        f2.pack(fill=tk.X)
        ttk.Entry(f2, textvariable=self.excel_path, width=60).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(f2, text="选择已有...", command=self._browse_excel_open, width=12).pack(side=tk.LEFT)
        ttk.Button(f2, text="新建模板", command=self._create_template, width=10).pack(side=tk.LEFT, padx=5)

        self.excel_status = ttk.Label(step2, text="（未选择）", foreground="gray")
        self.excel_status.pack(anchor=tk.W, pady=(5, 0))

        # ── 第三步：字段映射 ──
        step3 = ttk.LabelFrame(main_frame, text="🔤 第三步：字段映射（字段名 → Excel 列）", padding=10)
        step3.pack(fill=tk.BOTH, expand=True, pady=5)

        # 工具栏
        f3_toolbar = ttk.Frame(step3)
        f3_toolbar.pack(fill=tk.X, pady=(0, 5))
        ttk.Button(f3_toolbar, text="➕ 添加字段", command=self._add_field_row).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(f3_toolbar, text="📷 从样例图片识别字段名", command=self._auto_detect_fields).pack(side=tk.LEFT, padx=5)
        ttk.Button(f3_toolbar, text="🗑️ 删除选中行", command=self._delete_selected_field).pack(side=tk.LEFT, padx=5)
        ttk.Button(f3_toolbar, text="从 Excel 表头导入", command=self._import_from_excel).pack(side=tk.LEFT, padx=5)
        ttk.Label(f3_toolbar, text="  | 列字母如 A B C ...").pack(side=tk.LEFT, padx=10)

        # 字段映射表格
        columns = ('field_name', 'column_letter')
        self.field_tree = ttk.Treeview(step3, columns=columns, show='headings', height=8)
        self.field_tree.heading('field_name', text='字段名（OCR 中查找的文本）')
        self.field_tree.heading('column_letter', text='Excel 列 (A, B, C...)')
        self.field_tree.column('field_name', width=350, anchor=tk.W)
        self.field_tree.column('column_letter', width=150, anchor=tk.CENTER)

        # 滚动条
        vsb = ttk.Scrollbar(step3, orient=tk.VERTICAL, command=self.field_tree.yview)
        self.field_tree.configure(yscrollcommand=vsb.set)
        self.field_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # ── 第四步：日志 & 操作 ──
        step4 = ttk.LabelFrame(main_frame, text="📝 运行日志", padding=10)
        step4.pack(fill=tk.BOTH, pady=5)

        self.log_area = scrolledtext.ScrolledText(
            step4, height=10, font=('Consolas', 10),
            bg="#1e1e1e", fg="#d4d4d4", insertbackground="white"
        )
        self.log_area.pack(fill=tk.BOTH, expand=True)
        self.log_area.tag_config("info", foreground="#d4d4d4")
        self.log_area.tag_config("success", foreground="#4ec9b0")
        self.log_area.tag_config("error", foreground="#f44747")
        self.log_area.tag_config("warn", foreground="#ce9178")
        self.log_area.tag_config("title", foreground="#569cd6", font=('Consolas', 11, 'bold'))

        # ── 底部按钮 ──
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(5, 0))

        self.progress = ttk.Progressbar(bottom_frame, mode='determinate')
        self.progress.pack(fill=tk.X, pady=(0, 8))

        btn_frame = ttk.Frame(bottom_frame)
        btn_frame.pack(fill=tk.X)

        self.process_btn = ttk.Button(
            btn_frame, text="🚀 开始识别并填表", width=25,
            command=self._start_process, style='Accent.TButton'
        )
        self.process_btn.pack(side=tk.RIGHT, padx=5)

        ttk.Button(
            btn_frame, text="清空日志", width=10,
            command=lambda: self.log_area.delete('1.0', tk.END)
        ).pack(side=tk.RIGHT, padx=5)

        ttk.Button(
            btn_frame, text="打开输出目录", width=12,
            command=self._open_output_dir
        ).pack(side=tk.LEFT, padx=5)

    # ── UI 回调 ──

    def log(self, msg, tag="info"):
        self.log_area.insert(tk.END, msg + "\n", tag)
        self.log_area.see(tk.END)
        self.root.update_idletasks()

    def _browse_folder(self):
        folder = filedialog.askdirectory(title="选择图片文件夹")
        if folder:
            self.image_folder.set(folder)
            self.folder_status.config(text=f"已选择: {folder}", foreground="green")
            self._scan_images()

    def _scan_images(self):
        folder = self.image_folder.get()
        if not folder or not os.path.isdir(folder):
            return
        self.image_files = []
        for f in sorted(os.listdir(folder)):
            full = os.path.join(folder, f)
            if is_image_file(full):
                self.image_files.append(full)
        count = len(self.image_files)
        self.folder_status.config(
            text=f"已选择: {folder}  |  找到 {count} 张图片",
            foreground="green"
        )
        self.log(f"📂 扫描文件夹: {folder}  →  共 {count} 张图片", "info")

    def _browse_excel_open(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if path:
            self.excel_path.set(path)
            self.excel_status.config(text=f"已选择: {path}", foreground="green")
            self._import_from_excel()

    def _create_template(self):
        if not self.field_tree.get_children():
            messagebox.showwarning("提示", "请先添加至少一个字段映射再创建模板。")
            return

        path = filedialog.asksaveasfilename(
            title="保存 Excel 模板",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if not path:
            return

        fields = self._get_fields_from_tree()
        headers = [f[0] for f in fields]
        create_excel_template(path, headers)
        self.excel_path.set(path)
        self.excel_status.config(text=f"✅ 模板已创建: {path}", foreground="green")
        self.log(f"📋 Excel 模板创建成功: {os.path.basename(path)}", "success")

    def _add_field_row(self, field_name="", col_letter=""):
        # 弹窗输入
        dialog = tk.Toplevel(self.root)
        dialog.title("添加字段映射")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="字段名（OCR 文本中要查找的内容，如：姓名）：").pack(pady=(15, 5), padx=15, anchor=tk.W)
        name_var = tk.StringVar(value=field_name)
        ttk.Entry(dialog, textvariable=name_var, width=40).pack(padx=15, fill=tk.X)

        ttk.Label(dialog, text="对应 Excel 列（如 A、B、C）：").pack(pady=(10, 5), padx=15, anchor=tk.W)
        col_var = tk.StringVar(value=col_letter)
        ttk.Entry(dialog, textvariable=col_var, width=10).pack(padx=15, anchor=tk.W)

        def confirm():
            n = name_var.get().strip()
            c = col_var.get().strip().upper()
            if not n:
                messagebox.showwarning("提示", "字段名不能为空")
                return
            if not re.match(r'^[A-Z]{1,3}$', c):
                messagebox.showwarning("提示", "列名格式不对，请输入 A ~ ZZZ 之间的列字母")
                return
            # 检查重复
            for item in self.field_tree.get_children():
                vals = self.field_tree.item(item, 'values')
                if vals[1] == c:
                    if not messagebox.askyesno("确认", f"列 {c} 已有字段 {vals[0]}，确定要覆盖？"):
                        return
                    self.field_tree.delete(item)
                    break
            self.field_tree.insert('', tk.END, values=(n, c))
            dialog.destroy()

        ttk.Button(dialog, text="确认添加", command=confirm).pack(pady=15)
        dialog.bind('<Return>', lambda e: confirm())

    def _delete_selected_field(self):
        selected = self.field_tree.selection()
        for item in selected:
            self.field_tree.delete(item)

    def _get_fields_from_tree(self):
        result = []
        for item in self.field_tree.get_children():
            vals = self.field_tree.item(item, 'values')
            result.append((vals[0], vals[1]))
        return result

    def _import_from_excel(self):
        path = self.excel_path.get()
        if not path or not os.path.isfile(path):
            return
        try:
            headers = get_excel_headers(path)
        except Exception as e:
            self.log(f"❌ 读取 Excel 表头失败: {e}", "error")
            return

        # 清空当前，导入表头作为字段名，列从 A 开始
        for item in self.field_tree.get_children():
            self.field_tree.delete(item)

        col_letters = []
        for i in range(len(headers)):
            # 生成列字母 A, B, C, ..., Z, AA, AB...
            col = ""
            n = i + 1
            while n > 0:
                n -= 1
                col = chr(65 + n % 26) + col
                n //= 26
            col_letters.append(col)

        for h, c in zip(headers, col_letters):
            self.field_tree.insert('', tk.END, values=(h, c))

        self.log(f"📋 已从 Excel 导入 {len(headers)} 个字段映射", "success")

    def _auto_detect_fields(self):
        """用第一张图片 OCR 识别，让用户从中选字段名"""
        if not self.image_files:
            messagebox.showwarning("提示", "请先选择一个包含图片的文件夹并扫描。")
            return

        sample = self.image_files[0]
        self.log(f"🔍 正在 OCR 识别样例图片: {os.path.basename(sample)}", "info")
        self.root.update_idletasks()

        lines = run_ocr_on_image(sample)
        texts = extract_all_text(lines)
        self.log(f"📝 识别到 {len(texts)} 段文本: ", "info")
        for t in texts[:30]:
            self.log(f"   · {t}", "info")

        # 弹窗让用户选择哪些作为字段名
        dialog = tk.Toplevel(self.root)
        dialog.title("选择字段名")
        dialog.geometry("550x450")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="勾选要作为字段名的文本（系统会自动匹配其后的值）：").pack(pady=10, padx=10, anchor=tk.W)

        frame = ttk.Frame(dialog)
        frame.pack(fill=tk.BOTH, expand=True, padx=10)

        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        vars_list = []
        for t in texts:
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(scrollable_frame, text=t, variable=var)
            chk.pack(anchor=tk.W, padx=5, pady=2)
            vars_list.append((t, var))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def confirm_fields():
            selected = []
            col_idx = 0
            for text, var in vars_list:
                if var.get():
                    # 去除冒号等后缀
                    clean = text.rstrip(':： ')
                    col = ""
                    n = col_idx + 1
                    while n > 0:
                        n -= 1
                        col = chr(65 + n % 26) + col
                        n //= 26
                    selected.append((clean, col))
                    col_idx += 1
            if not selected:
                messagebox.showwarning("提示", "至少选择一个字段")
                return

            # 检查覆盖
            existing = self._get_fields_from_tree()
            if existing:
                if not messagebox.askyesno("确认", f"当前已有 {len(existing)} 个字段映射，确定要替换吗？"):
                    return
            for item in self.field_tree.get_children():
                self.field_tree.delete(item)
            for n, c in selected:
                self.field_tree.insert('', tk.END, values=(n, c))
            dialog.destroy()
            self.log(f"✅ 已添加 {len(selected)} 个字段映射", "success")

        ttk.Button(dialog, text="确认选择", command=confirm_fields).pack(pady=10)

    # ── 核心处理逻辑 ──

    def _start_process(self):
        if self.is_processing:
            return

        # 校验
        folder = self.image_folder.get()
        excel = self.excel_path.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("提示", "请选择有效的图片文件夹")
            return
        if not excel:
            messagebox.showwarning("提示", "请选择或创建 Excel 文件")
            return
        fields = self._get_fields_from_tree()
        if not fields:
            messagebox.showwarning("提示", "请添加至少一个字段映射")
            return

        if not self.image_files:
            self._scan_images()
        if not self.image_files:
            messagebox.showwarning("提示", "文件夹中没有找到图片")
            return

        # 如果 Excel 不存在，用当前字段创建模板
        if not os.path.isfile(excel):
            headers = [f[0] for f in fields]
            create_excel_template(excel, headers)
            self.log(f"📋 自动创建 Excel 模板: {os.path.basename(excel)}", "info")

        self.is_processing = True
        self.process_btn.config(state=tk.DISABLED, text="⏳ 处理中...")
        self.progress['maximum'] = len(self.image_files)
        self.progress['value'] = 0
        self.total_processed = 0
        self.log("=" * 55, "title")
        self.log("🚀 开始批量处理...", "title")
        self.log("=" * 55, "title")

        # 后台线程执行
        Thread(target=self._process_all, daemon=True).start()

    def _process_all(self):
        try:
            fields = self._get_fields_from_tree()
            field_names = [f[0] for f in fields]
            field_to_col = {f[0]: f[1] for f in fields}
            excel = self.excel_path.get()

            for idx, img_path in enumerate(self.image_files):
                if not self.is_processing:
                    break

                name = os.path.basename(img_path)
                self.log(f"\n[{idx+1}/{len(self.image_files)}] 📷 {name}", "info")

                try:
                    lines = run_ocr_on_image(img_path)
                    all_texts = extract_all_text(lines)
                    self.log(f"   OCR 识别到 {len(all_texts)} 段文本", "info")

                    row_data = parse_image_fileds(lines, field_names)
                    identified = {k for k, v in row_data.items() if v}
                    self.log(f"   ✅ 识别到 {len(identified)}/{len(field_names)} 个字段: {identified}", "success")

                    for fn in field_names:
                        val = row_data.get(fn, "（未识别）")
                        self.log(f"      · {fn} → {val}", "info")

                    if row_data:
                        append_row_to_excel(excel, field_to_col, row_data, self.total_processed)
                        self.total_processed += 1
                        file_name = os.path.splitext(name)[0]
                        self.log(f"   💾 已写入 Excel (行 {self.total_processed+1})", "success")
                    else:
                        self.log(f"   ⚠️ 未识别到任何字段，跳过", "warn")

                except Exception as e:
                    self.log(f"   ❌ 处理失败: {e}", "error")

                self.progress['value'] = idx + 1
                self.root.update_idletasks()

            self.log("=" * 55, "title")
            if self.total_processed > 0:
                self.log(f"✅ 处理完成！共处理 {len(self.image_files)} 张图片，"
                         f"成功写入 {self.total_processed} 行数据到 Excel", "success")
                self.log(f"📁 输出文件: {self.excel_path.get()}", "success")
            else:
                self.log("⚠️ 未写入任何数据，请检查字段映射和图片内容", "warn")
            self.log("=" * 55, "title")

        except Exception as e:
            self.log(f"❌ 处理过程异常: {e}", "error")
        finally:
            self.is_processing = False
            self.root.after(0, self._process_done)

    def _process_done(self):
        self.process_btn.config(state=tk.NORMAL, text="🚀 开始识别并填表")
        self.progress['value'] = 0
        self._save_config()

    def _open_output_dir(self):
        path = self.excel_path.get()
        if path and os.path.isfile(path):
            dirpath = os.path.dirname(path)
            os.startfile(dirpath)

    # ── 配置持久化 ──

    def _save_config(self):
        try:
            fields = self._get_fields_from_tree()
            config = {
                'image_folder': self.image_folder.get(),
                'excel_path': self.excel_path.get(),
                'fields': fields,
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _load_config(self):
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                if config.get('image_folder'):
                    self.image_folder.set(config['image_folder'])
                    self._scan_images()
                if config.get('excel_path'):
                    p = config['excel_path']
                    if os.path.isfile(p):
                        self.excel_path.set(p)
                        self.excel_status.config(text=f"已选择: {p}", foreground="green")
                fields = config.get('fields', [])
                for f in fields:
                    if len(f) >= 2:
                        self.field_tree.insert('', tk.END, values=(f[0], f[1]))
        except Exception:
            pass


# =═══════════════════════════════════════════════════════════
#  入口
# =═══════════════════════════════════════════════════════════

def main():
    root = tk.Tk()
    app = OCRExcelApp(root)
    root.mainloop()


if __name__ == "__main__":
    # 首次运行提示安装依赖
    try:
        import paddleocr
    except ImportError:
        print("=" * 55)
        print("📦 首次使用请先安装依赖，运行：")
        print("   pip install -r requirements.txt")
        print("=" * 55)
        sys.exit(1)
    main()

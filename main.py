#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标本照片OCR识别 → Excel自动填表工具
适用于：采集号、采集时间、采集人、采集地点、经度、纬度、海拔、习性、生态环境、高度
Windows 10/11 后台自动运行
"""

import os, re, sys, json, time, threading, platform
from pathlib import Path
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
except ImportError:
    tk = None

# ═══════ 字段配置（硬编码，用户无需修改） ═══════
FIELDS = [
    "采集号", "采集时间", "采集人", "采集地点",
    "经度", "纬度", "海拔",
    "习性", "生态环境", "高度"
]

SUPPORTED_EXT = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp'}
CONFIG_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "标本OCR工具"
CONFIG_FILE = CONFIG_DIR / "config.json"

# ═══════ OCR 引擎（懒加载） ═══════
_ocr = None
def init_ocr():
    """首次调用时下载并加载模型，后续直接返回"""
    global _ocr
    if _ocr is not None:
        return _ocr
    # PaddleOCR 首次会自动下载模型（约15MB），静默下载
    from paddleocr import PaddleOCR
    _ocr = PaddleOCR(use_angle_cls=True, lang='ch', show_log=False, use_gpu=False)
    return _ocr

def is_img(path):
    return Path(path).suffix.lower() in SUPPORTED_EXT

# ─── OCR 识别 ───
def ocr_image(img_path):
    """返回置信度 >0.3 的文本行列表"""
    ocr = init_ocr()
    res = ocr.ocr(str(img_path), cls=True)
    texts = []
    if res and res[0]:
        for box_info in res[0]:
            txt, conf = box_info[1][0].strip(), box_info[1][1]
            if conf > 0.3 and txt:
                texts.append(txt)
    return texts

def parse_fields(texts):
    """
    从 OCR 文本中提取各字段值。
    支持格式：字段名:值 / 字段名：值 / 字段名 值（值无空格时）
    """
    full = "\n".join(texts)
    result = {}

    for field in FIELDS:
        val = ""
        # 模式A：字段名:值 或 字段名：值
        m = re.search(
            re.escape(field) + r'[\s]*[:：]\s*([^\n]{1,200})',
            full
        )
        if m:
            v = m.group(1).strip().rstrip('，。.;,;')
            if v:
                val = v
        else:
            # 模式B：字段名后紧跟值（无分隔符或仅空格、tab）
            m = re.search(
                re.escape(field) + r'[\t ]{1,4}([^\s]{1,100})',
                full
            )
            if m:
                v = m.group(1).strip().rstrip('，。.;,;')
                if v:
                    val = v
        result[field] = val
    return result

# ─── Excel 操作 ───
def create_excel(path):
    """创建带格式的模板"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "标本数据"
    hf = Font(bold=True, size=11, color="FFFFFF")
    hfill = PatternFill("solid", fgColor="4472C4")
    halign = Alignment(horizontal="center", vertical="center", wrap_text=True)
    bdr = Border(left=Side('thin'), right=Side('thin'),
                 top=Side('thin'), bottom=Side('thin'))
    headers = ["序号"] + FIELDS
    widths = [8] + [12, 16, 12, 20, 14, 14, 10, 16, 16, 10]
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font, c.fill, c.alignment, c.border = hf, hfill, halign, bdr
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    wb.save(path)
    return path

def write_row(excel_path, row_data, row_num):
    """row_num: Excel 行号，1-based。表头=1，数据从2开始"""
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    bdr = Border(left=Side('thin'), right=Side('thin'),
                 top=Side('thin'), bottom=Side('thin'))
    # 序号列
    c = ws.cell(row=row_num, column=1, value=row_num - 1)
    c.border = bdr
    for i, field in enumerate(FIELDS, 2):
        v = row_data.get(field, "") or "（未识别）"
        c = ws.cell(row=row_num, column=i, value=v)
        c.border = bdr
    wb.save(excel_path)

def count_data_rows(excel_path):
    """返回已有数据行数（不含表头）"""
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    return max(0, ws.max_row - 1)

# ─── 核心处理函数 ───
def process_folder(image_folder, excel_path, progress_cb=None, log_cb=None):
    """
    处理一个文件夹内所有图片，OCR 后填入 Excel。
    progress_cb(current, total): 进度回调
    log_cb(msg): 日志回调
    返回 (处理数, 成功数)
    """
    # 扫描图片
    images = sorted([
        p for p in Path(image_folder).iterdir()
        if is_img(p)
    ])
    total = len(images)
    if total == 0:
        if log_cb:
            log_cb("⚠️ 未找到图片文件")
        return 0, 0

    # 创建 Excel（如果不存在）
    if not os.path.isfile(excel_path):
        create_excel(excel_path)
        if log_cb:
            log_cb(f"📋 已创建 Excel 模板: {Path(excel_path).name}")

    # 计算已有行数（追加模式）
    base_rows = count_data_rows(excel_path)  # 已有数据行数
    succeed = 0

    for idx, img in enumerate(images):
        name = img.name
        if log_cb:
            log_cb(f"[{idx+1}/{total}] 📷 {name}")
        if progress_cb:
            progress_cb(idx + 1, total)

        try:
            texts = ocr_image(str(img))
            if not texts:
                if log_cb:
                    log_cb(f"   ⚠️ 未识别到文字，跳过")
                continue

            row_data = parse_fields(texts)
            found = {k for k, v in row_data.items() if v}
            if log_cb:
                for f in FIELDS:
                    v = row_data.get(f, "") or "（未识别）"
                    log_cb(f"   · {f} → {v}")

            # 写入 Excel
            row_num = base_rows + succeed + 2  # +1 表头，+1 已有数据
            write_row(excel_path, row_data, row_num)
            succeed += 1
            if log_cb:
                log_cb(f"   ✅ 已写入 → 行{row_num}")

        except Exception as e:
            if log_cb:
                log_cb(f"   ❌ 处理失败: {e}")

    if log_cb:
        log_cb(f"\n{'='*45}")
        log_cb(f"✅ 处理完成！共 {total} 张图片，成功写入 {succeed} 行")
        log_cb(f"📁 输出: {excel_path}")
    return total, succeed


# ═════════════════════════════════════════════
#  GUI 界面（轻量，专为 Windows 优化）
# ═════════════════════════════════════════════

class App:
    def __init__(self):
        if tk is None:
            print("❌ 错误：需要 tkinter 支持（Windows 自带）")
            sys.exit(1)

        self.root = tk.Tk()
        self.root.title("标本照片 → Excel 自动填表")
        self.root.geometry("720x620")
        self.root.minsize(600, 500)

        # 尝试设置图标（内置图标）
        try:
            self.root.iconbitmap(default="")
        except:
            pass

        self.image_folder = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.running = False
        self._load_config()
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()

    def _build_ui(self):
        root = self.root
        main = ttk.Frame(root, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        # ── 标题 ──
        title = ttk.Label(main, text="标本照片 OCR → Excel 自动填表",
                          font=('微软雅黑', 14, 'bold'))
        title.pack(pady=(0, 8))

        # ── 图片文件夹 ──
        f1 = ttk.LabelFrame(main, text="📁 图片文件夹（按文件名顺序处理）", padding=8)
        f1.pack(fill=tk.X, pady=4)
        row1 = ttk.Frame(f1); row1.pack(fill=tk.X)
        ttk.Entry(row1, textvariable=self.image_folder, width=55).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(row1, text="浏览...", command=self._pick_folder, width=8).pack(side=tk.LEFT)
        self._folder_status = ttk.Label(f1, foreground="gray")
        self._folder_status.pack(anchor=tk.W, pady=(3, 0))

        # ── Excel 文件 ──
        f2 = ttk.LabelFrame(main, text="📊 Excel 输出文件", padding=8)
        f2.pack(fill=tk.X, pady=4)
        row2 = ttk.Frame(f2); row2.pack(fill=tk.X)
        ttk.Entry(row2, textvariable=self.excel_path, width=55).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(row2, text="选择已有...", command=self._pick_excel, width=10).pack(side=tk.LEFT)
        ttk.Button(row2, text="新建...", command=self._new_excel, width=8).pack(side=tk.LEFT, padx=4)
        self._excel_status = ttk.Label(f2, foreground="gray")
        self._excel_status.pack(anchor=tk.W, pady=(3, 0))

        # 自动刷新状态
        self.image_folder.trace_add("write", lambda *a: self._refresh_status())
        self.excel_path.trace_add("write", lambda *a: self._refresh_status())
        self._refresh_status()

        # ── 日志 ──
        f3 = ttk.LabelFrame(main, text="📝 运行日志", padding=6)
        f3.pack(fill=tk.BOTH, expand=True, pady=4)

        self.log = scrolledtext.ScrolledText(
            f3, height=12, font=('Consolas', 10),
            bg="#1e1e1e", fg="#d4d4d4", insertbackground="white",
            state=tk.DISABLED
        )
        self.log.pack(fill=tk.BOTH, expand=True)
        self.log.tag_config("info", foreground="#d4d4d4")
        self.log.tag_config("ok", foreground="#4ec9b0")
        self.log.tag_config("err", foreground="#f44747")
        self.log.tag_config("warn", foreground="#ce9178")
        self.log.tag_config("bold", foreground="#569cd6",
                            font=('Consolas', 11, 'bold'))

        # ── 进度条 ──
        self.progress = ttk.Progressbar(main, mode='determinate')
        self.progress.pack(fill=tk.X, pady=2)
        self.progress_label = ttk.Label(main, text="", foreground="gray")
        self.progress_label.pack()

        # ── 操作按钮 ──
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=(4, 0))
        self.run_btn = ttk.Button(
            btn_frame, text="🚀 开始识别并填表", width=24,
            command=self._start_process
        )
        self.run_btn.pack(side=tk.RIGHT, padx=4)
        ttk.Button(btn_frame, text="清空日志",
                   command=self._clear_log).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btn_frame, text="打开输出目录",
                   command=self._open_dir).pack(side=tk.LEFT, padx=4)

    # ── UI 工具 ──
    def _wlog(self, msg, tag="info"):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, msg + "\n", tag)
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def _clear_log(self):
        self.log.config(state=tk.NORMAL)
        self.log.delete('1.0', tk.END)
        self.log.config(state=tk.DISABLED)

    def _refresh_status(self):
        img_dir = self.image_folder.get()
        exl = self.excel_path.get()
        if img_dir and os.path.isdir(img_dir):
            n = len([p for p in Path(img_dir).iterdir() if is_img(p)])
            self._folder_status.config(
                text=f"已选（{n} 张图片）", foreground="green")
        else:
            self._folder_status.config(text="（未选择）", foreground="gray")

        if exl and os.path.isfile(exl):
            self._excel_status.config(text="✅ 文件存在", foreground="green")
        elif exl:
            self._excel_status.config(text="⚠️ 不存在，处理时将自动创建", foreground="orange")
        else:
            self._excel_status.config(text="（未选择）", foreground="gray")

    def _pick_folder(self):
        d = filedialog.askdirectory(title="选择标本照片文件夹")
        if d:
            self.image_folder.set(d)

    def _pick_excel(self):
        p = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if p:
            self.excel_path.set(p)

    def _new_excel(self):
        p = filedialog.asksaveasfilename(
            title="保存 Excel 文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if p:
            create_excel(p)
            self.excel_path.set(p)
            self._wlog(f"📋 已创建模板: {Path(p).name}", "ok")

    def _open_dir(self):
        p = self.excel_path.get()
        if p and os.path.isfile(p):
            os.startfile(os.path.dirname(p))

    def _on_close(self):
        self._save_config()
        self.root.destroy()

    # ── 核心逻辑 ──
    def _start_process(self):
        if self.running:
            return

        img_dir = self.image_folder.get()
        exl = self.excel_path.get()
        if not img_dir or not os.path.isdir(img_dir):
            messagebox.showwarning("提示", "请选择有效的图片文件夹")
            return
        if not exl:
            messagebox.showwarning("提示", "请选择或新建 Excel 文件")
            return

        self.running = True
        self.run_btn.config(state=tk.DISABLED, text="⏳ 处理中...")
        self.progress['value'] = 0
        self.progress_label.config(text="正在初始化 OCR 引擎（首次会下载模型，请稍候）...")
        self._wlog("=" * 45, "bold")
        self._wlog("🚀 开始处理标本照片...", "bold")
        self._wlog("=" * 45, "bold")

        def worker():
            def progress_cb(cur, tot):
                self.root.after(0, lambda: self.progress.configure(value=cur, maximum=tot))
                self.root.after(0, lambda: self.progress_label.config(
                    text=f"处理中 {cur}/{tot}  [{cur*100//tot}%]"))

            def log_cb(msg):
                tag = "info"
                if "✅" in msg or "写入" in msg:
                    tag = "ok"
                elif "⚠️" in msg:
                    tag = "warn"
                elif "❌" in msg:
                    tag = "err"
                self.root.after(0, lambda m=msg, t=tag: self._wlog(m, t))

            try:
                process_folder(img_dir, exl, progress_cb, log_cb)
            except Exception as e:
                self.root.after(0, lambda: self._wlog(f"❌ 处理异常: {e}", "err"))
            finally:
                self.running = False
                self.root.after(0, self._on_done)

        self.root.after(100, lambda: threading.Thread(target=worker, daemon=True).start())

    def _on_done(self):
        self.run_btn.config(state=tk.NORMAL, text="🚀 开始识别并填表")
        self.progress_label.config(text="✅ 完成")
        self._save_config()
        # 完成后弹窗提示
        self.root.after(200, lambda: messagebox.showinfo(
            "完成", f"处理完毕！\n输出文件：{self.excel_path.get()}"))

    # ── 配置持久化 ──
    def _save_config(self):
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump({
                    "image_folder": self.image_folder.get(),
                    "excel_path": self.excel_path.get(),
                }, f)
        except Exception:
            pass

    def _load_config(self):
        try:
            if CONFIG_FILE.exists():
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    d = json.load(f)
                if d.get("image_folder") and os.path.isdir(d["image_folder"]):
                    self.image_folder.set(d["image_folder"])
                if d.get("excel_path"):
                    self.excel_path.set(d["excel_path"])
        except Exception:
            pass


# ═════════════════════════════════════════════
#  入口
# ═════════════════════════════════════════════

if __name__ == "__main__":
    try:
        import paddleocr
    except ImportError:
        print("=" * 50)
        print("📦 首次使用请先安装依赖，运行：")
        print("   pip install -r requirements.txt")
        print("=" * 50)
        sys.exit(1)

    App()

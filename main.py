#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标本照片 → Excel 自动填表
基于 Surya OCR（轻量级，纯 Python 无额外系统依赖）
"""

import os, re, sys, json, threading
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
except ImportError:
    tk = None

FIELDS = [
    "采集号", "采集时间", "采集人", "采集地点",
    "经度", "纬度", "海拔",
    "习性", "生态环境", "高度"
]

SUPPORTED_EXT = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp'}
CONFIG_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "标本OCR工具"
CONFIG_FILE = CONFIG_DIR / "config.json"


# ─── OCR 引擎 ───
_ocr = None
def get_ocr():
    global _ocr
    if _ocr is None:
        from rapidocr_onnxruntime import RapidOCR
        _ocr = RapidOCR()
    return _ocr

def is_img(path):
    return Path(path).suffix.lower() in SUPPORTED_EXT

def ocr_image(img_path):
    ocr = get_ocr()
    result, _ = ocr(str(img_path))
    texts = []
    if result:
        for box in result:
            txt = box[1].strip()
            if txt:
                texts.append(txt)
    return texts


def parse_fields(texts):
    full = "\n".join(texts)
    result = {}
    for field in FIELDS:
        val = ""
        m = re.search(re.escape(field) + r'[\s]*[:：]\s*([^\n]{1,200})', full)
        if m:
            v = m.group(1).strip().rstrip('，。.;,;')
            if v: val = v
        else:
            m = re.search(re.escape(field) + r'[\t ]{1,4}([^\s]{1,100})', full)
            if m:
                v = m.group(1).strip().rstrip('，。.;,;')
                if v: val = v
        result[field] = val
    return result


def create_excel(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "标本数据"
    hf = Font(bold=True, size=11, color="FFFFFF")
    hfill = PatternFill("solid", fgColor="4472C4")
    halign = Alignment(horizontal="center", vertical="center", wrap_text=True)
    bdr = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
    headers = ["序号"] + FIELDS
    widths = [8, 12, 16, 12, 20, 14, 14, 10, 16, 16, 10]
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font, c.fill, c.alignment, c.border = hf, hfill, halign, bdr
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'
    wb.save(path)
    return path


def write_row(excel_path, row_data, row_num):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    bdr = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
    c = ws.cell(row=row_num, column=1, value=row_num - 1); c.border = bdr
    for i, field in enumerate(FIELDS, 2):
        v = row_data.get(field, "") or ""
        c = ws.cell(row=row_num, column=i, value=v); c.border = bdr
    wb.save(excel_path)


def count_data_rows(excel_path):
    wb = openpyxl.load_workbook(excel_path)
    return max(0, wb.active.max_row - 1)


def process_folder(image_folder, excel_path, progress_cb=None, log_cb=None):
    images = sorted([p for p in Path(image_folder).iterdir() if is_img(p)])
    total = len(images)
    if total == 0:
        if log_cb: log_cb("⚠️ 未找到图片文件")
        return 0, 0
    if not os.path.isfile(excel_path):
        create_excel(excel_path)
        if log_cb: log_cb(f"📋 已创建 Excel 模板: {Path(excel_path).name}")
    base_rows = count_data_rows(excel_path)
    succeed = 0
    for idx, img in enumerate(images):
        name = img.name
        if log_cb: log_cb(f"[{idx+1}/{total}] 📷 {name}")
        if progress_cb: progress_cb(idx + 1, total)
        try:
            texts = ocr_image(str(img))
            if not texts:
                if log_cb: log_cb(f"   ⚠️ 未识别到文字")
                continue
            row_data = parse_fields(texts)
            for f in FIELDS:
                v = row_data.get(f, "") or "（未识别）"
                if log_cb: log_cb(f"   · {f} → {v}")
            row_num = base_rows + succeed + 2
            write_row(excel_path, row_data, row_num)
            succeed += 1
            if log_cb: log_cb(f"   ✅ 已写入 → 行{row_num}")
        except Exception as e:
            if log_cb: log_cb(f"   ❌ 处理失败: {e}")
    if log_cb:
        log_cb(f"\n{'='*45}")
        log_cb(f"✅ 完成！共 {total} 张图片，成功写入 {succeed} 行")
    return total, succeed


class App:
    def __init__(self):
        if tk is None:
            print("❌ 需要 tkinter（Windows 自带）")
            sys.exit(1)
        self.root = tk.Tk()
        self.root.title("标本照片 → Excel 自动填表")
        self.root.geometry("720x620")
        self.root.minsize(600, 500)
        self.image_folder = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.running = False
        self._load_config()
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()

    def _build_ui(self):
        m = ttk.Frame(self.root, padding=12); m.pack(fill=tk.BOTH, expand=True)
        ttk.Label(m, text="标本照片 OCR → Excel 自动填表", font=('微软雅黑',14,'bold')).pack(pady=(0,8))

        f1 = ttk.LabelFrame(m, text="📁 图片文件夹（按文件名排序处理）", padding=8); f1.pack(fill=tk.X, pady=4)
        r1 = ttk.Frame(f1); r1.pack(fill=tk.X)
        ttk.Entry(r1, textvariable=self.image_folder, width=55).pack(side=tk.LEFT, padx=(0,6))
        ttk.Button(r1, text="浏览...", command=self._pick_folder, width=8).pack(side=tk.LEFT)
        self.fs = ttk.Label(f1, foreground="gray"); self.fs.pack(anchor=tk.W, pady=(3,0))

        f2 = ttk.LabelFrame(m, text="📊 Excel 输出文件", padding=8); f2.pack(fill=tk.X, pady=4)
        r2 = ttk.Frame(f2); r2.pack(fill=tk.X)
        ttk.Entry(r2, textvariable=self.excel_path, width=55).pack(side=tk.LEFT, padx=(0,6))
        ttk.Button(r2, text="选择已有...", command=self._pick_excel, width=10).pack(side=tk.LEFT)
        ttk.Button(r2, text="新建...", command=self._new_excel, width=8).pack(side=tk.LEFT, padx=4)
        self.es = ttk.Label(f2, foreground="gray"); self.es.pack(anchor=tk.W, pady=(3,0))

        self.image_folder.trace_add("write", lambda *a: self._refresh())
        self.excel_path.trace_add("write", lambda *a: self._refresh())
        self._refresh()

        f3 = ttk.LabelFrame(m, text="📝 运行日志", padding=6); f3.pack(fill=tk.BOTH, expand=True, pady=4)
        self.log = scrolledtext.ScrolledText(f3, height=12, font=('Consolas',10),
            bg="#1e1e1e", fg="#d4d4d4", insertbackground="white", state=tk.DISABLED)
        self.log.pack(fill=tk.BOTH, expand=True)
        for tag, color in [("i","#d4d4d4"),("ok","#4ec9b0"),("err","#f44747"),("w","#ce9178"),("b","#569cd6")]:
            self.log.tag_config(tag, foreground=color)
        self.log.tag_config("b", font=('Consolas',11,'bold'))

        self.pb = ttk.Progressbar(m, mode='determinate'); self.pb.pack(fill=tk.X, pady=2)
        self.pl = ttk.Label(m, text="", foreground="gray"); self.pl.pack()

        bf = ttk.Frame(m); bf.pack(fill=tk.X, pady=(4,0))
        self.run_btn = ttk.Button(bf, text="🚀 开始识别并填表", width=24, command=self._start)
        self.run_btn.pack(side=tk.RIGHT, padx=4)
        ttk.Button(bf, text="清空日志", command=lambda: (self.log.config(state=tk.NORMAL),
            self.log.delete('1.0', tk.END), self.log.config(state=tk.DISABLED))).pack(side=tk.RIGHT, padx=4)
        ttk.Button(bf, text="打开输出目录", command=self._open_dir).pack(side=tk.LEFT, padx=4)

    def _wl(self, msg, tag="i"):
        self.log.config(state=tk.NORMAL); self.log.insert(tk.END, msg+"\n", tag)
        self.log.see(tk.END); self.log.config(state=tk.DISABLED); self.root.update_idletasks()

    def _refresh(self):
        d, e = self.image_folder.get(), self.excel_path.get()
        n = len([p for p in Path(d).iterdir() if is_img(p)]) if d and os.path.isdir(d) else 0
        self.fs.config(text=f"已选（{n} 张图片）" if n else "（未选择）",
                       foreground="green" if n else "gray")
        self.es.config(text="✅ 文件存在" if e and os.path.isfile(e) else
                       ("⚠️ 不存在，自动创建" if e else "（未选择）"),
                       foreground="green" if e and os.path.isfile(e) else ("orange" if e else "gray"))

    def _pick_folder(self): d = filedialog.askdirectory(); d and self.image_folder.set(d)
    def _pick_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx")]); p and self.excel_path.set(p)
    def _new_excel(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")])
        if p: create_excel(p); self.excel_path.set(p); self._wl(f"📋 已创建: {Path(p).name}", "ok")
    def _open_dir(self):
        p = self.excel_path.get()
        p and os.path.isfile(p) and os.startfile(os.path.dirname(p))
    def _on_close(self): self._save_config(); self.root.destroy()

    def _start(self):
        if self.running: return
        d, e = self.image_folder.get(), self.excel_path.get()
        if not d or not os.path.isdir(d): messagebox.showwarning("提示", "请选择图片文件夹"); return
        if not e: messagebox.showwarning("提示", "请选择或新建 Excel"); return
        self.running = True
        self.run_btn.config(state=tk.DISABLED, text="⏳ 处理中...")
        self.pb['value'] = 0; self.pl.config(text="正在初始化 OCR 引擎...")
        self._wl("="*45,"b"); self._wl("🚀 开始处理标本照片...","b"); self._wl("="*45,"b")
        def worker():
            def pc(c,t): self.root.after(0,lambda: (self.pb.configure(value=c,maximum=t),
                self.pl.config(text=f"处理中 {c}/{t} [{c*100//t}%]")))
            def lc(m,t="i"): self.root.after(0,lambda: self._wl(m,t))
            try: process_folder(d,e,pc,lc)
            except Exception as ex: self.root.after(0,lambda: self._wl(f"❌ {ex}","err"))
            finally: self.running=False; self.root.after(0,self._done)
        threading.Thread(target=worker, daemon=True).start()

    def _done(self):
        self.run_btn.config(state=tk.NORMAL, text="🚀 开始识别并填表")
        self.pl.config(text="✅ 完成"); self._save_config()
    def _save_config(self):
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            Path(CONFIG_FILE).write_text(json.dumps(
                {"image_folder": self.image_folder.get(), "excel_path": self.excel_path.get()}),
                encoding='utf-8')
        except: pass
    def _load_config(self):
        try:
            if CONFIG_FILE.exists():
                d = json.loads(CONFIG_FILE.read_text(encoding='utf-8'))
                d.get("image_folder") and os.path.isdir(d["image_folder"]) and self.image_folder.set(d["image_folder"])
                d.get("excel_path") and self.excel_path.set(d["excel_path"])
        except: pass


if __name__ == "__main__":
    try:
        from rapidocr_onnxruntime import RapidOCR
    except ImportError:
        print("="*50)
        print("📦 请先安装依赖：")
        print("   pip install rapidocr-onnxruntime openpyxl Pillow")
        print("="*50)
        sys.exit(1)
    App()

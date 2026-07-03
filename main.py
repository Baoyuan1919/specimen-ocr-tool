#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标本照片 → Excel 自动填表 v2.0
- 字段自定义增删（设置中自由增减，也可从 Excel 表头同步）
- 系统托盘后台运行（关闭窗口最小化到托盘区）
- 文件夹自动监听（检测新图片自动识别填表）
"""

import os, re, sys, json, threading, time
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
except ImportError:
    tk = None

# 默认字段
DEFAULT_FIELDS = [
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

def parse_fields(texts, fields):
    full = "\n".join(texts)
    result = {}
    for field in fields:
        val = ""
        m = re.search(re.escape(field) + r'[\s]*[:：]\s*([^\n]{1,200})', full)
        if m:
            v = m.group(1).strip().rstrip('，。.;,;）\)')
            if v: val = v
        else:
            m = re.search(re.escape(field) + r'[\t ]{1,4}([^\s]{1,100})', full)
            if m:
                v = m.group(1).strip().rstrip('，。.;,;）\)')
                if v: val = v
        result[field] = val
    return result

def create_excel(path, fields):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "标本数据"
    hf = Font(bold=True, size=11, color="FFFFFF")
    hfill = PatternFill("solid", fgColor="4472C4")
    halign = Alignment(horizontal="center", vertical="center", wrap_text=True)
    bdr = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
    headers = ["序号"] + fields
    widths = [8] + [16] * len(fields)
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font, c.fill, c.alignment, c.border = hf, hfill, halign, bdr
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'
    wb.save(path)
    return path

def write_row(excel_path, row_data, field_order, row_num):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    bdr = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
    c = ws.cell(row=row_num, column=1, value=row_num - 1); c.border = bdr
    for i, field in enumerate(field_order, 2):
        v = row_data.get(field, "") or ""
        c = ws.cell(row=row_num, column=i, value=v); c.border = bdr
    wb.save(excel_path)

def count_data_rows(excel_path):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    return max(0, ws.max_row - 1)

def get_excel_headers(excel_path):
    """读取 Excel 表头字段（跳过「序号」列）"""
    if not os.path.isfile(excel_path):
        return None
    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        headers = []
        for c in range(2, ws.max_column + 1):
            v = ws.cell(row=1, column=c).value
            if v and str(v).strip() and str(v).strip() != "序号":
                headers.append(str(v).strip())
        return headers if headers else None
    except Exception:
        return None


# ─── 主应用 ───
class App:
    def __init__(self):
        if tk is None:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, "需要 tkinter（Windows 自带 Python 已包含）", "错误", 0)
            sys.exit(1)

        self.fields = DEFAULT_FIELDS.copy()
        self.image_folder = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.running = False
        self.monitoring = False
        self.monitor_thread = None
        self._stop_monitor = threading.Event()

        self._load_config()
        self._build_ui()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()

    def _build_ui(self):
        self.root = tk.Tk()
        self.root.title("标本照片 → Excel 自动填表")
        self.root.geometry("780x680")
        self.root.minsize(640, 520)

        m = ttk.Frame(self.root, padding=12)
        m.pack(fill=tk.BOTH, expand=True)

        ttk.Label(m, text="标本照片 OCR → Excel 自动填表",
                  font=('微软雅黑', 14, 'bold')).pack(pady=(0, 4))

        # ── 字段管理 ──
        ft = ttk.Frame(m)
        ft.pack(fill=tk.X, pady=(0, 4))
        self.fields_btn = ttk.Button(
            ft, text=f"⚙️ 管理字段（当前 {len(self.fields)} 个）",
            command=self._manage_fields, width=28
        )
        self.fields_btn.pack(side=tk.LEFT)
        ttk.Label(ft, text="增删改字段，自定义识别内容；也支持从 Excel 表头同步",
                  foreground="gray").pack(side=tk.LEFT, padx=8)

        # ── 图片文件夹 ──
        f1 = ttk.LabelFrame(m, text="📁 图片文件夹（按文件名排序处理）", padding=8)
        f1.pack(fill=tk.X, pady=3)
        r1 = ttk.Frame(f1)
        r1.pack(fill=tk.X)
        ttk.Entry(r1, textvariable=self.image_folder, width=55).pack(
            side=tk.LEFT, padx=(0, 6))
        ttk.Button(r1, text="浏览...", command=self._pick_folder,
                   width=8).pack(side=tk.LEFT)
        self.fs = ttk.Label(f1, foreground="gray")
        self.fs.pack(anchor=tk.W, pady=(3, 0))

        # ── Excel 输出 ──
        f2 = ttk.LabelFrame(m, text="📊 Excel 输出文件", padding=8)
        f2.pack(fill=tk.X, pady=3)
        r2 = ttk.Frame(f2)
        r2.pack(fill=tk.X)
        ttk.Entry(r2, textvariable=self.excel_path, width=55).pack(
            side=tk.LEFT, padx=(0, 6))
        ttk.Button(r2, text="选择已有...", command=self._pick_excel,
                   width=10).pack(side=tk.LEFT)
        ttk.Button(r2, text="新建...", command=self._new_excel,
                   width=8).pack(side=tk.LEFT, padx=4)
        self.es = ttk.Label(f2, foreground="gray")
        self.es.pack(anchor=tk.W, pady=(3, 0))

        self.image_folder.trace_add("write", lambda *a: self._refresh())
        self.excel_path.trace_add("write", lambda *a: self._refresh())
        self._refresh()

        # ── 日志 ──
        f3 = ttk.LabelFrame(m, text="📝 运行日志", padding=6)
        f3.pack(fill=tk.BOTH, expand=True, pady=3)
        self.log = scrolledtext.ScrolledText(
            f3, height=12, font=('Consolas', 10),
            bg="#1e1e1e", fg="#d4d4d4", insertbackground="white",
            state=tk.DISABLED)
        self.log.pack(fill=tk.BOTH, expand=True)
        for tag, color in [("i", "#d4d4d4"), ("ok", "#4ec9b0"),
                           ("err", "#f44747"), ("w", "#ce9178"),
                           ("b", "#569cd6")]:
            self.log.tag_config(tag, foreground=color)
        self.log.tag_config("b", font=('Consolas', 11, 'bold'))

        self.pb = ttk.Progressbar(m, mode='determinate')
        self.pb.pack(fill=tk.X, pady=2)
        self.pl = ttk.Label(m, text="", foreground="gray")
        self.pl.pack()

        # ── 按钮区 ──
        bf = ttk.Frame(m)
        bf.pack(fill=tk.X, pady=(4, 0))

        self.run_btn = ttk.Button(
            bf, text="🚀 手动识别并填表", width=20, command=self._start)
        self.run_btn.pack(side=tk.RIGHT, padx=3)

        self.monitor_btn = ttk.Button(
            bf, text="👁 开始监听", width=14, command=self._toggle_monitor)
        self.monitor_btn.pack(side=tk.RIGHT, padx=3)

        ttk.Button(bf, text="清空日志", width=10,
                   command=lambda: (self.log.config(state=tk.NORMAL),
                                    self.log.delete('1.0', tk.END),
                                    self.log.config(state=tk.DISABLED))
                   ).pack(side=tk.RIGHT, padx=3)

        ttk.Button(bf, text="打开输出目录", width=12,
                   command=self._open_dir).pack(side=tk.LEFT, padx=3)

        self.tray_btn = ttk.Button(
            bf, text="— 最小化到托盘", width=14,
            command=self._minimize_to_tray)
        self.tray_btn.pack(side=tk.LEFT, padx=3)

    # ─── 日志 / 刷新 ───
    def _wl(self, msg, tag="i"):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, msg + "\n", tag)
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def _refresh(self):
        d, e = self.image_folder.get(), self.excel_path.get()
        n = len([p for p in Path(d).iterdir() if is_img(p)]) \
            if d and os.path.isdir(d) else 0
        self.fs.config(
            text=f"已选（{n} 张图片）" if n else "（未选择）",
            foreground="green" if n else "gray")
        self.es.config(
            text="✅ 文件存在" if e and os.path.isfile(e)
            else ("⚠️ 不存在，自动创建" if e else "（未选择）"),
            foreground="green" if e and os.path.isfile(e)
            else ("orange" if e else "gray"))

    # ─── 路径选择 ───
    def _pick_folder(self):
        d = filedialog.askdirectory()
        d and self.image_folder.set(d)

    def _pick_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        p and self.excel_path.set(p)

    def _new_excel(self):
        p = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if p:
            create_excel(p, self.fields)
            self.excel_path.set(p)
            self._wl(f"📋 已创建: {Path(p).name}", "ok")

    def _open_dir(self):
        p = self.excel_path.get()
        p and os.path.isfile(p) and os.startfile(os.path.dirname(p))

    # ─── 字段管理 ───
    def _manage_fields(self):
        win = tk.Toplevel(self.root)
        win.title("字段管理")
        win.geometry("440x520")
        win.transient(self.root)
        win.grab_set()

        ttk.Label(win, text="管理识别字段",
                  font=('微软雅黑', 12, 'bold')).pack(pady=(10, 2))
        ttk.Label(
            win,
            text="每行一个字段名，需与照片上的文字名称完全一致",
            foreground="gray").pack()
        ttk.Label(
            win,
            text='例如照片上写「采集号：A001」，字段名就写「采集号」',
            foreground="gray").pack(pady=(0, 6))

        text = scrolledtext.ScrolledText(win, height=14, font=('微软雅黑', 10))
        text.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)
        for f in self.fields:
            text.insert(tk.END, f + "\n")

        def sync_from_excel():
            e = self.excel_path.get()
            if not e or not os.path.isfile(e):
                messagebox.showinfo("提示", "请先选择 Excel 文件", parent=win)
                return
            headers = get_excel_headers(e)
            if not headers:
                messagebox.showinfo("提示", "Excel 中未找到有效表头", parent=win)
                return
            text.delete('1.0', tk.END)
            for h in headers:
                text.insert(tk.END, h + "\n")
            self._wl(f"📋 已从 Excel 同步 {len(headers)} 个字段", "ok")

        def save_fields():
            raw = text.get('1.0', tk.END).strip()
            new_fields = [l.strip() for l in raw.split('\n') if l.strip()]
            if not new_fields:
                messagebox.showwarning("提示", "至少保留一个字段", parent=win)
                return
            self.fields = new_fields
            self._save_config()
            self.fields_btn.config(
                text=f"⚙️ 管理字段（当前 {len(self.fields)} 个）")
            win.destroy()
            n = len(self.fields)
            disp = ', '.join(self.fields[:5])
            self._wl(
                f"✅ 字段已更新（{n} 个）：{disp}{'...' if n > 5 else ''}",
                "ok")

        bf = ttk.Frame(win)
        bf.pack(fill=tk.X, padx=12, pady=6)

        ttk.Button(bf, text="从 Excel 表头同步",
                   command=sync_from_excel).pack(side=tk.LEFT)
        ttk.Button(bf, text="重置默认",
                   command=lambda: (text.delete('1.0', tk.END),
                                    text.insert(
                                        tk.END,
                                        "\n".join(DEFAULT_FIELDS) + "\n"))
                   ).pack(side=tk.LEFT, padx=6)

        ttk.Button(bf, text="✅ 保存", width=10,
                   command=save_fields).pack(side=tk.RIGHT, padx=4)
        ttk.Button(bf, text="取消", width=8,
                   command=win.destroy).pack(side=tk.RIGHT, padx=4)

    # ─── 系统托盘 ───
    def _minimize_to_tray(self):
        try:
            from PIL import Image, ImageDraw, ImageFont
            import pystray

            self.root.withdraw()
            self._wl("🔽 已最小化到系统托盘，双击图标恢复窗口", "ok")

            img = Image.new('RGBA', (64, 64), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            draw.ellipse([2, 2, 62, 62], fill=(68, 114, 196))
            try:
                font = ImageFont.truetype("msyh.ttc", 28)
            except Exception:
                font = ImageFont.load_default()
            draw.text((16, 16), "标", fill=(255, 255, 255), font=font)

            def on_show(icon, item):
                icon.stop()
                self.root.after(0, self.root.deiconify)

            def on_run(icon, item):
                icon.stop()
                self.root.after(
                    0, lambda: (self.root.deiconify(), self._start()))

            def on_quit(icon, item):
                icon.stop()
                self._save_config()
                if self.monitoring:
                    self._stop_monitor.set()
                self.root.after(0, self._quit)

            menu = pystray.Menu(
                pystray.MenuItem("显示窗口", on_show, default=True),
                pystray.MenuItem("立即识别", on_run),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("退出", on_quit),
            )
            icon = pystray.Icon("specimen_ocr", img, "标本OCR填表工具", menu)
            threading.Thread(target=icon.run, daemon=True).start()

        except ImportError:
            self.root.iconify()
            self._wl("📌 已最小化到任务栏"
                      "（安装 pystray 可支持系统托盘）", "i")
        except Exception as e:
            self.root.iconify()
            self._wl(f"📌 已最小化到任务栏（托盘异常: {e}）", "w")

    # ─── 文件夹监听 ───
    def _toggle_monitor(self):
        if self.monitoring:
            self._stop_monitor.set()
            self.monitoring = False
            self.monitor_btn.config(text="👁 开始监听")
            self._wl("🛑 已停止文件夹监听", "w")
            return

        d = self.image_folder.get()
        if not d or not os.path.isdir(d):
            messagebox.showwarning("提示", "请先选择图片文件夹")
            return

        self._stop_monitor.clear()
        self.monitoring = True
        self.monitor_btn.config(text="⏹ 停止监听")
        self._wl("👁 开始监听文件夹", "ok")
        self._wl(f"   📂 {d}", "i")
        self._wl("   → 检测到新图片将自动识别并填入 Excel", "i")

        def monitor_loop():
            processed = set()
            try:
                for p in Path(d).iterdir():
                    if is_img(p):
                        processed.add(p.name)
            except Exception:
                pass
            self.root.after(
                0, lambda: self._wl(
                    f"   ℹ️ 已有 {len(processed)} 张图片，仅处理新增的", "i"))

            while not self._stop_monitor.is_set():
                time.sleep(3)
                try:
                    e = self.excel_path.get()
                    if not e or not os.path.isdir(d):
                        continue
                    current = sorted(
                        [p for p in Path(d).iterdir() if is_img(p)])
                    new_imgs = [p for p in current
                                if p.name not in processed]
                    if not new_imgs:
                        continue

                    fields = self.fields
                    hdrs = get_excel_headers(e)
                    if hdrs:
                        fields = hdrs

                    for img in new_imgs:
                        if self._stop_monitor.is_set():
                            break
                        name = img.name
                        self.root.after(
                            0, lambda n=name: self._wl(
                                f"[监听] 📷 {n}", "ok"))
                        try:
                            texts = ocr_image(str(img))
                            if texts:
                                row_data = parse_fields(texts, fields)
                                base = count_data_rows(e)
                                if not os.path.isfile(e):
                                    create_excel(e, fields)
                                write_row(e, row_data, fields, base + 2)
                                processed.add(name)
                                self.root.after(
                                    0, lambda n=name: self._wl(
                                        f"[监听] ✅ {n} → 已填入 Excel",
                                        "ok"))
                            else:
                                processed.add(name)
                                self.root.after(
                                    0, lambda n=name: self._wl(
                                        f"[监听] ⚠️ {n} 未识别到文字",
                                        "w"))
                        except Exception as ex:
                            self.root.after(
                                0, lambda n=name, e=str(ex): self._wl(
                                    f"[监听] ❌ {n}: {e}", "err"))
                except Exception:
                    pass

        self.monitor_thread = threading.Thread(
            target=monitor_loop, daemon=True)
        self.monitor_thread.start()

    # ─── 手动识别 ───
    def _start(self):
        if self.running:
            return
        d, e = self.image_folder.get(), self.excel_path.get()
        if not d or not os.path.isdir(d):
            messagebox.showwarning("提示", "请选择图片文件夹")
            return
        if not e:
            messagebox.showwarning("提示", "请选择或新建 Excel")
            return

        self.running = True
        self.run_btn.config(state=tk.DISABLED, text="⏳ 处理中...")
        self.pb['value'] = 0
        self.pl.config(text="正在初始化 OCR 引擎...")
        self._wl("=" * 45, "b")
        self._wl("🚀 开始处理标本照片...", "b")
        self._wl("=" * 45, "b")

        def worker():
            fields = self.fields
            hdrs = get_excel_headers(e)
            if hdrs:
                fields = hdrs
                if hdrs != self.fields:
                    self.root.after(
                        0, lambda: self._wl(
                            f"   ℹ️ 字段已与 Excel 表头同步"
                            f"（{len(hdrs)} 个）", "i"))

            def pc(c, t):
                self.root.after(0, lambda: (
                    self.pb.configure(value=c, maximum=t),
                    self.pl.config(
                        text=f"处理中 {c}/{t} [{c*100//t}%]")
                ))

            def lc(m, tag="i"):
                self.root.after(0, lambda: self._wl(m, tag))

            try:
                images = sorted(
                    [p for p in Path(d).iterdir() if is_img(p)])
                total = len(images)
                if total == 0:
                    lc("⚠️ 未找到图片文件", "w")
                    return
                if not os.path.isfile(e):
                    create_excel(e, fields)
                    lc(f"📋 已创建 Excel 模板: {Path(e).name}", "ok")

                base_rows = count_data_rows(e)
                succeed = 0

                for idx, img in enumerate(images):
                    name = img.name
                    lc(f"[{idx+1}/{total}] 📷 {name}")
                    pc(idx + 1, total)
                    try:
                        texts = ocr_image(str(img))
                        if not texts:
                            lc(f"   ⚠️ 未识别到文字", "w")
                            continue
                        row_data = parse_fields(texts, fields)
                        for f in fields:
                            v = row_data.get(f, "") or "（未识别）"
                            lc(f"   · {f} → {v}")
                        row_num = base_rows + succeed + 2
                        write_row(e, row_data, fields, row_num)
                        succeed += 1
                        lc(f"   ✅ 已写入 → 行{row_num}", "ok")
                    except Exception as ex:
                        lc(f"   ❌ 处理失败: {ex}", "err")

                lc(f"\n{'=' * 45}", "b")
                lc(f"✅ 完成！共 {total} 张图片，"
                   f"成功写入 {succeed} 行", "ok")
            except Exception as ex:
                lc(f"❌ 运行出错: {ex}", "err")
            finally:
                self.running = False
                self.root.after(0, self._done)

        threading.Thread(target=worker, daemon=True).start()

    def _done(self):
        self.run_btn.config(state=tk.NORMAL, text="🚀 手动识别并填表")
        self.pl.config(text="✅ 完成")
        self._save_config()

    def _quit(self):
        self._save_config()
        if self.monitoring:
            self._stop_monitor.set()
        self.root.quit()
        self.root.destroy()

    def _on_close(self):
        """关闭窗口 → 最小化到托盘（或任务栏）"""
        self._save_config()
        self._minimize_to_tray()

    # ─── 配置持久化 ───
    def _save_config(self):
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            Path(CONFIG_FILE).write_text(
                json.dumps({
                    "image_folder": self.image_folder.get(),
                    "excel_path": self.excel_path.get(),
                    "fields": self.fields
                }, ensure_ascii=False, indent=2),
                encoding='utf-8')
        except Exception:
            pass

    def _load_config(self):
        try:
            if CONFIG_FILE.exists():
                d = json.loads(
                    CONFIG_FILE.read_text(encoding='utf-8'))
                if isinstance(d.get("fields"), list) \
                        and len(d["fields"]) > 0:
                    self.fields = d["fields"]
                p1 = d.get("image_folder", "")
                if p1 and os.path.isdir(p1):
                    self.image_folder.set(p1)
                p2 = d.get("excel_path", "")
                if p2:
                    self.excel_path.set(p2)
        except Exception:
            pass


if __name__ == "__main__":
    try:
        from rapidocr_onnxruntime import RapidOCR
    except ImportError:
        print("=" * 50)
        print("📦 请先安装依赖：")
        print("   pip install rapidocr-onnxruntime openpyxl Pillow pystray")
        print("=" * 50)
        sys.exit(1)
    App()

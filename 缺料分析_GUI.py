# -*- coding: utf-8 -*-
"""
達特世生技 ─ 缺料分析系統 (GUI)
版本: v1.0  2026-05-13
"""
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import subprocess, sys, os, glob, threading, re, shutil, tempfile
from datetime import datetime

# ─── 路徑設定 ──────────────────────────────────────────────
BASE   = r"C:\Users\admin\OneDrive\桌面\CODE資料"
DESK   = r"C:\Users\admin\OneDrive\桌面"
SCRIPT = os.path.join(BASE, "缺料分析.py")
PYTHON = sys.executable

# ─── 顏色主題 ──────────────────────────────────────────────
C = {
    'bg'       : '#F4F6F9',
    'header_bg': '#1F4E79',
    'header_fg': '#FFFFFF',
    'sub_fg'   : '#AED6F1',
    'card_bg'  : '#FFFFFF',
    'btn_run'  : '#1F6FBF',
    'btn_run_h': '#1557A0',
    'btn_open' : '#1E8449',
    'btn_dir'  : '#616A6B',
    'ok_fg'    : '#1E8449',
    'err_fg'   : '#922B21',
    'warn_fg'  : '#D35400',
    'miss_fg'  : '#999999',
    'log_bg'   : '#1E2227',
    'log_ok'   : '#98C379',
    'log_err'  : '#E06C75',
    'log_warn' : '#E5C07B',
    'log_info' : '#61AFEF',
    'log_done' : '#C678DD',
    'log_def'  : '#ABB2BF',
    'stat_bg'  : '#EAF4FB',
    'stat_val' : '#1A5276',
}

FONT_MAIN  = ('Microsoft JhengHei UI', 10)
FONT_BOLD  = ('Microsoft JhengHei UI', 10, 'bold')
FONT_TITLE = ('Microsoft JhengHei UI', 15, 'bold')
FONT_MONO  = ('Consolas', 9)
FONT_STAT  = ('Microsoft JhengHei UI', 20, 'bold')
FONT_STAT_S= ('Microsoft JhengHei UI', 9)

# ─── 檔案偵測 ──────────────────────────────────────────────
def latest2(*pats):
    files = []
    for p in pats:
        files.extend(glob.glob(p))
    return max(files, key=os.path.getmtime) if files else None

FILE_DEFS = [
    # (key, label,  必/選, glob patterns)
    ('inv',   '原物料庫存', '必要',
     [os.path.join(DESK, '原物料庫存*.xlsx'),
      os.path.join(BASE, '原物料庫存*.xlsx')]),
    ('order', '訂單生產總表', '必要',
     [os.path.join(DESK, '訂單生產總表*.xlsx'),
      os.path.join(BASE, '訂單生產總表*.xlsx')]),
    ('ear',   '耳繩庫存',  '建議',
     [os.path.join(DESK, '耳繩庫存*.xls'),
      os.path.join(DESK, '耳繩庫存*.xlsx'),
      os.path.join(BASE, '耳繩庫存*.xls'),
      os.path.join(BASE, '耳繩庫存*.xlsx')]),
    ('box',   '盒子庫存',  '建議',
     [os.path.join(DESK, '盒子庫存*.xls'),
      os.path.join(BASE, '盒子庫存*.xls')]),
]


def detect_all():
    """偵測所有輸入檔案，回傳 {key: path_or_None}"""
    return {k: latest2(*pats) for k, _, _, pats in FILE_DEFS}


# ══════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("達特世生技｜缺料分析系統")
        self.geometry("820x700")
        self.minsize(700, 560)
        self.configure(bg=C['bg'])
        try:
            self.iconbitmap(default='')
        except Exception:
            pass

        self.running     = False
        self.result_path = None
        self._override   = {}   # key → full path (使用者手動選擇)

        self._build_ui()
        self._refresh_files()

    # ── UI 建構 ───────────────────────────────────────────
    def _build_ui(self):
        self._build_header()
        self._build_files()
        self._build_run_btn()
        self._build_log()
        self._build_stats()
        self._build_footer()

    def _build_header(self):
        h = tk.Frame(self, bg=C['header_bg'])
        h.pack(fill='x')
        tk.Label(h, text="🏭  達特世生技 ‧ 缺料分析系統",
                 font=FONT_TITLE, bg=C['header_bg'], fg=C['header_fg'],
                 pady=10).pack()
        tk.Label(h, text="自動比對庫存與訂單需求 ‧ 輸出缺料清單 Excel",
                 font=('Microsoft JhengHei UI', 9), bg=C['header_bg'],
                 fg=C['sub_fg'], pady=0).pack(pady=(0, 8))

    def _build_files(self):
        outer = tk.Frame(self, bg=C['bg'], padx=16, pady=8)
        outer.pack(fill='x')

        card = tk.Frame(outer, bg=C['card_bg'],
                        relief='flat', bd=0,
                        highlightbackground='#D5DBDB',
                        highlightthickness=1)
        card.pack(fill='x')

        tk.Label(card, text=" 📁  輸入檔案",
                 font=FONT_BOLD, bg=C['card_bg'], fg='#1A5276',
                 anchor='w', pady=6, padx=10).pack(fill='x')
        ttk.Separator(card, orient='horizontal').pack(fill='x', padx=10)

        self._file_name_vars  = {}   # key → StringVar  (顯示檔名)
        self._file_name_lbl   = {}   # key → Label
        self._file_status_lbl = {}   # key → Label (狀態圖示)

        grid = tk.Frame(card, bg=C['card_bg'], padx=12, pady=8)
        grid.pack(fill='x')
        grid.columnconfigure(1, weight=1)

        for i, (key, label, req, _) in enumerate(FILE_DEFS):
            req_color = C['err_fg'] if req == '必要' else C['warn_fg']
            # 標籤
            tk.Label(grid, text=f"[{req}]", fg=req_color,
                     font=('Microsoft JhengHei UI', 8, 'bold'),
                     bg=C['card_bg'], width=5, anchor='e'
                     ).grid(row=i, column=0, padx=(0, 4), pady=3, sticky='e')
            tk.Label(grid, text=f"{label}",
                     font=FONT_MAIN, bg=C['card_bg'], width=10, anchor='e'
                     ).grid(row=i, column=1, padx=(0, 6), pady=3, sticky='e')

            # 路徑顯示
            var = tk.StringVar()
            self._file_name_vars[key] = var
            nm_lbl = tk.Label(grid, textvariable=var,
                              font=FONT_MONO, bg=C['card_bg'],
                              fg=C['ok_fg'], anchor='w', width=46)
            nm_lbl.grid(row=i, column=2, sticky='ew', pady=3)
            self._file_name_lbl[key] = nm_lbl

            # 更換按鈕
            btn = tk.Button(grid, text='更換',
                            command=lambda k=key: self._pick_file(k),
                            font=('Microsoft JhengHei UI', 8),
                            bg='#ECF0F1', fg='#333',
                            relief='flat', padx=8, cursor='hand2')
            btn.grid(row=i, column=3, padx=(6, 0), pady=3)

        # 重新偵測
        refresh_row = tk.Frame(card, bg=C['card_bg'])
        refresh_row.pack(fill='x', padx=12, pady=(0, 10))
        tk.Label(refresh_row,
                 text="⚡ 系統自動抓取桌面最新版本",
                 font=('Microsoft JhengHei UI', 8), bg=C['card_bg'],
                 fg='#888').pack(side='left')
        tk.Button(refresh_row, text='🔄 重新偵測',
                  command=self._refresh_files,
                  font=('Microsoft JhengHei UI', 8),
                  bg='#ECF0F1', fg='#555', relief='flat',
                  padx=8, cursor='hand2').pack(side='right')

    def _build_run_btn(self):
        f = tk.Frame(self, bg=C['bg'], pady=6)
        f.pack(fill='x', padx=16)
        self.run_btn = tk.Button(
            f, text='▶   執行缺料分析',
            command=self._run,
            font=('Microsoft JhengHei UI', 13, 'bold'),
            bg=C['btn_run'], fg='white', activebackground=C['btn_run_h'],
            activeforeground='white', relief='flat',
            padx=40, pady=10, cursor='hand2'
        )
        self.run_btn.pack(expand=True)
        self.prog = ttk.Progressbar(f, mode='indeterminate', length=600)
        self.prog.pack(fill='x', pady=(6, 0))

    def _build_log(self):
        f = tk.Frame(self, bg=C['bg'], padx=16)
        f.pack(fill='both', expand=True)

        hdr = tk.Frame(f, bg=C['bg'])
        hdr.pack(fill='x')
        tk.Label(hdr, text="📋  執行紀錄",
                 font=FONT_BOLD, bg=C['bg'], fg='#1A5276').pack(side='left')
        self.status_lbl = tk.Label(hdr, text="",
                                    font=FONT_MAIN, bg=C['bg'], fg='#555')
        self.status_lbl.pack(side='right')

        self.log = scrolledtext.ScrolledText(
            f, font=FONT_MONO, bg=C['log_bg'], fg=C['log_def'],
            insertbackground='white', relief='flat',
            height=12, state='disabled'
        )
        self.log.pack(fill='both', expand=True, pady=(4, 0))

        # 顏色 tag
        for tag, fg in [('ok',   C['log_ok']),
                        ('err',  C['log_err']),
                        ('warn', C['log_warn']),
                        ('info', C['log_info']),
                        ('done', C['log_done'])]:
            kw = {'foreground': fg}
            if tag == 'done':
                kw['font'] = ('Consolas', 9, 'bold')
            self.log.tag_config(tag, **kw)

    def _build_stats(self):
        f = tk.Frame(self, bg=C['stat_bg'],
                     highlightbackground='#AED6F1',
                     highlightthickness=1,
                     padx=16, pady=8)
        f.pack(fill='x', padx=16, pady=(6, 0))

        self._stat_vars = {}
        for col, (key, label, color) in enumerate([
            ('orders',    '訂單數',  '#1A5276'),
            ('materials', '材料項數', '#1A5276'),
            ('shortage',  '缺料項數', '#922B21'),
        ]):
            cell = tk.Frame(f, bg=C['stat_bg'])
            cell.pack(side='left', expand=True, fill='x')
            var = tk.StringVar(value='—')
            self._stat_vars[key] = var
            tk.Label(cell, textvariable=var,
                     font=FONT_STAT, bg=C['stat_bg'], fg=color).pack()
            tk.Label(cell, text=label,
                     font=FONT_STAT_S, bg=C['stat_bg'], fg='#555').pack()
            if col < 2:
                ttk.Separator(f, orient='vertical').pack(
                    side='left', fill='y', padx=8)

    def _build_footer(self):
        f = tk.Frame(self, bg=C['bg'], padx=16, pady=8)
        f.pack(fill='x')

        self.open_btn = tk.Button(
            f, text='📂  開啟結果 Excel',
            command=self._open_result,
            font=FONT_BOLD, bg=C['btn_open'], fg='white',
            activebackground='#1A6B3C', activeforeground='white',
            relief='flat', padx=16, pady=7,
            state='disabled', cursor='hand2'
        )
        self.open_btn.pack(side='left', padx=(0, 8))

        self.dir_btn = tk.Button(
            f, text='📁  開啟資料夾',
            command=self._open_dir,
            font=FONT_MAIN, bg=C['btn_dir'], fg='white',
            activebackground='#4D5656', activeforeground='white',
            relief='flat', padx=12, pady=7,
            cursor='hand2'
        )
        self.dir_btn.pack(side='left')

        self.time_lbl = tk.Label(f, text="",
                                  font=('Microsoft JhengHei UI', 8),
                                  bg=C['bg'], fg='#888')
        self.time_lbl.pack(side='right')

    # ── 檔案偵測 ──────────────────────────────────────────
    def _refresh_files(self):
        found = detect_all()
        for key, label, req, _ in FILE_DEFS:
            # 若使用者已手動覆寫，優先使用
            path = self._override.get(key) or found.get(key)
            if path:
                self._file_name_vars[key].set(os.path.basename(path))
                self._file_name_lbl[key].config(fg=C['ok_fg'])
            else:
                self._file_name_vars[key].set('❌  未找到')
                self._file_name_lbl[key].config(fg=C['miss_fg'])

    def _pick_file(self, key):
        ft = [('Excel/xls', '*.xlsx *.xls'), ('All', '*.*')]
        path = filedialog.askopenfilename(
            title=f'選擇{dict((k, l) for k,l,*_ in FILE_DEFS).get(key,key)}檔案',
            initialdir=DESK, filetypes=ft)
        if not path:
            return
        self._override[key] = path
        self._file_name_vars[key].set(os.path.basename(path))
        self._file_name_lbl[key].config(fg='#2E86C1')  # 藍色表示手動選擇

    # ── Log 工具 ──────────────────────────────────────────
    def _log_write(self, text, tag=''):
        self.log.configure(state='normal')
        self.log.insert('end', text + '\n', tag)
        self.log.see('end')
        self.log.configure(state='disabled')

    def _log_clear(self):
        self.log.configure(state='normal')
        self.log.delete('1.0', 'end')
        self.log.configure(state='disabled')

    def _classify(self, line):
        if re.search(r'[✅完成]', line):                return 'done'
        if '✓' in line:                                 return 'ok'
        if '❌' in line:                                 return 'err'
        if '⚠' in line or 'warn' in line.lower():      return 'warn'
        if re.search(r'===|找到檔案|^\s*\[', line):    return 'info'
        return ''

    # ── 執行分析 ──────────────────────────────────────────
    def _run(self):
        if self.running:
            return

        # 確認必要檔案
        found = detect_all()
        for key, label, req, _ in FILE_DEFS:
            if req == '必要':
                path = self._override.get(key) or found.get(key)
                if not path:
                    messagebox.showerror('缺少必要檔案',
                                         f'找不到「{label}」\n請確認檔案已放在桌面或 CODE資料 資料夾')
                    return

        self.running = True
        self.run_btn.config(state='disabled', text='⏳  分析中，請稍候…')
        self.open_btn.config(state='disabled')
        self.prog.start(12)
        self.result_path = None
        self._stat_vars['orders'].set('—')
        self._stat_vars['materials'].set('—')
        self._stat_vars['shortage'].set('—')

        self._log_clear()
        self._log_write(
            f"▶  開始執行缺料分析  [{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]",
            'info')
        self._log_write('─' * 64)

        def worker():
            # 如果使用者有手動選擇檔案，先複製到桌面（使其成為最新）
            _copied = []
            try:
                for key, *_ in FILE_DEFS:
                    if key in self._override:
                        src = self._override[key]
                        ext = os.path.splitext(src)[1]
                        prefixes = {'inv': '原物料庫存',
                                    'order': '訂單生產總表',
                                    'ear': '耳繩庫存',
                                    'box': '盒子庫存'}
                        dst = os.path.join(
                            DESK,
                            f"_gui_override_{prefixes[key]}{ext}")
                        shutil.copy2(src, dst)
                        # 更新修改時間確保最新
                        os.utime(dst, None)
                        _copied.append(dst)
            except Exception as ce:
                self.after(0, self._log_write,
                           f'⚠  複製覆寫檔案失敗: {ce}', 'warn')

            try:
                proc = subprocess.Popen(
                    [PYTHON, '-u', SCRIPT],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    cwd=BASE,
                    encoding='utf-8',
                    errors='replace',
                    bufsize=1
                )

                result_path = None
                orders = materials = shortage = None

                for raw in proc.stdout:
                    line = raw.rstrip('\r\n')
                    tag  = self._classify(line)
                    self.after(0, self._log_write, line, tag)

                    # 擷取輸出路徑
                    m = re.search(r'輸出:\s*(.+\.xlsx)', line)
                    if m:
                        result_path = m.group(1).strip()

                    # 擷取統計數字
                    m2 = re.search(
                        r'訂單:\s*(\d+)\s*筆.*材料:\s*(\d+)\s*項.*缺料:\s*(\d+)\s*項',
                        line)
                    if m2:
                        orders, materials, shortage = m2.groups()

                proc.wait()

                # 清除暫存覆寫檔
                for f in _copied:
                    try:
                        os.remove(f)
                    except Exception:
                        pass

                self.after(0, self._finish,
                           proc.returncode, result_path,
                           orders, materials, shortage)

            except Exception as e:
                self.after(0, self._log_write,
                           f'❌  執行錯誤: {e}', 'err')
                self.after(0, self._finish, -1, None, None, None, None)

        threading.Thread(target=worker, daemon=True).start()

    def _finish(self, rc, result_path, orders, materials, shortage):
        self.running = False
        self.prog.stop()
        self.run_btn.config(state='normal', text='▶   執行缺料分析')

        now = datetime.now().strftime('%H:%M:%S')
        self._log_write('─' * 64)

        if rc == 0 and result_path:
            self.result_path = result_path
            self.open_btn.config(state='normal')
            self._log_write(f'✅  分析完成！  [{now}]', 'done')
            self._log_write(f'   結果檔案：{result_path}', 'done')
            self.status_lbl.config(
                text=f'✅ 完成 {now}', fg=C['ok_fg'])
            self.time_lbl.config(text=f'完成時間：{now}')

            if orders:
                self._stat_vars['orders'].set(orders)
            if materials:
                self._stat_vars['materials'].set(materials)
            if shortage:
                self._stat_vars['shortage'].set(shortage)
                # 缺料項數顯示紅色提醒
                n = int(shortage)
                color = '#922B21' if n > 0 else C['ok_fg']
                self._stat_vars['shortage'].set(shortage)
                for w in self.winfo_children():
                    self._recolor_stat(w, 'shortage', color)
        else:
            self._log_write(
                f'⚠  執行結束（代碼: {rc}）請檢查上方 log', 'warn')
            self.status_lbl.config(
                text=f'⚠ 請檢查 log  {now}', fg=C['warn_fg'])

    def _recolor_stat(self, widget, key, color):
        """遞迴找 _stat_vars[key] 對應的 Label 並改色（簡易實作）"""
        try:
            if isinstance(widget, tk.Label):
                if str(widget.cget('textvariable')) == str(self._stat_vars[key]):
                    widget.config(fg=color)
        except Exception:
            pass
        for child in widget.winfo_children():
            self._recolor_stat(child, key, color)

    # ── 開啟結果 ──────────────────────────────────────────
    def _open_result(self):
        if self.result_path and os.path.exists(self.result_path):
            os.startfile(self.result_path)
        else:
            messagebox.showwarning('找不到檔案',
                                   f'找不到結果檔案：\n{self.result_path}')

    def _open_dir(self):
        today  = datetime.now().strftime('%m%d')
        outdir = os.path.join(BASE, f'缺料分析{today}')
        target = outdir if os.path.exists(outdir) else BASE
        os.startfile(target)


# ══════════════════════════════════════════════════════════
if __name__ == '__main__':
    app = App()
    app.mainloop()

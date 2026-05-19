# -*- coding: utf-8 -*-
"""
達特世生技 ─ 缺料分析系統 (GUI)
版本: v1.6  2026-05-19
更新: 損耗率調整 — 印刷布6%、一般布料4%
"""
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import subprocess, sys, os, glob, threading, re, shutil, tempfile, math
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
FONT_SM    = ('Microsoft JhengHei UI', 9)

# ─── 檔案偵測 ──────────────────────────────────────────────
def latest2(*pats):
    files = []
    for p in pats:
        files.extend(glob.glob(p))
    return max(files, key=os.path.getmtime) if files else None

FILE_DEFS = [
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
    return {k: latest2(*pats) for k, _, _, pats in FILE_DEFS}


# ══════════════════════════════════════════════════════════
# DRX-3D 品項設定
# ══════════════════════════════════════════════════════════
DRX_DESIGNS = [
    '小車與飛行隊',
    '卡皮巴拉',
    '粉嫩獨角獸',
    '企鵝寶寶',
    '花花兔',
    '恐龍樂園',
    '淘氣狗',
    '愛心兔',
    '英雄救援車',
]
DRX_CHANNELS = ['啄木鳥', '家樂福', '躍獅', '其他']

# 上次訂單預設值（可清空重填）
DRX_DEFAULT = {
    '小車與飛行隊': {'啄木鳥': 360, '家樂福': 1464, '躍獅': 0,   '其他': 0, '庫存捲': 0},
    '卡皮巴拉':     {'啄木鳥': 0,   '家樂福': 1464, '躍獅': 120, '其他': 0, '庫存捲': 5},
    '粉嫩獨角獸':   {'啄木鳥': 0,   '家樂福': 1464, '躍獅': 0,   '其他': 0, '庫存捲': 0},
    '企鵝寶寶':     {'啄木鳥': 120, '家樂福': 0,    '躍獅': 0,   '其他': 0, '庫存捲': 0},
    '花花兔':       {'啄木鳥': 120, '家樂福': 0,    '躍獅': 0,   '其他': 0, '庫存捲': 0},
}


# ══════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("達特世生技｜缺料分析 + DRX-3D 叫料試算")
        self.geometry("900x720")
        self.minsize(800, 600)
        self.configure(bg=C['bg'])
        try:
            self.iconbitmap(default='')
        except Exception:
            pass

        self.running     = False
        self.result_path = None
        self._override   = {}

        self._build_ui()
        self._refresh_files()

    # ── UI 建構 ───────────────────────────────────────────
    def _build_ui(self):
        self._build_header()
        # 建立 Notebook
        nb_frame = tk.Frame(self, bg=C['bg'])
        nb_frame.pack(fill='both', expand=True, padx=10, pady=(6,0))

        style = ttk.Style()
        style.configure('TNotebook.Tab', font=('Microsoft JhengHei UI', 10, 'bold'),
                        padding=[12, 5])

        self.nb = ttk.Notebook(nb_frame)
        self.nb.pack(fill='both', expand=True)

        # Tab 1：缺料分析
        tab1 = tk.Frame(self.nb, bg=C['bg'])
        self.nb.add(tab1, text='🔍  缺料分析')
        self._build_shortage_tab(tab1)

        # Tab 2：DRX-3D 叫料試算
        tab2 = tk.Frame(self.nb, bg=C['bg'])
        self.nb.add(tab2, text='📦  DRX-3D 叫料試算')
        self._build_drx_tab(tab2)

        self._build_footer()

    # ── Header ───────────────────────────────────────────
    def _build_header(self):
        h = tk.Frame(self, bg=C['header_bg'])
        h.pack(fill='x')
        tk.Label(h, text="🏭  達特世生技 ‧ 缺料分析系統",
                 font=FONT_TITLE, bg=C['header_bg'], fg=C['header_fg'],
                 pady=10).pack()
        tk.Label(h, text="自動比對庫存與訂單需求 ‧ 輸出缺料清單 Excel　　"
                         "v1.6  2026-05-19（DRX-3D：印刷布損耗6% / 一般布料4%）",
                 font=FONT_SM, bg=C['header_bg'],
                 fg=C['sub_fg'], pady=0).pack(pady=(0, 8))

    # ══════════════════════════════════════════════════════
    # Tab 1：缺料分析（原有功能）
    # ══════════════════════════════════════════════════════
    def _build_shortage_tab(self, parent):
        self._build_files(parent)
        self._build_run_btn(parent)
        self._build_log(parent)
        self._build_stats(parent)

    def _build_files(self, parent):
        outer = tk.Frame(parent, bg=C['bg'], padx=10, pady=8)
        outer.pack(fill='x')
        card = tk.Frame(outer, bg=C['card_bg'], relief='flat', bd=0,
                        highlightbackground='#D5DBDB', highlightthickness=1)
        card.pack(fill='x')
        tk.Label(card, text=" 📁  輸入檔案",
                 font=FONT_BOLD, bg=C['card_bg'], fg='#1A5276',
                 anchor='w', pady=6, padx=10).pack(fill='x')
        ttk.Separator(card, orient='horizontal').pack(fill='x', padx=10)

        self._file_name_vars  = {}
        self._file_name_lbl   = {}
        self._file_status_lbl = {}

        grid = tk.Frame(card, bg=C['card_bg'], padx=12, pady=8)
        grid.pack(fill='x')
        grid.columnconfigure(1, weight=1)

        for i, (key, label, req, _) in enumerate(FILE_DEFS):
            req_color = C['err_fg'] if req == '必要' else C['warn_fg']
            tk.Label(grid, text=f"[{req}]", fg=req_color,
                     font=('Microsoft JhengHei UI', 8, 'bold'),
                     bg=C['card_bg'], width=5, anchor='e'
                     ).grid(row=i, column=0, padx=(0, 4), pady=3, sticky='e')
            tk.Label(grid, text=f"{label}",
                     font=FONT_MAIN, bg=C['card_bg'], width=10, anchor='e'
                     ).grid(row=i, column=1, padx=(0, 6), pady=3, sticky='e')

            var = tk.StringVar()
            self._file_name_vars[key] = var
            nm_lbl = tk.Label(grid, textvariable=var,
                              font=FONT_MONO, bg=C['card_bg'],
                              fg=C['ok_fg'], anchor='w', width=46)
            nm_lbl.grid(row=i, column=2, sticky='ew', pady=3)
            self._file_name_lbl[key] = nm_lbl

            btn = tk.Button(grid, text='更換',
                            command=lambda k=key: self._pick_file(k),
                            font=FONT_SM, bg='#ECF0F1', fg='#333',
                            relief='flat', padx=8, cursor='hand2')
            btn.grid(row=i, column=3, padx=(6, 0), pady=3)

        refresh_row = tk.Frame(card, bg=C['card_bg'])
        refresh_row.pack(fill='x', padx=12, pady=(0, 10))
        tk.Label(refresh_row, text="⚡ 系統自動抓取桌面最新版本",
                 font=FONT_SM, bg=C['card_bg'], fg='#888').pack(side='left')
        tk.Button(refresh_row, text='🔄 重新偵測',
                  command=self._refresh_files,
                  font=FONT_SM, bg='#ECF0F1', fg='#555',
                  relief='flat', padx=8, cursor='hand2').pack(side='right')

    def _build_run_btn(self, parent):
        f = tk.Frame(parent, bg=C['bg'], pady=6)
        f.pack(fill='x', padx=10)
        self.run_btn = tk.Button(
            f, text='▶   執行缺料分析',
            command=self._run,
            font=('Microsoft JhengHei UI', 13, 'bold'),
            bg=C['btn_run'], fg='white', activebackground=C['btn_run_h'],
            activeforeground='white', relief='flat',
            padx=40, pady=10, cursor='hand2')
        self.run_btn.pack(expand=True)
        self.prog = ttk.Progressbar(f, mode='indeterminate', length=600)
        self.prog.pack(fill='x', pady=(6, 0))

    def _build_log(self, parent):
        f = tk.Frame(parent, bg=C['bg'], padx=10)
        f.pack(fill='both', expand=True)
        hdr = tk.Frame(f, bg=C['bg'])
        hdr.pack(fill='x')
        tk.Label(hdr, text="📋  執行紀錄",
                 font=FONT_BOLD, bg=C['bg'], fg='#1A5276').pack(side='left')
        self.status_lbl = tk.Label(hdr, text="", font=FONT_MAIN,
                                    bg=C['bg'], fg='#555')
        self.status_lbl.pack(side='right')

        self.log = scrolledtext.ScrolledText(
            f, font=FONT_MONO, bg=C['log_bg'], fg=C['log_def'],
            insertbackground='white', relief='flat',
            height=10, state='disabled')
        self.log.pack(fill='both', expand=True, pady=(4, 0))

        for tag, fg in [('ok', C['log_ok']), ('err', C['log_err']),
                        ('warn', C['log_warn']), ('info', C['log_info']),
                        ('done', C['log_done'])]:
            kw = {'foreground': fg}
            if tag == 'done':
                kw['font'] = ('Consolas', 9, 'bold')
            self.log.tag_config(tag, **kw)

    def _build_stats(self, parent):
        f = tk.Frame(parent, bg=C['stat_bg'],
                     highlightbackground='#AED6F1', highlightthickness=1,
                     padx=10, pady=8)
        f.pack(fill='x', padx=10, pady=(6, 0))
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

    # ══════════════════════════════════════════════════════
    # Tab 2：DRX-3D 叫料試算
    # ══════════════════════════════════════════════════════
    def _build_drx_tab(self, parent):
        # ── 參數列 ──
        param_frame = tk.Frame(parent, bg='#EDE7F6',
                               highlightbackground='#9C27B0',
                               highlightthickness=1)
        param_frame.pack(fill='x', padx=10, pady=(8, 4))

        tk.Label(param_frame,
                 text="  📐 計算參數（老闆算法 / M號量測值）",
                 font=FONT_BOLD, bg='#EDE7F6', fg='#4A148C',
                 pady=5).pack(side='left')

        # 參數輸入
        self._drx_params = {}
        params = [
            ('pitch',    '步距 Pitch (cm)', 12.5),
            ('roll_m',   '每捲長度 (M)',    1200),
            ('loss_pct', '損耗率 (%)',       6.0),
            ('per_box',  '每盒片數',         20),
        ]
        for key, lbl, default in params:
            tk.Label(param_frame, text=f'  {lbl}:', font=FONT_SM,
                     bg='#EDE7F6', fg='#333').pack(side='left')
            var = tk.StringVar(value=str(default))
            self._drx_params[key] = var
            e = tk.Entry(param_frame, textvariable=var, width=6,
                         font=FONT_MONO, relief='flat',
                         bg='white', fg='#4A148C',
                         highlightbackground='#CE93D8',
                         highlightthickness=1)
            e.pack(side='left', padx=(2, 8))

        # 參數顯示
        self._drx_param_info = tk.Label(
            param_frame, text='', font=FONT_SM, bg='#EDE7F6', fg='#666')
        self._drx_param_info.pack(side='left', padx=6)
        self._update_drx_param_info()

        # 每次參數改變自動更新說明
        for var in self._drx_params.values():
            var.trace_add('write', lambda *_: self._update_drx_param_info())

        # ── 尺寸 Pitch 快選列 ──
        size_frame = tk.Frame(parent, bg='#F3E5F5',
                              highlightbackground='#CE93D8',
                              highlightthickness=1)
        size_frame.pack(fill='x', padx=10, pady=(0, 4))

        tk.Label(size_frame, text='  📏 各尺寸步距（點選自動帶入）：',
                 font=FONT_SM, bg='#F3E5F5', fg='#4A148C',
                 pady=4).pack(side='left')

        # 尺寸定義：(標籤, pitch值, 狀態, 顏色)
        SIZE_PITCHES = [
            ('XL  待確認', None,  '⏳', '#BBBBBB'),
            ('L = 13.5cm', 13.5, '✅', '#1E8449'),
            ('M = 12.5cm', 12.5, '✅', '#1F6FBF'),
            ('S  待確認',  None,  '⏳', '#BBBBBB'),
        ]
        for sz_label, sz_pitch, sz_icon, sz_color in SIZE_PITCHES:
            if sz_pitch is not None:
                btn = tk.Button(
                    size_frame,
                    text=f'{sz_icon} {sz_label}',
                    font=('Microsoft JhengHei UI', 9, 'bold'),
                    bg='white', fg=sz_color,
                    relief='flat', padx=10, pady=3,
                    cursor='hand2',
                    highlightbackground=sz_color,
                    highlightthickness=1,
                    command=lambda p=sz_pitch: (
                        self._drx_params['pitch'].set(str(p)),
                        self._update_drx_param_info()
                    )
                )
                btn.pack(side='left', padx=4, pady=4)
            else:
                tk.Label(size_frame,
                         text=f'{sz_icon} {sz_label}',
                         font=('Microsoft JhengHei UI', 9),
                         bg='#F3E5F5', fg=sz_color,
                         padx=10, pady=4).pack(side='left', padx=4)

        # ── 表格 ──
        tbl_outer = tk.Frame(parent, bg=C['bg'])
        tbl_outer.pack(fill='both', expand=True, padx=10, pady=4)

        # 捲動支援
        canvas = tk.Canvas(tbl_outer, bg=C['bg'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(tbl_outer, orient='vertical',
                                   command=canvas.yview)
        self._drx_tbl = tk.Frame(canvas, bg=C['bg'])
        self._drx_tbl.bind('<Configure>',
            lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=self._drx_tbl, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        self._build_drx_table()

        # ── 按鈕列 ──
        btn_row = tk.Frame(parent, bg=C['bg'], pady=6)
        btn_row.pack(fill='x', padx=10)

        tk.Button(btn_row, text='🔢  計算叫料數量',
                  command=self._drx_calc,
                  font=('Microsoft JhengHei UI', 11, 'bold'),
                  bg='#7B1FA2', fg='white',
                  activebackground='#6A1B9A',
                  relief='flat', padx=20, pady=8,
                  cursor='hand2').pack(side='left', padx=(0, 8))

        tk.Button(btn_row, text='🗑  清空訂單',
                  command=self._drx_clear,
                  font=FONT_MAIN, bg='#B0BEC5', fg='white',
                  relief='flat', padx=12, pady=8,
                  cursor='hand2').pack(side='left', padx=(0, 8))

        tk.Button(btn_row, text='↩  填入上次訂單',
                  command=self._drx_fill_default,
                  font=FONT_MAIN, bg='#546E7A', fg='white',
                  relief='flat', padx=12, pady=8,
                  cursor='hand2').pack(side='left')

        # ── 結果區 ──
        self._drx_result = tk.Label(
            parent, text='← 填入訂單盒數後，點「計算叫料數量」',
            font=FONT_MAIN, bg='#F3E5F5', fg='#555',
            anchor='w', padx=12, pady=8,
            wraplength=860, justify='left')
        self._drx_result.pack(fill='x', padx=10, pady=(0, 8))

    def _update_drx_param_info(self):
        try:
            pitch   = float(self._drx_params['pitch'].get())
            roll_m  = float(self._drx_params['roll_m'].get())
            loss    = float(self._drx_params['loss_pct'].get()) / 100
            per_box = int(self._drx_params['per_box'].get())
            per100  = round(100 * 100 / pitch)
            per_roll = int(roll_m * 100 / pitch * (1 - loss))
            per_roll_box = per_roll // per_box
            self._drx_param_info.config(
                text=f'→  每100M = {per100}片 ｜ 每捲({int(roll_m)}M) = {per_roll:,}片 ≈ {per_roll_box}盒')
        except Exception:
            self._drx_param_info.config(text='（參數格式有誤）')

    def _build_drx_table(self):
        tbl = self._drx_tbl
        for w in tbl.winfo_children():
            w.destroy()

        COLS  = DRX_CHANNELS + ['庫存捲']
        COL_W = [8] * len(DRX_CHANNELS) + [6]  # 各欄寬度
        HDR_BG  = '#6A1B9A'
        HDR_FG  = '#FFFFFF'
        EVEN_BG = '#F3E5F5'
        ODD_BG  = '#FFFFFF'

        # ── 表頭 ──
        hdr_cols = ['品項'] + COLS + ['合計盒', '合計片', '需M數', '捲數', '結果']
        hdr_widths = [12] + COL_W + [8, 9, 8, 5, 16]
        for ci, (h, w) in enumerate(zip(hdr_cols, hdr_widths)):
            bg = '#4A148C' if h in ('合計盒','合計片','需M數','捲數','結果') else HDR_BG
            lbl = tk.Label(tbl, text=h, font=('Microsoft JhengHei UI', 9, 'bold'),
                           bg=bg, fg=HDR_FG, width=w, relief='flat',
                           bd=0, pady=5,
                           highlightbackground='#7B1FA2', highlightthickness=1)
            lbl.grid(row=0, column=ci, padx=1, pady=1, sticky='ew')

        # ── 資料列 ──
        self._drx_entries = {}   # design → {col → Entry/Label var}
        self._drx_result_lbls = {}  # design → Label

        for ri, design in enumerate(DRX_DESIGNS, 1):
            bg = EVEN_BG if ri % 2 == 0 else ODD_BG
            defvals = DRX_DEFAULT.get(design, {})

            # 品項名稱
            tk.Label(tbl, text=design,
                     font=('Microsoft JhengHei UI', 9, 'bold'),
                     bg=bg, fg='#1A237E', width=12, anchor='w',
                     padx=4, pady=4).grid(row=ri, column=0, padx=1, pady=1, sticky='ew')

            entry_row = {}
            all_cols = DRX_CHANNELS + ['庫存捲']
            for ci, col in enumerate(all_cols, 1):
                var = tk.StringVar(value=str(defvals.get(col, 0) or ''))
                entry_row[col] = var
                e = tk.Entry(tbl, textvariable=var, width=hdr_widths[ci],
                             font=FONT_MONO, relief='flat',
                             bg='white', fg='#333',
                             justify='center',
                             highlightbackground='#CE93D8',
                             highlightthickness=1)
                e.grid(row=ri, column=ci, padx=1, pady=1, sticky='ew')
            self._drx_entries[design] = entry_row

            # 合計盒（唯讀計算）
            for ci2, key in enumerate(['合計盒', '合計片', '需M數', '捲數'], len(all_cols)+1):
                lbl = tk.Label(tbl, text='—', font=FONT_MONO,
                               bg=bg, fg='#555', width=hdr_widths[ci2],
                               anchor='e', padx=4)
                lbl.grid(row=ri, column=ci2, padx=1, pady=1, sticky='ew')
                entry_row[key] = lbl

            # 結果標籤
            res_lbl = tk.Label(tbl, text='', font=FONT_SM,
                               bg=bg, fg='#1A5276', width=16,
                               anchor='w', padx=4)
            res_lbl.grid(row=ri, column=len(all_cols)+5, padx=1, pady=1, sticky='ew')
            self._drx_result_lbls[design] = res_lbl

    def _get_drx_param(self, key, default):
        try:
            return float(self._drx_params[key].get())
        except Exception:
            return default

    def _drx_calc(self):
        pitch   = self._get_drx_param('pitch', 12.5)
        roll_m  = self._get_drx_param('roll_m', 1200)
        loss    = self._get_drx_param('loss_pct', 0.4) / 100
        per_box = int(self._get_drx_param('per_box', 20))

        per_roll_pieces = roll_m * 100 / pitch * (1 - loss)

        results = []
        for design in DRX_DESIGNS:
            row = self._drx_entries[design]
            # 加總各通路
            total_boxes = 0
            for ch in DRX_CHANNELS:
                try:
                    total_boxes += int(row[ch].get() or 0)
                except Exception:
                    pass
            stock_rolls = 0
            try:
                stock_rolls = float(row['庫存捲'].get() or 0)
            except Exception:
                pass

            total_pieces = total_boxes * per_box
            m_need = total_pieces * pitch / 100 / (1 - loss) if total_pieces > 0 else 0
            rolls_need = math.ceil(m_need / roll_m) if m_need > 0 else 0

            # 扣庫存
            stock_boxes = int(stock_rolls * per_roll_pieces / per_box)
            rolls_order = max(0, rolls_need - int(stock_rolls))

            # 更新唯讀欄位
            EVEN_BG = '#F3E5F5'; ODD_BG = '#FFFFFF'
            ri = DRX_DESIGNS.index(design) + 1
            bg = EVEN_BG if ri % 2 == 0 else ODD_BG

            row['合計盒'].config(text=f'{total_boxes:,}' if total_boxes else '—',
                                 fg='#7B1FA2' if total_boxes else '#BBB')
            row['合計片'].config(text=f'{total_pieces:,}' if total_pieces else '—',
                                 fg='#333')
            row['需M數'].config(text=f'{math.ceil(m_need):,}' if m_need else '—',
                                fg='#E65100' if m_need else '#BBB')
            roll_color = '#7B1FA2' if rolls_need > 0 else '#BBB'
            row['捲數'].config(text=str(rolls_need) if rolls_need else '—',
                               fg=roll_color,
                               font=('Microsoft JhengHei UI', 10, 'bold'))

            # 結果說明
            if total_boxes == 0:
                res_text, res_color = '（無訂單）', '#AAAAAA'
            elif stock_rolls >= rolls_need:
                res_text  = f'✅ 庫存足 (有{int(stock_rolls)}捲)'
                res_color = '#1E8449'
            elif stock_rolls > 0:
                res_text  = f'▶ 叫 {rolls_order}捲（有庫存{int(stock_rolls)}捲）'
                res_color = '#E65100'
            else:
                res_text  = f'▶ 需叫 {rolls_need} 捲'
                res_color = '#7B1FA2'

            self._drx_result_lbls[design].config(text=res_text, fg=res_color,
                                                   bg='#E8F5E9' if '✅' in res_text else bg)
            results.append((design, total_boxes, rolls_need, rolls_order,
                            stock_rolls, res_text))

        # ── 彙整摘要 ──
        active = [(d, b, rn, ro, sk, tx)
                  for d, b, rn, ro, sk, tx in results if b > 0]
        if not active:
            self._drx_result.config(
                text='⚠ 請先填入訂單盒數', fg='#E65100', bg='#FFF3E0')
            return

        max_design = max(active, key=lambda x: x[1])
        total_order_rolls = sum(ro for _, _, _, ro, _, _ in active)
        lines = [
            f"📊 計算完成  ─  Pitch={pitch}cm ｜ 每捲{int(roll_m)}M ｜ 印刷布損耗{loss*100:.1f}%",
            f"",
            f"⭐ 片數最多：{max_design[0]}（{max_design[1]:,}盒）→ 基準捲數 {max_design[2]}捲",
            f"📦 本次共需叫料：{total_order_rolls} 捲",
            f"",
        ]
        for d, b, rn, ro, sk, tx in active:
            sk_note = f'（庫存{int(sk)}捲，已扣）' if sk > 0 else ''
            lines.append(f"   {'▶' if ro > 0 else '✅'}  {d:<10} {b:>5}盒 → 需{rn}捲{sk_note}  {tx}")

        self._drx_result.config(
            text='\n'.join(lines), fg='#1A237E', bg='#EDE7F6')

    def _drx_clear(self):
        for design in DRX_DESIGNS:
            row = self._drx_entries[design]
            for col in DRX_CHANNELS + ['庫存捲']:
                row[col].set('')
            for col in ['合計盒', '合計片', '需M數', '捲數']:
                row[col].config(text='—', fg='#BBB')
            self._drx_result_lbls[design].config(text='', bg='#FFFFFF')
        self._drx_result.config(
            text='← 已清空。請填入訂單盒數後點「計算叫料數量」',
            fg='#555', bg='#F3E5F5')

    def _drx_fill_default(self):
        for design in DRX_DESIGNS:
            row = self._drx_entries[design]
            defvals = DRX_DEFAULT.get(design, {})
            for col in DRX_CHANNELS + ['庫存捲']:
                val = defvals.get(col, 0)
                row[col].set(str(val) if val else '')
        self._drx_result.config(
            text='✅ 已填入 5/15 訂單資料，點「計算叫料數量」試算',
            fg='#7B1FA2', bg='#F3E5F5')

    # ── 共用 Footer ───────────────────────────────────────
    def _build_footer(self):
        f = tk.Frame(self, bg=C['bg'], padx=16, pady=8)
        f.pack(fill='x')
        self.open_btn = tk.Button(
            f, text='📂  開啟結果 Excel',
            command=self._open_result,
            font=FONT_BOLD, bg=C['btn_open'], fg='white',
            activebackground='#1A6B3C', activeforeground='white',
            relief='flat', padx=16, pady=7,
            state='disabled', cursor='hand2')
        self.open_btn.pack(side='left', padx=(0, 8))
        self.dir_btn = tk.Button(
            f, text='📁  開啟資料夾',
            command=self._open_dir,
            font=FONT_MAIN, bg=C['btn_dir'], fg='white',
            activebackground='#4D5656', activeforeground='white',
            relief='flat', padx=12, pady=7, cursor='hand2')
        self.dir_btn.pack(side='left')
        self.time_lbl = tk.Label(f, text="",
                                  font=FONT_SM, bg=C['bg'], fg='#888')
        self.time_lbl.pack(side='right')

    # ── 檔案偵測 ──────────────────────────────────────────
    def _refresh_files(self):
        found = detect_all()
        for key, label, req, _ in FILE_DEFS:
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
        self._file_name_lbl[key].config(fg='#2E86C1')

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

    # ── 執行缺料分析 ──────────────────────────────────────
    def _run(self):
        if self.running:
            return
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
            f"▶  開始執行缺料分析  [{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]", 'info')
        self._log_write('─' * 64)

        def worker():
            _copied = []
            try:
                for key, *_ in FILE_DEFS:
                    if key in self._override:
                        src = self._override[key]
                        ext = os.path.splitext(src)[1]
                        prefixes = {'inv': '原物料庫存', 'order': '訂單生產總表',
                                    'ear': '耳繩庫存', 'box': '盒子庫存'}
                        dst = os.path.join(DESK, f"_gui_override_{prefixes[key]}{ext}")
                        shutil.copy2(src, dst)
                        os.utime(dst, None)
                        _copied.append(dst)
            except Exception as ce:
                self.after(0, self._log_write, f'⚠  複製覆寫檔案失敗: {ce}', 'warn')

            try:
                proc = subprocess.Popen(
                    [PYTHON, '-u', SCRIPT],
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    cwd=BASE, encoding='utf-8', errors='replace', bufsize=1)

                result_path = None
                orders = materials = shortage = None

                for raw in proc.stdout:
                    line = raw.rstrip('\r\n')
                    tag  = self._classify(line)
                    self.after(0, self._log_write, line, tag)
                    m = re.search(r'輸出:\s*(.+\.xlsx)', line)
                    if m:
                        result_path = m.group(1).strip()
                    m2 = re.search(
                        r'訂單:\s*(\d+)\s*筆.*材料:\s*(\d+)\s*項.*缺料:\s*(\d+)\s*項', line)
                    if m2:
                        orders, materials, shortage = m2.groups()

                proc.wait()
                for f in _copied:
                    try: os.remove(f)
                    except Exception: pass

                self.after(0, self._finish, proc.returncode, result_path,
                           orders, materials, shortage)
            except Exception as e:
                self.after(0, self._log_write, f'❌  執行錯誤: {e}', 'err')
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
            self.status_lbl.config(text=f'✅ 完成 {now}', fg=C['ok_fg'])
            self.time_lbl.config(text=f'完成時間：{now}')
            if orders:   self._stat_vars['orders'].set(orders)
            if materials: self._stat_vars['materials'].set(materials)
            if shortage:
                self._stat_vars['shortage'].set(shortage)
                n = int(shortage)
                color = '#922B21' if n > 0 else C['ok_fg']
                for w in self.winfo_children():
                    self._recolor_stat(w, 'shortage', color)
        else:
            self._log_write(f'⚠  執行結束（代碼: {rc}）請檢查上方 log', 'warn')
            self.status_lbl.config(text=f'⚠ 請檢查 log  {now}', fg=C['warn_fg'])

    def _recolor_stat(self, widget, key, color):
        try:
            if isinstance(widget, tk.Label):
                if str(widget.cget('textvariable')) == str(self._stat_vars[key]):
                    widget.config(fg=color)
        except Exception:
            pass
        for child in widget.winfo_children():
            self._recolor_stat(child, key, color)

    def _open_result(self):
        if self.result_path and os.path.exists(self.result_path):
            os.startfile(self.result_path)
        else:
            messagebox.showwarning('找不到檔案', f'找不到結果檔案：\n{self.result_path}')

    def _open_dir(self):
        today  = datetime.now().strftime('%m%d')
        outdir = os.path.join(BASE, f'缺料分析{today}')
        target = outdir if os.path.exists(outdir) else BASE
        os.startfile(target)


# ══════════════════════════════════════════════════════════
if __name__ == '__main__':
    app = App()
    app.mainloop()

"""
Startup Settings UI (Tkinter)
- Provide detailed Chinese descriptions for each parameter.
- Load defaults from config.settings and config.runtime (JSON)
- Save to runtime JSON and apply to process.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from typing import Dict, Any
import config.settings as settings
from config.runtime import load_runtime_settings, save_runtime_settings, apply_to_settings

PARAMS_SPEC = [
    {
        'key': 'WATCH_FOLDERS',
        'label': '監控資料夾 (每行一個路徑)',
        'help': '指定需要監控的資料夾（或檔案）。支援網路磁碟。系統會遞迴監控子資料夾。',
        'type': 'multiline',
    },
    {
        'key': 'SUPPORTED_EXTS',
        'label': '檔案類型 (Excel 為 .xlsx,.xlsm)',
        'help': '設定需要監控的檔案副檔名，逗號分隔（例如 .xlsx,.xlsm）。',
        'type': 'text',
    },
    {
        'key': 'SCAN_ALL_MODE',
        'label': '啟動時掃描所有 Excel 並建立基準線',
        'help': '開啟後，啟動時會掃描 WATCH_FOLDERS 內所有支援檔案並建立初始基準線。關閉可縮短大型磁碟啟動時間。',
        'type': 'bool',
    },
    {
        'key': 'USE_LOCAL_CACHE',
        'label': '啟用本地快取',
        'help': '讀取網路檔前先複製到本地快取，提高穩定性與速度。',
        'type': 'bool',
    },
    {
        'key': 'CACHE_FOLDER',
        'label': '本地快取資料夾',
        'help': '設定本地快取位置。需具備讀寫權限。',
        'type': 'text',
    },
    {
        'key': 'ENABLE_TIMEOUT',
        'label': '啟用檔案處理超時保護',
        'help': '當單一檔案處理超過 FILE_TIMEOUT_SECONDS 時中止該檔處理，避免長時間卡住。',
        'type': 'bool',
    },
    {
        'key': 'FILE_TIMEOUT_SECONDS',
        'label': '單檔超時秒數',
        'help': '超過此秒數仍未完成讀取/比較會視為超時。',
        'type': 'int',
    },
    {
        'key': 'ENABLE_MEMORY_MONITOR',
        'label': '啟用記憶體監控',
        'help': '當行程記憶體超過限制時自動觸發垃圾回收並告警。',
        'type': 'bool',
    },
    {
        'key': 'MEMORY_LIMIT_MB',
        'label': '記憶體上限 (MB)',
        'help': '超過此數值時會嘗試釋放記憶體並提示。',
        'type': 'int',
    },
    {
        'key': 'DEFAULT_COMPRESSION_FORMAT',
        'label': '基準線壓縮格式',
        'help': '選擇基準線儲存格式：lz4 (讀寫快), zstd (壓縮高), gzip (相容性)。',
        'type': 'choice',
        'choices': ['lz4','zstd','gzip']
    },
    {
        'key': 'AUTO_UPDATE_BASELINE_AFTER_COMPARE',
        'label': '比較後自動更新基準線',
        'help': '當偵測到變更後，是否自動以「目前內容」更新成新的基準線。',
        'type': 'bool',
    },
    {
        'key': 'TRACK_DIRECT_VALUE_CHANGES',
        'label': '追蹤直接值變更',
        'help': '勾選後，若某格為輸入文字/數字（非公式），其值變更會被記錄為「直接值變更」。',
        'type': 'bool',
    },
    {
        'key': 'TRACK_FORMULA_CHANGES',
        'label': '追蹤公式變更',
        'help': '勾選後，只要儲存格的公式字串有改動（例如 =A1+B1 → =A1+B2）便會記錄為「公式變更」。',
        'type': 'bool',
    },
    {
        'key': 'TRACK_EXTERNAL_REFERENCES',
        'label': '追蹤外部參照更新',
        'help': '勾選後，若儲存格的公式不變，但其引用的外部連結刷新導致結果變動，會記錄為「外部參照更新」。',
        'type': 'bool',
    },
    {
        'key': 'IGNORE_INDIRECT_CHANGES',
        'label': '忽略間接影響變更',
        'help': '勾選後，若僅因工作簿內其他儲存格改動導致此格公式結果變化（公式本身不變），則忽略此類變更。',
        'type': 'bool',
    },
    {
        'key': 'MAX_CHANGES_TO_DISPLAY',
        'label': '畫面顯示變更上限 (0=不限制)',
        'help': '限制 console 表格一次展示的變更數，有助於大檔案閱讀。',
        'type': 'int',
    },
    {
        'key': 'POLLING_SIZE_THRESHOLD_MB',
        'label': '輪詢大小分界 (MB)',
        'help': '小於此大小的檔案採用較密集的輪詢間隔；大於則採用較稀疏的間隔。',
        'type': 'int',
    },
    {
        'key': 'DENSE_POLLING_INTERVAL_SEC',
        'label': '密集輪詢間隔 (秒)',
        'help': '適用於較小檔案的輪詢頻率。',
        'type': 'int',
    },
    {
        'key': 'DENSE_POLLING_DURATION_SEC',
        'label': '密集輪詢總時長 (秒)',
        'help': '沒有進一步變更時，密集輪詢會在總時長用盡後停止。',
        'type': 'int',
    },
    {
        'key': 'SPARSE_POLLING_INTERVAL_SEC',
        'label': '稀疏輪詢間隔 (秒)',
        'help': '適用於較大檔案的輪詢頻率。',
        'type': 'int',
    },
    {
        'key': 'SPARSE_POLLING_DURATION_SEC',
        'label': '稀疏輪詢總時長 (秒)',
        'help': '如需使用舊版 watcher 的稀疏輪詢策略可參考 legacy；現版本用自適應穩定檢查。',
        'type': 'int',
    },
]

class SettingsDialog(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title('Excel Watchdog 設定')
        self.geometry('900x700')
        self.grab_set()
        self._widgets: Dict[str, Any] = {}

        # Load defaults (config + runtime overrides)
        runtime = load_runtime_settings()

        frm = ttk.Frame(self)
        frm.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(frm)
        scrollbar = ttk.Scrollbar(frm, orient='vertical', command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=scroll_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        for spec in PARAMS_SPEC:
            row = ttk.Frame(scroll_frame)
            row.pack(fill='x', padx=10, pady=6)
            ttk.Label(row, text=spec['label']).pack(anchor='w')
            help_lbl = ttk.Label(row, text=spec['help'], foreground='#666')
            help_lbl.pack(anchor='w')
            key = spec['key']
            cur_val = getattr(settings, key, '')
            if key in runtime:
                cur_val = runtime[key]
            w = None
            if spec['type'] == 'text':
                # Special display for SUPPORTED_EXTS: show as comma-separated string without parentheses
                display_val = ''
                if key == 'SUPPORTED_EXTS':
                    if isinstance(cur_val, (list, tuple)):
                        display_val = ','.join([str(x) for x in cur_val])
                    else:
                        display_val = str(cur_val)
                else:
                    display_val = str(cur_val)
                var = tk.StringVar(value=display_val)
                w = ttk.Entry(row, textvariable=var, width=80)
                w.pack(anchor='w', fill='x')
            elif spec['type'] == 'multiline':
                text = tk.Text(row, height=4, width=80)
                if isinstance(cur_val, (list, tuple)):
                    text.insert('1.0', '\n'.join(cur_val))
                else:
                    text.insert('1.0', str(cur_val))
                text.pack(anchor='w', fill='x')
                w = text
            elif spec['type'] == 'bool':
                var = tk.BooleanVar(value=bool(cur_val))
                w = ttk.Checkbutton(row, variable=var, text='啟用/勾選')
                w.var = var
                w.pack(anchor='w')
            elif spec['type'] == 'int':
                var = tk.StringVar(value=str(cur_val))
                w = ttk.Entry(row, textvariable=var, width=20)
                w.pack(anchor='w')
            elif spec['type'] == 'choice':
                var = tk.StringVar(value=str(cur_val))
                w = ttk.Combobox(row, textvariable=var, values=spec['choices'], state='readonly', width=20)
                w.pack(anchor='w')
            self._widgets[key] = (spec, w)

        btn_row = ttk.Frame(self)
        btn_row.pack(fill='x', padx=10, pady=10)
        ttk.Button(btn_row, text='選擇資料夾...', command=self._pick_folder).pack(side='left')
        ttk.Button(btn_row, text='還原預設', command=self._reset_defaults).pack(side='left', padx=6)
        ttk.Button(btn_row, text='儲存並開始', command=self._save_and_apply).pack(side='right')

    def _pick_folder(self):
        folder = filedialog.askdirectory()
        if not folder:
            return
        key = 'WATCH_FOLDERS'
        spec, widget = self._widgets.get(key, (None, None))
        if widget is None:
            return
        if spec['type'] == 'multiline':
            existing = widget.get('1.0', 'end').strip()
            lines = [l for l in existing.split('\n') if l.strip()]
            lines.append(folder)
            widget.delete('1.0', 'end')
            widget.insert('1.0', '\n'.join(lines))

    def _reset_defaults(self):
        for key, (spec, widget) in self._widgets.items():
            val = getattr(settings, key, '')
            if spec['type'] == 'text':
                widget.delete(0, 'end')
                widget.insert(0, str(val))
            elif spec['type'] == 'multiline':
                widget.delete('1.0', 'end')
                if isinstance(val, (list, tuple)):
                    widget.insert('1.0', '\n'.join(val))
                else:
                    widget.insert('1.0', str(val))
            elif spec['type'] == 'bool':
                widget.var.set(bool(val))
            elif spec['type'] == 'int':
                widget.delete(0, 'end')
                widget.insert(0, str(val))
            elif spec['type'] == 'choice':
                widget.set(str(val))

    def _collect_values(self) -> Dict[str, Any]:
        data: Dict[str, Any] = {}
        for key, (spec, widget) in self._widgets.items():
            if spec['type'] == 'text' or spec['type'] == 'int' or spec['type'] == 'choice':
                data[key] = widget.get().strip()
            elif spec['type'] == 'multiline':
                raw = widget.get('1.0', 'end').strip()
                lines = [l.strip() for l in raw.split('\n') if l.strip()]
                data[key] = lines
            elif spec['type'] == 'bool':
                data[key] = bool(widget.var.get())
        # normalize SUPPORTED_EXTS string to tuple-like list
        exts = data.get('SUPPORTED_EXTS')
        if isinstance(exts, str):
            items = [x.strip() for x in exts.split(',') if x.strip()]
            norm = []
            for x in items:
                x = x.strip(" ' \"(){}").lower()
                if not x:
                    continue
                if not x.startswith('.'):
                    x = '.' + x
                norm.append(x)
            if norm:
                data['SUPPORTED_EXTS'] = norm
            else:
                # Do not override if user left it blank
                data.pop('SUPPORTED_EXTS', None)
        return data

    def _save_and_apply(self):
        try:
            data = self._collect_values()
            # persist and apply
            save_runtime_settings(data)
            apply_to_settings(data)
            self.destroy()
        except Exception as e:
            messagebox.showerror('錯誤', f'儲存設定失敗: {e}')


def show_settings_ui():
    root = tk.Tk()
    root.withdraw()
    dlg = SettingsDialog(root)
    root.wait_window(dlg)
    root.destroy()

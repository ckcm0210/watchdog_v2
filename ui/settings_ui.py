"""
Startup Settings UI (Tkinter)
- Provide detailed Chinese descriptions for each parameter.
- Load defaults from config.settings and config.runtime (JSON)
- Save to runtime JSON and apply to process.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os, json, sys
from typing import Dict, Any
import config.settings as settings
import config.runtime as runtime
from config.runtime import load_runtime_settings, save_runtime_settings, apply_to_settings

PARAMS_SPEC = [
    # 監控與檔案類型
    {
        'key': 'WATCH_FOLDERS',
        'priority': 1,
        'label': '監控資料夾（可多個）',
        'help': '指定需要監控的資料夾（可多個）。支援網路磁碟。系統會遞迴監控子資料夾。可用下方「新增資料夾」按鈕加入。',
        'type': 'paths',
        'path_kind': 'dir',
    },
    {
        'key': 'SUPPORTED_EXTS',
        'label': '檔案類型 (Excel 為 .xlsx,.xlsm)',
        'help': '設定需要監控的檔案副檔名，逗號分隔（例如 .xlsx,.xlsm）。會自動正規化為小寫並加上點號。',
        'type': 'text',
    },
    {
        'key': 'MANUAL_BASELINE_TARGET',
        'priority': 3,
        'label': '手動建立基準線的檔案清單',
        'help': '啟動時會先對這些檔案建立基準線（可多個）。使用「新增檔案」加入。',
        'type': 'paths',
        'path_kind': 'file',
    },
    {
        'key': 'MONITOR_ONLY_FOLDERS',
        'priority': 4,
        'label': '只監控變更的根目錄（Issue B）',
        'help': '在此清單內的根目錄底下，第一次偵測到 Excel 檔變更時，系統只會記錄路徑、最後修改時間與最後儲存者，並建立首次基準線；下一次變更才進入普通比較流程。若某子資料夾同時在 WATCH_FOLDERS，則以 WATCH_FOLDERS 的即時比較為優先。',
        'type': 'paths',
        'path_kind': 'dir',
    },
    {
        'key': 'WATCH_EXCLUDE_FOLDERS',
        'priority': 2,
        'label': '即時比較的排除清單（子資料夾）',
        'help': '若在 WATCH_FOLDERS 中，這些子資料夾會被排除，不進行即時比較。',
        'type': 'paths',
        'path_kind': 'dir',
    },
    {
        'key': 'MONITOR_ONLY_EXCLUDE_FOLDERS',
        'priority': 5,
        'label': '只監控變更的排除清單（子資料夾）',
        'help': '若在 MONITOR_ONLY_FOLDERS 中，這些子資料夾會被排除，不進行 monitor-only。',
        'type': 'paths',
        'path_kind': 'dir',
    },
    {
        'key': 'SCAN_TARGET_FOLDERS',
        'priority': 3,
        'label': '啟動掃描的指定目錄（可多個）',
        'help': '啟動掃描時建立基準線的目錄清單。預設會以 WATCH_FOLDERS 全部為準；你可在此列表移除不想掃描的目錄或自行新增。',
        'type': 'paths',
        'path_kind': 'dir',
    },
    {
        'key': 'AUTO_SYNC_SCAN_TARGETS',
        'priority': 3,
        'label': '啟動掃描清單自動同步監控資料夾',
        'help': '開啟後，「啟動掃描的指定目錄」會自動與「監控資料夾」一致；關閉可手動指定子集。',
        'type': 'bool',
    },
    {
        'key': 'SCAN_ALL_MODE',
        'priority': 3,
        'label': '啟動時掃描所有 Excel 並建立基準線',
        'help': '開啟後，啟動時會掃描 WATCH_FOLDERS 內所有支援檔案並建立初始基準線。關閉可縮短大型磁碟啟動時間。',
        'type': 'bool',
    },

    # 快取與暫存
    {
        'key': 'USE_LOCAL_CACHE',
        'label': '啟用本地快取',
        'help': '讀取網路檔前先複製到本地快取，提高穩定性與速度。',
        'type': 'bool',
    },
    {
        'key': 'CACHE_FOLDER',
        'priority': 3,
        'label': '本地快取資料夾',
        'help': '設定本地快取位置。需具備讀寫權限。可透過「瀏覽」選擇資料夾。',
        'type': 'path',
        'path_kind': 'dir',
    },

    # 超時/記憶體/恢復
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
        'key': 'ENABLE_RESUME',
        'label': '啟用進度恢復',
        'help': '建立大量基準線時，將進度寫入 RESUME_LOG_FILE，重新啟動可續傳。',
        'type': 'bool',
    },
    {
        'key': 'RESUME_LOG_FILE',
        'priority': 4,
        'label': '進度紀錄檔路徑',
        'help': '保存基準線建立進度的檔案路徑，建議放在本機磁碟。',
        'type': 'path',
        'path_kind': 'save_file',
    },

    # 監視邏輯/防抖
    {
        'key': 'DEBOUNCE_INTERVAL_SEC',
        'label': '防抖動間隔 (秒)',
        'help': '相同檔案在短時間內多次事件，會合併為一次。',
        'type': 'int',
    },

    # 比較邏輯
    {
        'key': 'FORMULA_ONLY_MODE',
        'label': '只關注公式變更',
        'help': '啟用後，僅比較與顯示公式的變更。',
        'type': 'bool',
    },
    {
        'key': 'TRACK_DIRECT_VALUE_CHANGES',
        'label': '追蹤直接值變更',
        'help': '若某格為輸入文字/數字（非公式），其值變更會被記錄。',
        'type': 'bool',
    },
    {
        'key': 'TRACK_FORMULA_CHANGES',
        'label': '追蹤公式變更',
        'help': '只要儲存格的公式字串有改動（例如 =A1+B1 → =A1+B2）便會記錄。',
        'type': 'bool',
    },
    {
        'key': 'ENABLE_FORMULA_VALUE_CHECK',
        'label': '外部參照：值不變視為無變更',
        'help': '當外部參照公式的字串因刷新而有差異，但其儲存的數值（cached value）沒有改變時，忽略該變更（避免假警報）。只對快取副本進行 read-only 讀取。',
        'type': 'bool',
    },
    {
        'key': 'MAX_FORMULA_VALUE_CELLS',
        'label': '值比對的最大公式格數（跨表合計）',
        'help': '為了效能，只對前 N 個含公式的儲存格查詢其 cached value。超過此數量時跳過值比對（仍會比較公式字串）。',
        'type': 'int',
    },
    {
        'key': 'TRACK_EXTERNAL_REFERENCES',
        'label': '追蹤外部參照更新',
        'help': '公式不變、但外部連結刷新導致結果變更時記錄。',
        'type': 'bool',
    },
    {
        'key': 'IGNORE_INDIRECT_CHANGES',
        'label': '忽略間接影響變更',
        'help': '公式不變、僅因工作簿內其他儲存格改動導致結果變化時忽略。',
        'type': 'bool',
    },
    {
        'key': 'MAX_CHANGES_TO_DISPLAY',
        'label': '畫面顯示變更上限 (0=不限制)',
        'help': '限制 console 表格一次展示的變更數，有助於大檔案閱讀。',
        'type': 'int',
    },
    {
        'key': 'AUTO_UPDATE_BASELINE_AFTER_COMPARE',
        'label': '比較後自動更新基準線',
        'help': '當偵測到變更後，是否自動以「目前內容」更新成新的基準線。',
        'type': 'bool',
    },

    # 壓縮/歸檔
    {
        'key': 'DEFAULT_COMPRESSION_FORMAT',
        'label': '基準線壓縮格式',
        'help': '選擇基準線儲存格式：lz4 (讀寫快), zstd (壓縮高), gzip (相容性)。',
        'type': 'choice',
        'choices': ['lz4','zstd','gzip']
    },
    {
        'key': 'LZ4_COMPRESSION_LEVEL',
        'label': 'LZ4 壓縮等級 (0-16)',
        'help': '速度最快，讀取幾乎不受影響。等級越高壓縮率越好但寫入較慢：0-3 非常快、壓縮較低；4-9 折衝；10-16 壓縮更高但明顯變慢。',
        'type': 'int',
    },
    {
        'key': 'ZSTD_COMPRESSION_LEVEL',
        'label': 'Zstd 壓縮等級 (1-22)',
        'help': '高壓縮比通用首選：1-3 偏速度；4-9 折衝（常用 3-6）；10-18 更高壓縮但寫入耗時；19-22 極致壓縮、CPU時間很長（僅限空間極度敏感）。',
        'type': 'int',
    },
    {
        'key': 'GZIP_COMPRESSION_LEVEL',
        'label': 'gzip 壓縮等級 (1-9)',
        'help': '兼容性最佳：1-3 較快、壓縮一般；4-6 折衝（6 常用）；7-9 壓縮略升但耗時顯著，除非相容性/可攜性為先。',
        'type': 'int',
    },
    {
        'key': 'ENABLE_ARCHIVE_MODE',
        'label': '啟用歸檔模式',
        'help': '基準線建立一段時間後可轉存為較高壓縮格式以節省空間。',
        'type': 'bool',
    },
    {
        'key': 'ARCHIVE_AFTER_DAYS',
        'label': '轉為歸檔的天數',
        'help': '建立後超過此天數的基準線將轉為歸檔格式。',
        'type': 'int',
    },
    {
        'key': 'ARCHIVE_COMPRESSION_FORMAT',
        'label': '歸檔壓縮格式',
        'help': '歸檔時使用的壓縮格式。',
        'type': 'choice',
        'choices': ['lz4','zstd','gzip']
    },
    {
        'key': 'SHOW_COMPRESSION_STATS',
        'label': '顯示壓縮統計',
        'help': '在儲存/讀取基準線時顯示壓縮比與耗時。',
        'type': 'bool',
    },
    {
        'key': 'SHOW_DEBUG_MESSAGES',
        'label': '顯示除錯訊息',
        'help': '輸出較詳細的內部流程訊息。',
        'type': 'bool',
    },

    # 日誌/輸出
    {
        'key': 'LOG_FOLDER',
        'priority': 5,
        'label': '日誌資料夾（CSV/日誌輸出）',
        'help': '記錄 CSV 與其他日誌的資料夾。',
        'type': 'path',
        'path_kind': 'dir',
    },
    {
        'key': 'LOG_FILE_DATE',
        'label': '日誌日期（唯讀）',
        'help': '用於組合 CSV_LOG_FILE 的日期字串。',
        'type': 'readonly',
    },
    {
        'key': 'CSV_LOG_FILE',
        'label': 'CSV 記錄檔（唯讀）',
        'help': '比較結果輸出的壓縮 CSV 檔路徑，從 LOG_FOLDER + 日期組合而來。',
        'type': 'readonly',
    },
    {
        'key': 'CONSOLE_TEXT_LOG_ENABLED',
        'label': '將 Console 輸出寫入文字檔',
        'help': '將所有 Console 訊息（含表格）追加寫入指定的文字檔（UTF-8）。可用於長期留存或排查。',
        'type': 'bool',
    },
    {
        'key': 'CONSOLE_TEXT_LOG_FILE',
        'label': 'Console 文字檔路徑',
        'help': '預設為 LOG_FOLDER/console_log_YYYYMMDD.txt（依據日誌日期）。你亦可自訂儲存位置。',
        'type': 'path',
        'path_kind': 'save_file',
    },

    # 快速模式
    {
        'key': 'ENABLE_FAST_MODE',
        'label': '啟用快速模式',
        'help': '針對常見情境優化，可能略過部分詳細檢查以加速。',
        'type': 'bool',
    },

    # 重試/臨時複本
    {
        'key': 'MAX_RETRY',
        'label': '失敗重試次數',
        'help': '讀取或比較失敗時最多重試次數。',
        'type': 'int',
    },
    {
        'key': 'RETRY_INTERVAL_SEC',
        'label': '重試間隔 (秒)',
        'help': '兩次重試之間等待的秒數。',
        'type': 'int',
    },
    {
        'key': 'USE_TEMP_COPY',
        'label': '使用暫存複本',
        'help': '比較前先複製檔案到暫存位置以避免占用與鎖檔。',
        'type': 'bool',
    },

    # 進階／穩定性（複製與嚴格模式）
    {
        'key': 'STRICT_NO_ORIGINAL_READ',
        'label': '嚴格模式：永不開原始檔',
        'help': '啟用後，任何讀取都必須來自本地快取副本；若複製到快取最終失敗，會略過本次處理，絕不打開原始檔（避免鎖檔）。',
        'type': 'bool',
    },
    {
        'key': 'IGNORE_CACHE_FOLDER',
        'label': '忽略快取資料夾事件',
        'help': '忽略 CACHE_FOLDER 內所有檔案事件，避免快取本身引起的監控噪音（建議啟用）。',
        'type': 'bool',
    },
    {
        'key': 'COPY_RETRY_COUNT',
        'label': '複製重試次數',
        'help': '將來源檔案複製到本地快取時的最大重試次數。來源檔案正被儲存／網路不穩時可提高此值。',
        'type': 'int',
    },
    {
        'key': 'COPY_RETRY_BACKOFF_SEC',
        'label': '重試退避（秒）',
        'help': '兩次重試之間的等待秒數（可輸入小數）。會按嘗試次數逐步增加等待，例如 1.0 → 2.0 → 3.0 秒。',
        'type': 'text',
    },
    {
        'key': 'COPY_CHUNK_SIZE_MB',
        'label': '分塊複製大小 (MB)',
        'help': '以較小區塊逐段讀寫來源檔，可降低一次性長時間把持來源句柄的風險。0 表示關閉。',
        'type': 'int',
    },
    {
        'key': 'COPY_POST_SLEEP_SEC',
        'label': '複製完成後短暫等待（秒）',
        'help': '複製完成後等待短暫時間（可輸入小數）讓檔案系統穩定，避免隨後讀取時競態。',
        'type': 'text',
    },
    {
        'key': 'COPY_STABILITY_CHECKS',
        'label': '複製前穩定性檢查次數',
        'help': '開始複製前，連續 N 次檢查來源檔的修改時間（mtime）一致才開始複製。',
        'type': 'int',
    },
    {
        'key': 'COPY_STABILITY_INTERVAL_SEC',
        'label': '穩定性檢查間隔（秒）',
        'help': '兩次 mtime 穩定性檢查之間的等待秒數（可輸入小數）。',
        'type': 'text',
    },
    {
        'key': 'COPY_STABILITY_MAX_WAIT_SEC',
        'label': '穩定性檢查最大等待（秒）',
        'help': '最多等待多少秒來達到所需的穩定檢查次數；超時則本次複製跳過。',
        'type': 'text',
    },

    # 日誌／去重
    {
        'key': 'LOG_DEDUP_WINDOW_SEC',
        'label': 'CSV 去重時間窗（秒）',
        'help': '在此秒數內，相同檔案＋工作表＋相同內容的變更只記錄一次至 CSV（避免短時間重覆記錄同一批變更）。',
        'type': 'int',
    },

    # 白名單
    {
        'key': 'WHITELIST_USERS',
        'label': '使用者白名單 (每行一個)',
        'help': '在白名單內的使用者修改可選擇不顯示或單獨記錄。',
        'type': 'multiline',
    },
    {
        'key': 'LOG_WHITELIST_USER_CHANGE',
        'label': '記錄白名單使用者變更',
        'help': '啟用後，白名單使用者的變更也會寫入記錄。',
        'type': 'bool',
    },
    {
        'key': 'FORCE_BASELINE_ON_FIRST_SEEN',
        'label': '首次遇見即強制建立基準線 (每行一個關鍵字)',
        'help': '支援關鍵字或部分路徑比對。若檔案路徑包含其一，第一次掃描或偵測到時即建立基準線。',
        'type': 'multiline',
    },

    # 輪詢
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

    # Console 視窗
    {
        'key': 'ENABLE_BLACK_CONSOLE',
        'label': '啟用黑色 Console 視窗',
        'help': '額外顯示一個即時輸出視窗。',
        'type': 'bool',
    },
    {
        'key': 'CONSOLE_POPUP_ON_COMPARISON',
        'label': '偵測到比較時彈出視窗',
        'help': '有比較輸出時自動帶到前景。',
        'type': 'bool',
    },
    {
        'key': 'CONSOLE_ALWAYS_ON_TOP',
        'label': '視窗保持最上層',
        'help': '讓 Console 視窗始終置頂。',
        'type': 'bool',
    },
    {
        'key': 'CONSOLE_TEMP_TOPMOST_DURATION',
        'label': '臨時置頂秒數',
        'help': '收到比較輸出時，視窗臨時置頂的時間。',
        'type': 'int',
    },
    {
        'key': 'CONSOLE_INITIAL_TOPMOST_DURATION',
        'label': '啟動初期置頂秒數',
        'help': '啟動後短暫置頂以避免被其他視窗遮住。',
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
        self.protocol('WM_DELETE_WINDOW', self._on_close)

        # Load defaults (config + runtime overrides)
        runtime_data = load_runtime_settings()

        frm = ttk.Frame(self)
        frm.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(frm)
        # Enable mouse-wheel scroll on Windows
        def _on_mousewheel(event):
            # Standardize Windows wheel delta
            delta = int(-1 * (event.delta / 120))
            canvas.yview_scroll(delta, 'units')
            return 'break'
        canvas.bind_all('<MouseWheel>', _on_mousewheel)

        scrollbar = ttk.Scrollbar(frm, orient='vertical', command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        scroll_frame.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=scroll_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Helper: treat empty strings/lists/dicts as blank, but keep 0/False
        def _is_blank(v):
            if v is None:
                return True
            if isinstance(v, str) and v.strip() == '':
                return True
            if isinstance(v, (list, tuple, dict)) and len(v) == 0:
                return True
            return False

        specs = sorted(PARAMS_SPEC, key=lambda s: s.get('priority', 100))
        for spec in specs:
            row = ttk.Frame(scroll_frame)
            row.pack(fill='x', padx=10, pady=6)
            ttk.Label(row, text=spec['label']).pack(anchor='w')
            help_lbl = ttk.Label(row, text=spec['help'], foreground='#666', wraplength=820, justify='left')
            help_lbl.pack(anchor='w')
            key = spec['key']
            cur_val = getattr(settings, key, '')
            if key in runtime_data and not _is_blank(runtime_data[key]):
                cur_val = runtime_data[key]
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
            elif spec['type'] == 'watch_subselect':
                # Checkbox list built from WATCH_FOLDERS (subset selection)
                frame = ttk.Frame(row)
                frame.pack(anchor='w', fill='x')
                vars_list = []
                rt = runtime.load_runtime_settings()
                watch_list = rt.get('WATCH_FOLDERS') if rt.get('WATCH_FOLDERS') else getattr(settings, 'WATCH_FOLDERS', [])
                cur_selected = set(cur_val or [])
                if not cur_selected:
                    cur_selected = set(watch_list)
                for path in watch_list:
                    var = tk.BooleanVar(value=(path in cur_selected))
                    cb = ttk.Checkbutton(frame, text=os.path.normpath(path), variable=var)
                    cb.pack(anchor='w')
                    vars_list.append((path, var))
                def sync_from_watch():
                    nonlocal vars_list
                    for child in frame.winfo_children():
                        child.destroy()
                    vars_list = []
                    rt2 = runtime.load_runtime_settings()
                    watch_list2 = rt2.get('WATCH_FOLDERS') if rt2.get('WATCH_FOLDERS') else getattr(settings, 'WATCH_FOLDERS', [])
                    for path in watch_list2:
                        var = tk.BooleanVar(value=True)
                        cb = ttk.Checkbutton(frame, text=os.path.normpath(path), variable=var)
                        cb.pack(anchor='w')
                        vars_list.append((path, var))
                ttk.Button(row, text='從監控資料夾同步', command=sync_from_watch).pack(anchor='w', pady=4)
                frame.vars_list = vars_list
                w = frame
            elif spec['type'] == 'paths':
                # Multi-select paths: provide listbox + add/remove buttons
                frame = ttk.Frame(row)
                frame.pack(anchor='w', fill='x')
                listbox = tk.Listbox(frame, height=5, width=80, selectmode=tk.EXTENDED)
                listbox.pack(side='left', fill='both', expand=True)
                for v in (cur_val or []):
                    try:
                        listbox.insert('end', os.path.normpath(v))
                    except Exception:
                        listbox.insert('end', str(v))
                btns = ttk.Frame(frame)
                btns.pack(side='left', padx=6)
                def add_path(lb=listbox, sp=spec):
                    if sp.get('path_kind') == 'file':
                        p = filedialog.askopenfilename()
                    else:
                        p = filedialog.askdirectory()
                    if p:
                        # Normalize to Windows-style backslashes
                        p = os.path.normpath(p)
                        lb.insert('end', p)
                def remove_selected(lb=listbox):
                    sel = list(lb.curselection())
                    sel.reverse()
                    for idx in sel:
                        lb.delete(idx)
                def clear_all(lb=listbox):
                    lb.delete(0, 'end')
                ttk.Button(btns, text='新增', command=add_path).pack(fill='x')
                ttk.Button(btns, text='移除選取', command=remove_selected).pack(fill='x', pady=2)
                ttk.Button(btns, text='全部清除', command=clear_all).pack(fill='x')
                w = listbox
            elif spec['type'] == 'path':
                # Single path with browse button
                frame = ttk.Frame(row)
                frame.pack(anchor='w', fill='x')
                var = tk.StringVar(value=str(cur_val))
                entry = ttk.Entry(frame, textvariable=var, width=70)
                entry.pack(side='left', fill='x', expand=True)
                def browse():
                    kind = spec.get('path_kind')
                    if kind == 'file':
                        p = filedialog.askopenfilename()
                    elif kind == 'save_file':
                        p = filedialog.asksaveasfilename()
                    else:
                        p = filedialog.askdirectory()
                    if p:
                        var.set(os.path.normpath(p))
                ttk.Button(frame, text='瀏覽...', command=browse).pack(side='left', padx=6)
                w = entry
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

        # 保證空白欄位會自動填入 settings.py 的預設值
        self._ensure_defaults_filled()
        # 若啟用自動同步，保持 SCAN_TARGET_FOLDERS 與 WATCH_FOLDERS 一致
        try:
            rt0 = runtime.load_runtime_settings()
            auto = rt0.get('AUTO_SYNC_SCAN_TARGETS', False)
        except Exception:
            auto = False
        if auto:
            spec_targets, widget_targets = self._widgets.get('SCAN_TARGET_FOLDERS', (None, None))
            spec_watch, widget_watch = self._widgets.get('WATCH_FOLDERS', (None, None))
            if widget_targets and widget_watch and spec_targets['type'] == 'paths' and spec_watch['type'] == 'paths':
                # 清空並以 WATCH_FOLDERS 內容填入
                widget_targets.delete(0, 'end')
                for v in list(widget_watch.get(0, 'end')):
                    widget_targets.insert('end', v)

        btn_row = ttk.Frame(self)
        btn_row.pack(fill='x', padx=10, pady=10)
        ttk.Button(btn_row, text='載入預設', command=self._reset_defaults).pack(side='left')
        ttk.Button(btn_row, text='載入設定檔...', command=self._load_preset).pack(side='left', padx=6)
        ttk.Button(btn_row, text='匯出設定檔...', command=self._save_preset).pack(side='left')
        ttk.Button(btn_row, text='儲存並開始', command=self._save_and_apply).pack(side='right')

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
            elif spec['type'] == 'paths':
                widget.delete(0, 'end')
                for v in (val or []):
                    try:
                        widget.insert('end', os.path.normpath(v))
                    except Exception:
                        widget.insert('end', str(v))
            elif spec['type'] == 'bool':
                widget.var.set(bool(val))
            elif spec['type'] == 'int':
                widget.delete(0, 'end')
                widget.insert(0, str(val))
            elif spec['type'] == 'choice':
                widget.set(str(val))

    def _ensure_defaults_filled(self):
        # 將畫面上仍為空白的欄位填入 settings.py 的預設值
        for key, (spec, widget) in self._widgets.items():
            val = getattr(settings, key, '')
            if spec['type'] == 'text' and widget.get().strip() == '':
                widget.insert(0, str(val))
            elif spec['type'] == 'multiline':
                raw = widget.get('1.0', 'end').strip()
                if raw == '':
                    if isinstance(val, (list, tuple)):
                        widget.insert('1.0', '\n'.join(val))
                    else:
                        widget.insert('1.0', str(val))
            elif spec['type'] == 'paths':
                if widget.size() == 0:
                    for v in (val or []):
                        try:
                            widget.insert('end', os.path.normpath(v))
                        except Exception:
                            widget.insert('end', str(v))
            elif spec['type'] == 'path':
                if widget.get().strip() == '' and val:
                    try:
                        widget.delete(0, 'end')
                        widget.insert(0, os.path.normpath(val))
                    except Exception:
                        widget.insert(0, str(val))
            elif spec['type'] == 'bool':
                # 不覆蓋既有勾選狀態
                pass
            elif spec['type'] == 'int' and widget.get().strip() == '':
                widget.insert(0, str(val))
            elif spec['type'] == 'choice' and not widget.get().strip():
                widget.set(str(val))

    def _collect_values(self) -> Dict[str, Any]:
        data: Dict[str, Any] = {}
        for key, (spec, widget) in self._widgets.items():
            if spec['type'] in ('text','int','choice'):
                data[key] = widget.get().strip()
            elif spec['type'] == 'path':
                val = widget.get().strip()
                if val:
                    try:
                        import os
                        val = os.path.normpath(val)
                    except Exception:
                        pass
                data[key] = val
            elif spec['type'] == 'multiline':
                raw = widget.get('1.0', 'end').strip()
                lines = [l.strip() for l in raw.split('\n') if l.strip()]
                data[key] = lines
            elif spec['type'] == 'paths':
                items = list(widget.get(0, 'end'))
                data[key] = items
            elif spec['type'] == 'watch_subselect':
                # collect checked subset
                items = []
                for path, var in getattr(widget, 'vars_list', []):
                    if bool(var.get()):
                        items.append(path)
                data[key] = items
            elif spec['type'] == 'bool':
                data[key] = bool(widget.var.get())
        # normalize SUPPORTED_EXTS string to tuple-like list
        exts = data.get('SUPPORTED_EXTS')
        if isinstance(exts, str):
            items = [x.strip() for x in exts.replace(';', ',').split(',') if x.strip()]
            norm = []
            for x in items:
                x = x.strip(" ' \"()[]{}").lower()
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
            # 按下儲存並開始：確保移除取消旗標
            if 'STARTUP_CANCELLED' in data:
                data.pop('STARTUP_CANCELLED', None)
            # persist and apply
            save_runtime_settings(data)
            apply_to_settings(data)
            self.destroy()
        except Exception as e:
            messagebox.showerror('錯誤', f'儲存設定失敗: {e}')

    def _on_close(self):
        # 使用者按視窗右上角關閉：視為取消，停止 watchdog 啟動
        try:
            # 傳回一個特殊旗標到 runtime json，讓 main 判斷不要繼續執行
            data = runtime.load_runtime_settings() or {}
            data['STARTUP_CANCELLED'] = True
            save_runtime_settings(data)
        except Exception:
            pass
        self.destroy()

    def _save_preset(self):
        try:
            data = self._collect_values()
            path = filedialog.asksaveasfilename(defaultextension='.json', filetypes=[('JSON Files','*.json')])
            if not path:
                return
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo('完成', '已儲存為範本')
        except Exception as e:
            messagebox.showerror('錯誤', f'儲存範本失敗: {e}')

    def _load_preset(self):
        try:
            path = filedialog.askopenfilename(filetypes=[('JSON Files','*.json')])
            if not path:
                return
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            # 將值套到畫面（不直接 apply 到 settings）
            for key, (spec, widget) in self._widgets.items():
                if key not in data:
                    continue
                val = data[key]
                if spec['type'] == 'text':
                    widget.delete(0, 'end')
                    widget.insert(0, str(val))
                elif spec['type'] == 'int':
                    widget.delete(0, 'end')
                    widget.insert(0, str(val))
                elif spec['type'] == 'choice':
                    widget.set(str(val))
                elif spec['type'] == 'multiline':
                    widget.delete('1.0', 'end')
                    if isinstance(val, (list, tuple)):
                        widget.insert('1.0', '\n'.join(val))
                    else:
                        widget.insert('1.0', str(val))
                elif spec['type'] == 'paths':
                    widget.delete(0, 'end')
                    for v in val or []:
                        widget.insert('end', v)
                elif spec['type'] == 'bool':
                    widget.var.set(bool(val))
            messagebox.showinfo('完成', '已載入範本內容（尚未套用，請按「儲存並開始」）。')
        except Exception as e:
            messagebox.showerror('錯誤', f'載入範本失敗: {e}')


def show_settings_ui():
    root = tk.Tk()
    root.withdraw()
    dlg = SettingsDialog(root)
    # Keep window visible and interactive; don't block main thread longer than necessary
    root.wait_window(dlg)
    root.destroy()

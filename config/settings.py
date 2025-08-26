"""
系統配置設定
所有原始配置都在這裡，確保向後相容
"""
import os
from datetime import datetime

# =========== User Config ============
TRACK_EXTERNAL_REFERENCES = True       # 追蹤外部參照更新
TRACK_DIRECT_VALUE_CHANGES = True      # 追蹤直接值變更
TRACK_FORMULA_CHANGES = True           # 追蹤公式變更
IGNORE_INDIRECT_CHANGES = True         # 忽略間接影響
ENABLE_BLACK_CONSOLE = True
CONSOLE_POPUP_ON_COMPARISON = True
CONSOLE_ALWAYS_ON_TOP = False           # 新增：是否始終置頂
CONSOLE_TEMP_TOPMOST_DURATION = 5       # 新增：臨時置頂持續時間（秒）
CONSOLE_INITIAL_TOPMOST_DURATION = 2    # 新增：初始置頂持續時間（秒）
SHOW_COMPRESSION_STATS = False          # 關閉壓縮統計顯示
SHOW_DEBUG_MESSAGES = False             # 關閉調試訊息
AUTO_UPDATE_BASELINE_AFTER_COMPARE = True  # 比較後自動更新基準線
SCAN_ALL_MODE = True
MAX_CHANGES_TO_DISPLAY = 20 # 限制顯示的變更數量，0 表示不限制
USE_LOCAL_CACHE = True
CACHE_FOLDER = r"C:\Users\user\Desktop\watchdog\cache_folder"
ENABLE_FAST_MODE = True
ENABLE_TIMEOUT = True
FILE_TIMEOUT_SECONDS = 120
ENABLE_MEMORY_MONITOR = True
MEMORY_LIMIT_MB = 2048
ENABLE_RESUME = True
FORMULA_ONLY_MODE = True
DEBOUNCE_INTERVAL_SEC = 2

# =========== Compression Config ============
# 預設壓縮格式：'lz4' 用於頻繁讀寫, 'zstd' 用於長期存儲, 'gzip' 用於兼容性
DEFAULT_COMPRESSION_FORMAT = 'lz4'  # 'lz4', 'zstd', 'gzip'

# 壓縮級別設定
LZ4_COMPRESSION_LEVEL = 1       # LZ4: 0-16, 越高壓縮率越好但越慢
ZSTD_COMPRESSION_LEVEL = 3      # Zstd: 1-22, 推薦 3-6
GZIP_COMPRESSION_LEVEL = 6      # gzip: 1-9, 推薦 6

# 歸檔設定
ENABLE_ARCHIVE_MODE = True              # 是否啟用歸檔模式
ARCHIVE_AFTER_DAYS = 7                  # 多少天後轉為歸檔格式
ARCHIVE_COMPRESSION_FORMAT = 'zstd'     # 歸檔使用的壓縮格式

# 效能監控
SHOW_COMPRESSION_STATS = True           # 是否顯示壓縮統計

RESUME_LOG_FILE = r"C:\Users\user\Desktop\watchdog\resume_log\baseline_progress.log"
WATCH_FOLDERS = [
    r"C:\Users\user\Desktop\Test",
]
MANUAL_BASELINE_TARGET = []
LOG_FOLDER = r"C:\Users\user\Desktop\watchdog\log_folder"
LOG_FILE_DATE = datetime.now().strftime('%Y%m%d')
CSV_LOG_FILE = os.path.join(LOG_FOLDER, f"excel_change_log_{LOG_FILE_DATE}.csv.gz")
SUPPORTED_EXTS = ('.xlsx', '.xlsm')
MAX_RETRY = 10
RETRY_INTERVAL_SEC = 2
USE_TEMP_COPY = True
WHITELIST_USERS = ['ckcm0210', 'yourwhiteuser']
LOG_WHITELIST_USER_CHANGE = True
FORCE_BASELINE_ON_FIRST_SEEN = [
    r"\\network_drive\\your_folder1\\must_first_baseline.xlsx",
    "force_this_file.xlsx"
]

# =========== Polling Config ============
POLLING_SIZE_THRESHOLD_MB = 10
DENSE_POLLING_INTERVAL_SEC = 10
DENSE_POLLING_DURATION_SEC = 15
SPARSE_POLLING_INTERVAL_SEC = 15
SPARSE_POLLING_DURATION_SEC = 15

# =========== 全局變數 ============
current_processing_file = None
processing_start_time = None
force_stop = False
baseline_completed = False
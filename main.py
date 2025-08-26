"""
Excel Monitor 主執行檔案
這是唯一需要執行的檔案
"""
import os
import sys
import signal
import threading
import time
from datetime import datetime
import logging

# 確保能夠導入模組
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

# 導入各個模組
import config.settings as settings
from utils.logging import init_logging
from utils.memory import check_memory_limit
from utils.helpers import get_all_excel_files, timeout_handler
from utils.compression import CompressionFormat, test_compression_support  # 新增
from ui.console import init_console
from core.baseline import create_baseline_for_files_robust
from core.watcher import active_polling_handler, ExcelFileEventHandler
from core.comparison import set_current_event_number
from watchdog.observers import Observer

def signal_handler(signum, frame):
    """
    信號處理器，優雅地停止程序
    """
    if not settings.force_stop:
        settings.force_stop = True
        print("\n🛑 收到中斷信號，正在安全停止...")
        if settings.current_processing_file: 
            print(f"   目前處理檔案: {settings.current_processing_file}")
        active_polling_handler.stop()
        print("   (再按一次 Ctrl+C 強制退出)")
    else:
        print("\n💥 強制退出...")
        sys.exit(1)

def main():
    """
    主函數
    """
    print("🚀 Excel Monitor v2.1 啟動中...")
    
    # 測試壓縮支援
    test_compression_support()
    
    # 初始化日誌系統
    init_logging()
    
    # 啟動前設定 UI（可讓使用者覆寫 settings）
    try:
        from ui.settings_ui import show_settings_ui
        show_settings_ui()
        # 若使用者關閉設定視窗（取消啟動），不要繼續運行
        from config.runtime import load_runtime_settings
        if (load_runtime_settings() or {}).get('STARTUP_CANCELLED'):
            print('使用者取消啟動，退出程式。')
            return
    except Exception as e:
        print(f"⚠️ 設定 UI 啟動失敗，使用預設設定: {e}")
    
    # 初始化控制台
    console = init_console()
    
    # 設定信號處理器
    signal.signal(signal.SIGINT, signal_handler)
    
    # 啟動超時監控
    if settings.ENABLE_TIMEOUT:
        timeout_thread = threading.Thread(target=timeout_handler, daemon=True)
        timeout_thread.start()
    
    # 檢查壓縮格式支援
    available_formats = CompressionFormat.get_available_formats()
    print(f"🗜️  支援壓縮格式: {', '.join(available_formats)}")
    validated_format = CompressionFormat.validate_format(settings.DEFAULT_COMPRESSION_FORMAT)
    if validated_format != settings.DEFAULT_COMPRESSION_FORMAT:
        print(f"⚠️  格式已調整: {settings.DEFAULT_COMPRESSION_FORMAT} → {validated_format}")
        settings.DEFAULT_COMPRESSION_FORMAT = validated_format
    
    print(f"📁 監控資料夾: {settings.WATCH_FOLDERS}")
    if getattr(settings, 'MONITOR_ONLY_FOLDERS', None):
        print(f"🛈  只監控變更的根目錄: {settings.MONITOR_ONLY_FOLDERS}")
    print(f"📊 支援格式: {settings.SUPPORTED_EXTS}")
    print(f"⚙️  設定檔案: 已載入")
    
    # 🔥 處理手動基準線目標
    manual_files = []
    if settings.MANUAL_BASELINE_TARGET:
        print(f"📋 手動基準線目標: {len(settings.MANUAL_BASELINE_TARGET)} 個")
        for target in settings.MANUAL_BASELINE_TARGET:
            if os.path.exists(target):
                manual_files.append(target)
                print(f"   ✅ {os.path.basename(target)}")
            else:
                print(f"   ❌ 檔案不存在: {target}")
    
    # 獲取所有 Excel 檔案
    all_files = []
    if settings.SCAN_ALL_MODE:
        print("\n🔍 掃描所有 Excel 檔案...")
        scan_roots = list(settings.WATCH_FOLDERS or [])
        # 若使用者指定 SCAN_TARGET_FOLDERS，僅針對該子集掃描
        if getattr(settings, 'SCAN_TARGET_FOLDERS', None):
            scan_roots = list(dict.fromkeys([r for r in settings.SCAN_TARGET_FOLDERS if r]))
        all_files = get_all_excel_files(scan_roots)
        print(f"找到 {len(all_files)} 個 Excel 檔案")
    
    # 🔥 合併手動目標和掃描結果
    total_files = list(set(all_files + manual_files))
    
    # 建立基準線
    if total_files:
        print(f"\n📊 總共需要處理 {len(total_files)} 個檔案")
        create_baseline_for_files_robust(total_files)
    
    # 啟動檔案監控
    print("\n👀 啟動檔案監控...")
    event_handler = ExcelFileEventHandler(active_polling_handler)
    observer = Observer()
    
    # 對 WATCH_FOLDERS 與 MONITOR_ONLY_FOLDERS 都要註冊監控
    watch_roots = list(dict.fromkeys(list(settings.WATCH_FOLDERS or []) + list(getattr(settings, 'MONITOR_ONLY_FOLDERS', []) or [])))
    if not watch_roots:
        print("   ⚠️  沒有任何監控根目錄（WATCH_FOLDERS 或 MONITOR_ONLY_FOLDERS 為空）")
    for folder in watch_roots:
        if os.path.exists(folder):
            observer.schedule(event_handler, folder, recursive=True)
            print(f"   監控: {folder}")
        else:
            print(f"   ⚠️  資料夾不存在: {folder}")
    
    observer.start()
    
    print("\n✅ Excel Monitor 已啟動完成！")
    print("🎯 功能狀態:")
    print(f"   - 公式模式: {'開啟' if settings.FORMULA_ONLY_MODE else '關閉'}")
    print(f"   - 白名單過濾: {'開啟' if settings.WHITELIST_USERS else '關閉'}")
    print(f"   - 本地緩存: {'開啟' if settings.USE_LOCAL_CACHE else '關閉'}")
    print(f"   - 黑色控制台: {'開啟' if settings.ENABLE_BLACK_CONSOLE else '關閉'}")
    print(f"   - 記憶體監控: {'開啟' if settings.ENABLE_MEMORY_MONITOR else '關閉'}")
    print(f"   - 壓縮格式: {settings.DEFAULT_COMPRESSION_FORMAT.upper()}")
    print(f"   - 歸檔模式: {'開啟' if settings.ENABLE_ARCHIVE_MODE else '關閉'}")
    print("\n按 Ctrl+C 停止監控...")
    
    try:
        while not settings.force_stop:
            time.sleep(1)
    except KeyboardInterrupt:
        pass
    finally:
        print("\n🔄 正在停止監控...")
        observer.stop()
        observer.join()
        active_polling_handler.stop()
        print("✅ 監控已停止")

if __name__ == "__main__":
    main()
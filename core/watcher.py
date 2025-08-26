import os
import time
import threading
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import config.settings as settings
import logging
from datetime import datetime

class ActivePollingHandler:
    """
    主動輪詢處理器，採用新的智慧輪詢邏輯
    """
    def __init__(self):
        self.polling_tasks = {}
        self.lock = threading.Lock()
        self.stop_event = threading.Event()

    def start_polling(self, file_path, event_number):
        """
        根據檔案大小決定輪詢策略（用 mtime 穩定檢查，不再用與 baseline 的差異判斷）
        """
        try:
            file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
        except (FileNotFoundError, PermissionError, OSError) as e:
            logging.warning(f"獲取檔案大小失敗: {file_path}, 錯誤: {e}")
            file_size_mb = 0

        interval = settings.DENSE_POLLING_INTERVAL_SEC if file_size_mb < settings.POLLING_SIZE_THRESHOLD_MB else settings.SPARSE_POLLING_INTERVAL_SEC
        polling_type = "密集" if file_size_mb < settings.POLLING_SIZE_THRESHOLD_MB else "稀疏"
        
        print(f"[輪詢] 檔案: {os.path.basename(file_path)} ({polling_type}輪詢，每 {interval}s 檢查一次)")
        # 初始化 last_mtime
        try:
            last_mtime = os.path.getmtime(file_path)
        except Exception:
            last_mtime = 0
        self._start_adaptive_polling(file_path, event_number, interval, last_mtime)

    def _start_adaptive_polling(self, file_path, event_number, interval, last_mtime):
        """
        開始自適應輪詢
        """
        with self.lock:
            if file_path in self.polling_tasks:
                self.polling_tasks[file_path]['timer'].cancel()

            def task_wrapper():
                self._poll_for_stability(file_path, event_number, interval, last_mtime)

            timer = threading.Timer(interval, task_wrapper)
            self.polling_tasks[file_path] = {'timer': timer}
            timer.start()
            print(f"    [輪詢啟動] {interval} 秒後首次檢查 {os.path.basename(file_path)}")

    def _poll_for_stability(self, file_path, event_number, interval, last_mtime):
        """
        執行輪詢檢查，如果檔案變更則延長輪詢，否則結束
        """
        if self.stop_event.is_set():
            return

        print(f"    [輪詢檢查] 正在檢查 {os.path.basename(file_path)} 的變更...")

        # 以 mtime 穩定判斷：如果 mtime 沒再變，視為穩定
        try:
            current_mtime = os.path.getmtime(file_path)
        except Exception:
            current_mtime = last_mtime

        has_changes = False
        if current_mtime != last_mtime:
            # 有新變更，更新 last_mtime 並做一次比較
            last_mtime = current_mtime
            from core.comparison import compare_excel_changes, set_current_event_number
            set_current_event_number(event_number)
            has_changes = compare_excel_changes(file_path, silent=False, event_number=event_number, is_polling=True)

        with self.lock:
            if file_path not in self.polling_tasks:
                return

            if has_changes:
                print(f"    [輪詢] 檔案仍在變更，延長等待時間，{interval} 秒後再次檢查。")
                
                def task_wrapper():
                    self._poll_for_stability(file_path, event_number, interval, last_mtime)
                
                new_timer = threading.Timer(interval, task_wrapper)
                self.polling_tasks[file_path]['timer'] = new_timer
                new_timer.start()
            else:
                print(f"    [輪詢結束] {os.path.basename(file_path)} 檔案已穩定。")
                self.polling_tasks.pop(file_path, None)

    def stop(self):
        """
        停止所有輪詢任務
        """
        self.stop_event.set()
        with self.lock:
            for task in self.polling_tasks.values():
                task['timer'].cancel()
            self.polling_tasks.clear()

class ExcelFileEventHandler(FileSystemEventHandler):
    """
    Excel 檔案事件處理器
    """
    def __init__(self, polling_handler):
        self.polling_handler = polling_handler
        self.last_event_times = {}
        self.event_counter = 0
        
    def on_created(self, event):
        """
        檔案建立事件處理
        """
        if event.is_directory:
            return

        file_path = event.src_path

        # 檢查是否為支援的 Excel 檔案
        if not file_path.lower().endswith(settings.SUPPORTED_EXTS):
            return

        # 檢查是否為臨時檔案
        if os.path.basename(file_path).startswith('~$'):
            return

        print(f"\n✨ 發現新檔案: {os.path.basename(file_path)}")
        print(f"📊 正在建立基準線...")

        from core.baseline import create_baseline_for_files_robust
        create_baseline_for_files_robust([file_path])

        print(f"✅ 基準線建立完成，已納入監控: {os.path.basename(file_path)}")

    def _is_in_watch_folders(self, path: str) -> bool:
        try:
            p = os.path.abspath(path)
            for root in (settings.WATCH_FOLDERS or []):
                r = os.path.abspath(root)
                if os.path.commonpath([p, r]) == r:
                    # 排除清單
                    for ex in (getattr(settings, 'WATCH_EXCLUDE_FOLDERS', []) or []):
                        exa = os.path.abspath(ex)
                        if os.path.commonpath([p, exa]) == exa:
                            return False
                    return True
        except Exception:
            pass
        return False

    def _is_monitor_only(self, path: str) -> bool:
        # WATCH_FOLDERS 優先於 MONITOR_ONLY_FOLDERS
        if self._is_in_watch_folders(path):
            return False
        try:
            p = os.path.abspath(path)
            for root in (settings.MONITOR_ONLY_FOLDERS or []):
                r = os.path.abspath(root)
                if os.path.commonpath([p, r]) == r:
                    # 排除清單
                    for ex in (getattr(settings, 'MONITOR_ONLY_EXCLUDE_FOLDERS', []) or []):
                        exa = os.path.abspath(ex)
                        if os.path.commonpath([p, exa]) == exa:
                            return False
                    return True
        except Exception:
            pass
        return False

    def on_modified(self, event):
        """
        檔案修改事件處理
        """
        if event.is_directory:
            return
            
        file_path = event.src_path
        
        # 檢查是否為支援的 Excel 檔案
        if not file_path.lower().endswith(settings.SUPPORTED_EXTS):
            return
            
        # 檢查是否為臨時檔案
        if os.path.basename(file_path).startswith('~$'):
            return
            
        # 防抖動處理
        current_time = time.time()
        if file_path in self.last_event_times:
            if current_time - self.last_event_times[file_path] < settings.DEBOUNCE_INTERVAL_SEC:
                return
                
        self.last_event_times[file_path] = current_time
        self.event_counter += 1
        
        # 獲取檔案最後作者
        try:
            from core.excel_parser import get_excel_last_author
            last_author = get_excel_last_author(file_path)
            author_info = f" (最後儲存者: {last_author})" if last_author != 'Unknown' else ""
        except Exception as e:
            author_info = ""
        
        # 先做一次靜默比對，若無變更則不噪音輸出（仍可後續輪詢）
        from core.comparison import compare_excel_changes, set_current_event_number
        set_current_event_number(self.event_counter)
        has_changes_preview = compare_excel_changes(file_path, silent=True, event_number=self.event_counter, is_polling=False)
        
        if has_changes_preview:
            print(f"\n🔔 檔案變更偵測: {os.path.basename(file_path)} (事件 #{self.event_counter}){author_info}")
        
        # 監控但不預先 baseline 的區域：首次變更只紀錄資訊並建立 baseline，之後才比較
        if self._is_monitor_only(file_path):
            try:
                from utils.helpers import get_file_mtime
                mtime = get_file_mtime(file_path)
                print(f"    [MONITOR-ONLY] {file_path}\n       - 最後修改時間: {mtime}\n       - 最後儲存者: {last_author}")
                # 若尚未有 baseline，先建立一份；已存在則繼續走下面的比較流程
                from core.baseline import get_baseline_file_with_extension, save_baseline
                from core.excel_parser import dump_excel_cells_with_timeout, hash_excel_content
                from utils.helpers import _baseline_key_for_path
                base_key = _baseline_key_for_path(file_path)
                baseline_exists = bool(get_baseline_file_with_extension(base_key))
                if not baseline_exists:
                    cur = dump_excel_cells_with_timeout(file_path)
                    if cur:
                        bdata = {"last_author": last_author, "content_hash": hash_excel_content(cur), "cells": cur, "timestamp": datetime.now().isoformat()}
                        save_baseline(base_key, bdata)
                        print("    [MONITOR-ONLY] 已建立首次基準線（本次不比較）。")
                        return
            except Exception as e:
                logging.warning(f"monitor-only 初始化失敗: {e}")
                return
        
        # 🔥 設定事件編號並立即執行一次比較
        from core.comparison import compare_excel_changes, set_current_event_number
        set_current_event_number(self.event_counter)
        
        # 檢查檔案是否已經在輪詢中
        if file_path in self.polling_handler.polling_tasks:
            print(f"    [偵測] {os.path.basename(file_path)} 正在輪詢中，忽略本次即時檢查。")
            return

        print(f"📊 立即檢查變更...")
        has_changes = compare_excel_changes(file_path, silent=False, event_number=self.event_counter, is_polling=False)
        
        if has_changes:
            print(f"✅ 偵測到變更，啟動輪詢以監控後續活動...")
        else:
            print(f"ℹ️  未發現即時變更，啟動輪詢以監控後續活動...")
        
        # 開始輪詢
        self.polling_handler.start_polling(file_path, self.event_counter)

# 創建全局輪詢處理器實例
active_polling_handler = ActivePollingHandler()

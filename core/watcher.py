import os
import time
import threading
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import config.settings as settings
import logging

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
        根據檔案大小決定輪詢策略
        """
        try:
            file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
        except (FileNotFoundError, PermissionError, OSError) as e:
            logging.warning(f"獲取檔案大小失敗: {file_path}, 錯誤: {e}")
            file_size_mb = 0

        interval = settings.DENSE_POLLING_INTERVAL_SEC if file_size_mb < settings.POLLING_SIZE_THRESHOLD_MB else settings.SPARSE_POLLING_INTERVAL_SEC
        polling_type = "密集" if file_size_mb < settings.POLLING_SIZE_THRESHOLD_MB else "稀疏"
        
        print(f"[輪詢] 檔案: {os.path.basename(file_path)} ({polling_type}輪詢，每 {interval}s 檢查一次)")
        self._start_adaptive_polling(file_path, event_number, interval)

    def _start_adaptive_polling(self, file_path, event_number, interval):
        """
        開始自適應輪詢
        """
        with self.lock:
            if file_path in self.polling_tasks:
                self.polling_tasks[file_path]['timer'].cancel()

            def task_wrapper():
                self._poll_for_stability(file_path, event_number, interval)

            timer = threading.Timer(interval, task_wrapper)
            self.polling_tasks[file_path] = {'timer': timer}
            timer.start()
            print(f"    [輪詢啟動] {interval} 秒後首次檢查 {os.path.basename(file_path)}")

    def _poll_for_stability(self, file_path, event_number, interval):
        """
        執行輪詢檢查，如果檔案變更則延長輪詢，否則結束
        """
        if self.stop_event.is_set():
            return

        print(f"    [輪詢檢查] 正在檢查 {os.path.basename(file_path)} 的變更...")

        from core.comparison import compare_excel_changes, set_current_event_number
        set_current_event_number(event_number)
        has_changes = compare_excel_changes(file_path, silent=False, event_number=event_number, is_polling=True)

        with self.lock:
            if file_path not in self.polling_tasks:
                return

            if has_changes:
                print(f"    [輪詢] 檔案仍在變更，延長等待時間，{interval} 秒後再次檢查。")
                
                def task_wrapper():
                    self._poll_for_stability(file_path, event_number, interval)
                
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
        
        print(f"\n🔔 檔案變更偵測: {os.path.basename(file_path)} (事件 #{self.event_counter}){author_info}")
        
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

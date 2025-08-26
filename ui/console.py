"""
黑色控制台視窗
"""
import tkinter as tk
from tkinter import scrolledtext
import queue
import threading
import time
import config.settings as settings
import logging

class BlackConsoleWindow:
    def __init__(self):
        self.root = None
        self.text_widget = None
        self.message_queue = queue.Queue()
        self.running = False
        self.is_minimized = False
        self.popup_on_comparison = settings.CONSOLE_POPUP_ON_COMPARISON
        self.is_topmost = False  # 追蹤置頂狀態
        self.topmost_timer = None  # 🔥 新增：追蹤計時器
        
    def create_window(self):
        """創建黑色 console 視窗"""
        self.root = tk.Tk()
        self.root.title("Excel Watchdog Console")
        self.root.geometry("1200x1000")
        self.root.configure(bg='black')
        
        # 🔥 程式啟動時彈出並短暫置頂
        self.root.attributes('-topmost', True)
        self.root.lift()
        self.root.focus_force()
        self.is_topmost = True
        
        # 3秒後取消置頂
        self.schedule_disable_topmost(3000)
        
        # 監控視窗狀態變化
        self.root.bind('<Unmap>', self.on_minimize)
        self.root.bind('<Map>', self.on_restore)
        
        # 創建滾動文字區域
        self.text_widget = scrolledtext.ScrolledText(
            self.root,
            bg='black',
            fg='white',
            font=('Consolas', 10),
            insertbackground='white',
            selectbackground='darkgray',
            wrap=tk.WORD
        )
        self.text_widget.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 設定視窗關閉事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.running = True
        self.check_messages()
        
    def schedule_disable_topmost(self, delay_ms):
        """安排取消置頂 - 避免重複計時器"""
        # 🔥 取消之前的計時器
        if self.topmost_timer:
            self.root.after_cancel(self.topmost_timer)
        
        # 設定新的計時器
        self.topmost_timer = self.root.after(delay_ms, self.disable_topmost)
        
    def disable_topmost(self):
        """取消置頂狀態"""
        if self.root and self.running and self.is_topmost:
            self.root.attributes('-topmost', False)
            self.is_topmost = False
            self.topmost_timer = None  # 清除計時器引用
        
    def on_minimize(self, event):
        """當視窗被最小化時觸發"""
        self.is_minimized = True
        
    def on_restore(self, event):
        """當視窗被恢復時觸發"""
        self.is_minimized = False
        
    def popup_window(self):
        """彈出視窗到最上層 - 有新訊息時"""
        if self.root and self.running:
            try:
                self.root.deiconify()  # 將視窗從最小化還原
                
                # 🔥 關鍵修正：短暫置頂以強制彈出，然後立即取消
                self.root.attributes('-topmost', True)
                self.root.lift()
                self.root.focus_force()
                # 使用 after 來確保置頂先生效，再取消
                self.root.after(100, lambda: self.root.attributes('-topmost', False))

                # 短暫閃爍效果來吸引注意
                def flash_window():
                    original_bg = self.root.cget('bg')
                    self.root.configure(bg='darkred')
                    self.root.after(200, lambda: self.root.configure(bg=original_bg))
                
                flash_window()

                self.is_minimized = False
                
            except Exception as e:
                logging.error(f"彈出視窗失敗: {e}")
                    
    def check_messages(self):
        """檢查並顯示新訊息"""
        try:
            has_new_messages = False
            message_count = 0
            
            while not self.message_queue.empty():
                message_data = self.message_queue.get_nowait()
                has_new_messages = True
                message_count += 1
                
                # 判斷是普通訊息還是特殊訊息
                if isinstance(message_data, dict):
                    message = message_data.get('message', '')
                    is_comparison = message_data.get('is_comparison', False)
                    
                    # 如果是比較表格，彈出視窗
                    if is_comparison and self.popup_on_comparison:
                        self.popup_window()
                else:
                    # 向後兼容：如果是普通字串
                    message = str(message_data)
                
                self.text_widget.insert(tk.END, message + '\n')
                self.text_widget.see(tk.END)
            
            # 如果有新訊息且視窗被最小化，彈出一次就好
            if has_new_messages and self.is_minimized:
                self.popup_window()
                # 🔥 移除這行 - 不需要顯示給用戶
                # print(f"[DEBUG] 收到 {message_count} 條新訊息，彈出視窗")
                
        except queue.Empty:
            pass
        
        if self.running:
            self.root.after(100, self.check_messages)
    
    def add_message(self, message, is_comparison=False):
        """添加訊息到佇列"""
        if self.running:
            message_data = {
                'message': message,
                'is_comparison': is_comparison
            }
            self.message_queue.put(message_data)
    
    def toggle_topmost(self):
        """手動切換置頂狀態"""
        if self.root and self.running:
            self.is_topmost = not self.is_topmost
            self.root.attributes('-topmost', self.is_topmost)
            
            # 如果手動關閉置頂，取消自動計時器
            if not self.is_topmost and self.topmost_timer:
                self.root.after_cancel(self.topmost_timer)
                self.topmost_timer = None
            
            status = "開啟" if self.is_topmost else "關閉"

    
    def on_closing(self):
        """關閉視窗時的處理"""
        self.running = False
        # 清理計時器
        if self.topmost_timer:
            self.root.after_cancel(self.topmost_timer)
        self.root.destroy()
    
    def start(self):
        """在新線程中啟動視窗"""
        def run_window():
            self.create_window()
            self.root.mainloop()
        
        window_thread = threading.Thread(target=run_window, daemon=True)
        window_thread.start()
        
        # 等待視窗創建完成
        while self.root is None:
            time.sleep(0.1)

# 全局 console 視窗實例
black_console = None

def init_console():
    """初始化控制台"""
    global black_console
    if settings.ENABLE_BLACK_CONSOLE:
        black_console = BlackConsoleWindow()
        black_console.start()
    return black_console
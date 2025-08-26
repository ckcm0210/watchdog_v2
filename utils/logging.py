"""
日誌和打印功能
"""
import builtins
from datetime import datetime
from io import StringIO
from wcwidth import wcswidth, wcwidth

# 保存原始 print 函數
_original_print = builtins.print

def timestamped_print(*args, **kwargs):
    """
    帶時間戳的打印函數
    """
    # 如果有 file=... 參數，直接用原生 print
    if 'file' in kwargs:
        _original_print(*args, **kwargs)
        return

    output_buffer = StringIO()
    _original_print(*args, file=output_buffer, **kwargs)
    message = output_buffer.getvalue()
    output_buffer.close()

    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    # 簡化邏輯：所有行都加時間戳記
    lines = message.rstrip().split('\n')
    timestamped_lines = []
    
    for line in lines:
        timestamped_lines.append(f"[{timestamp}] {line}")
    
    timestamped_message = '\n'.join(timestamped_lines)
    _original_print(timestamped_message)
    
    # 檢查是否為比較表格訊息
    is_comparison = any(keyword in message for keyword in [
        'Address', 'Baseline', 'Current', 
        '[SUMMARY]', '====', '----',
        '[MOD]', '[ADD]', '[DEL]'
    ])
    
    # 同時送到黑色 console - 使用延遲導入避免循環導入
    try:
        from ui.console import black_console
        if black_console and black_console.running:
            black_console.add_message(timestamped_message, is_comparison=is_comparison)
    except ImportError:
        pass

def init_logging():
    """
    初始化日誌系統
    """
    builtins.print = timestamped_print

def wrap_text_with_cjk_support(text, width):
    """
    自研的、支持 CJK 字符寬度的智能文本換行函數
    """
    lines = []
    line = ""
    current_width = 0
    for char in text:
        char_width = wcwidth(char)
        if char_width < 0: 
            continue # 跳過控制字符

        if current_width + char_width > width:
            lines.append(line)
            line = char
            current_width = char_width
        else:
            line += char
            current_width += char_width
    if line:
        lines.append(line)
    return lines or ['']

def _get_display_width(text):
    """
    精準計算一個字串的顯示闊度，處理 CJK 全形字元
    """
    return wcswidth(str(text))
"""
比較和差異顯示功能 - 確保 TABLE 一定顯示
"""
import os
import csv
import gzip
import json
import time
from datetime import datetime
from wcwidth import wcwidth
import config.settings as settings
from utils.logging import _get_display_width
from utils.helpers import get_file_mtime
from core.excel_parser import pretty_formula, extract_external_refs, get_excel_last_author
from core.baseline import load_baseline, baseline_file_path
import logging

# ... [print_aligned_console_diff 和其他輔助函數保持不變] ...
def print_aligned_console_diff(old_data, new_data, file_info=None, max_display_changes=0):
    """
    三欄式顯示，能處理中英文對齊，並正確顯示 formula。
    Address 欄固定闊度，Baseline/Current 平均分配。
    """
    try:
        term_width = os.get_terminal_size().columns
    except OSError:
        term_width = 120

    address_col_width = 12
    separators_width = 4
    remaining_width = term_width - address_col_width - separators_width
    baseline_col_width = remaining_width // 2
    current_col_width = remaining_width - baseline_col_width

    def wrap_text(text, width):
        lines = []
        current_line = ""
        current_width = 0
        for char in str(text):
            char_width = wcwidth(char)
            if char_width < 0:
                continue
            if current_width + char_width > width:
                lines.append(current_line)
                current_line = char
                current_width = char_width
            else:
                current_line += char
                current_width += char_width
        if current_line:
            lines.append(current_line)
        return lines or ['']

    def pad_line(line, width):
        line_width = _get_display_width(line)
        if line_width is None:
            line_width = len(str(line))
        padding = width - line_width
        return str(line) + ' ' * padding if padding > 0 else str(line)

    def format_cell(cell_value):
        if cell_value is None or cell_value == {}:
            return "(Empty)"
        if isinstance(cell_value, dict):
            formula = cell_value.get("formula")
            if formula:
                return f"={formula}"
            if "value" in cell_value:
                return repr(cell_value["value"])
        return repr(cell_value)
    
    print()
    print("=" * term_width)
    if file_info:
        filename = file_info.get('filename', 'Unknown')
        worksheet = file_info.get('worksheet', '')
        event_number = file_info.get('event_number')
        file_path = file_info.get('file_path', filename)

        event_str = f"(事件#{event_number}) " if event_number else ""
        caption = f"{event_str}{file_path} [Worksheet: {worksheet}]" if worksheet else f"{event_str}{file_path}"
        for cap_line in wrap_text(caption, term_width):
            print(cap_line)
    print("=" * term_width)

    baseline_time = file_info.get('baseline_time', 'N/A')
    current_time = file_info.get('current_time', 'N/A')
    old_author = file_info.get('old_author', 'N/A')
    new_author = file_info.get('new_author', 'N/A')

    header_addr = pad_line("Address", address_col_width)
    header_base = pad_line(f"Baseline ({baseline_time} by {old_author})", baseline_col_width)
    header_curr = pad_line(f"Current ({current_time} by {new_author})", current_col_width)
    print(f"{header_addr} | {header_base} | {header_curr}")
    print("-" * term_width)

    all_keys = sorted(list(set(old_data.keys()) | set(new_data.keys())))
    if not all_keys:
        print("(No cell changes)")
    else:
        displayed_changes_count = 0
        for key in all_keys:
            if max_display_changes > 0 and displayed_changes_count >= max_display_changes:
                print(f"...(僅顯示前 {max_display_changes} 個變更，總計 {len(all_keys)} 個變更)...")
                break

            old_val = old_data.get(key)
            new_val = new_data.get(key)

            if old_val is not None and new_val is not None:
                if old_val != new_val:
                    old_text = format_cell(old_val)
                    new_text = "[MOD] " + format_cell(new_val)
                else:
                    old_text = format_cell(old_val)
                    new_text = format_cell(new_val)
            elif old_val is not None:
                old_text = format_cell(old_val)
                new_text = "[DEL] (Deleted)"
            else:
                old_text = "(Empty)"
                new_text = "[ADD] " + format_cell(new_val)

            addr_lines = wrap_text(key, address_col_width)
            old_lines = wrap_text(old_text, baseline_col_width)
            new_lines = wrap_text(new_text, current_col_width)
            num_lines = max(len(addr_lines), len(old_lines), len(new_lines))
            for i in range(num_lines):
                a_line = addr_lines[i] if i < len(addr_lines) else ""
                o_line = old_lines[i] if i < len(old_lines) else ""
                n_line = new_lines[i] if i < len(new_lines) else ""
                formatted_a = pad_line(a_line, address_col_width)
                formatted_o = pad_line(o_line, baseline_col_width)
                formatted_n = n_line
                print(f"{formatted_a} | {formatted_o} | {formatted_n}")
            displayed_changes_count += 1
    print("=" * term_width)
    print()

def format_timestamp_for_display(timestamp_str):
    if not timestamp_str or timestamp_str == 'N/A':
        return 'N/A'
    try:
        if 'T' in timestamp_str:
            if '.' in timestamp_str:
                timestamp_str = timestamp_str.split('.')[0]
            return timestamp_str.replace('T', ' ')
        return timestamp_str
    except ValueError as e:
        logging.error(f"格式化時間戳失敗: {timestamp_str}, 錯誤: {e}")
        return timestamp_str

def compare_excel_changes(file_path, silent=False, event_number=None, is_polling=False):
    """
    [最終修正版] 統一日誌記錄和顯示邏輯
    """
    try:
        from core.excel_parser import dump_excel_cells_with_timeout
        
        base_name = os.path.basename(file_path)
        
        old_baseline = load_baseline(base_name)
        if old_baseline is None:
            old_baseline = {}

        current_data = dump_excel_cells_with_timeout(file_path, show_sheet_detail=False, silent=True)
        if not current_data:
            time.sleep(1)
            current_data = dump_excel_cells_with_timeout(file_path, show_sheet_detail=False, silent=True)
            if not current_data:
                if not silent:
                    print(f"❌ 重試後仍無法讀取檔案: {base_name}")
                return False
        
        baseline_cells = old_baseline.get('cells', {})
        if baseline_cells == current_data:
            # 如果是輪詢且無變化，則不顯示任何內容
            if is_polling:
                print(f"    [輪詢檢查] {base_name} 內容無變化。")
            return False
        
        any_sheet_has_changes = False
        
        old_author = old_baseline.get('last_author', 'N/A')
        try:
            new_author = get_excel_last_author(file_path)
        except Exception:
            new_author = 'Unknown'

        for worksheet_name in set(baseline_cells.keys()) | set(current_data.keys()):
            old_ws = baseline_cells.get(worksheet_name, {})
            new_ws = current_data.get(worksheet_name, {})
            
            if old_ws == new_ws:
                continue

            any_sheet_has_changes = True
            
            # 只有在非靜默模式下才顯示和記錄
            if not silent:
                baseline_timestamp = old_baseline.get('timestamp', 'N/A')
                current_timestamp = get_file_mtime(file_path)
                
                # 準備顯示的資料
                all_addresses = set(old_ws.keys()) | set(new_ws.keys())
                display_old = {addr: old_ws.get(addr) for addr in all_addresses if old_ws.get(addr) != new_ws.get(addr)}
                display_new = {addr: new_ws.get(addr) for addr in all_addresses if old_ws.get(addr) != new_ws.get(addr)}

                # 確保比較表格一定顯示
                print_aligned_console_diff(
                    display_old,
                    display_new,
                    {
                        'filename': base_name,
                        'file_path': file_path,
                        'event_number': event_number,
                        'worksheet': worksheet_name,
                        'baseline_time': format_timestamp_for_display(baseline_timestamp),
                        'current_time': format_timestamp_for_display(current_timestamp),
                        'old_author': old_author,
                        'new_author': new_author,
                    },
                    max_display_changes=settings.MAX_CHANGES_TO_DISPLAY
                )
                
                # 分析並記錄有意義的變更
                meaningful_changes = analyze_meaningful_changes(old_ws, new_ws)
                if meaningful_changes:
                    # 只在非輪詢的第一次檢查時記錄日誌，避免重複
                    if not is_polling:
                        log_meaningful_changes_to_csv(file_path, worksheet_name, meaningful_changes, new_author)

        # 只有在非輪詢的第一次檢查且有變更時才更新基準線
        if any_sheet_has_changes and not silent and not is_polling:
            if settings.AUTO_UPDATE_BASELINE_AFTER_COMPARE:
                print(f"🔄 自動更新基準線: {base_name}")
                updated_baseline = {
                    "last_author": new_author,
                    "content_hash": f"updated_{int(time.time())}",
                    "cells": current_data,
                    "timestamp": datetime.now().isoformat()
                }
                from core.baseline import save_baseline
                if not save_baseline(base_name, updated_baseline):
                    print(f"[WARNING] 基準線更新失敗: {base_name}")
        
        return any_sheet_has_changes
        
    except Exception as e:
        if not silent:
            logging.error(f"比較過程出錯: {e}")
        return False

def analyze_meaningful_changes(old_ws, new_ws):
    """
    🧠 分析有意義的變更
    """
    meaningful_changes = []
    all_addresses = set(old_ws.keys()) | set(new_ws.keys())
    
    for addr in all_addresses:
        old_cell = old_ws.get(addr, {})
        new_cell = new_ws.get(addr, {})
        
        if old_cell == new_cell:
            continue

        change_type = classify_change_type(old_cell, new_cell)
        
        # 根據設定過濾變更
        if (change_type == 'FORMULA_CHANGE' and not settings.TRACK_FORMULA_CHANGES) or \
           (change_type == 'DIRECT_VALUE_CHANGE' and not settings.TRACK_DIRECT_VALUE_CHANGES) or \
           (change_type == 'EXTERNAL_REF_UPDATE' and not settings.TRACK_EXTERNAL_REFERENCES) or \
           (change_type == 'INDIRECT_CHANGE' and settings.IGNORE_INDIRECT_CHANGES):
            continue

        meaningful_changes.append({
            'address': addr,
            'old_value': old_cell.get('value'),
            'new_value': new_cell.get('value'),
            'old_formula': old_cell.get('formula'),
            'new_formula': new_cell.get('formula'),
            'change_type': change_type
        })
    
    return meaningful_changes

def classify_change_type(old_cell, new_cell):
    """
    🔍 分類變更類型
    """
    old_val = old_cell.get('value')
    new_val = new_cell.get('value')
    old_formula = old_cell.get('formula')
    new_formula = new_cell.get('formula')
    
    if not old_cell and new_cell: return 'CELL_ADDED'
    if old_cell and not new_cell: return 'CELL_DELETED'
    if old_formula != new_formula: return 'FORMULA_CHANGE'
    if not old_formula and not new_formula and old_val != new_val: return 'DIRECT_VALUE_CHANGE'
    if old_formula and new_formula and old_formula == new_formula and old_val != new_val:
        return 'EXTERNAL_REF_UPDATE' if has_external_reference(old_formula) else 'INDIRECT_CHANGE'
    return 'NO_CHANGE'

def has_external_reference(formula):
    if not formula: return False
    return "['" in formula or "!'" in formula

def log_meaningful_changes_to_csv(file_path, worksheet_name, changes, current_author):
    """
    📝 記錄有意義的變更到 CSV (最終統一版)
    """
    if not current_author or current_author == 'N/A' or not changes:
        return

    try:
        os.makedirs(os.path.dirname(settings.CSV_LOG_FILE), exist_ok=True)
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        file_exists = os.path.exists(settings.CSV_LOG_FILE)
        
        with gzip.open(settings.CSV_LOG_FILE, 'at', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            
            if not file_exists:
                writer.writerow([
                    'Timestamp', 'Filename', 'Worksheet', 'Cell', 'Change_Type',
                    'Old_Value', 'New_Value', 'Old_Formula', 'New_Formula', 'Last_Author'
                ])
            
            for change in changes:
                writer.writerow([
                    timestamp,
                    os.path.basename(file_path),
                    worksheet_name,
                    change['address'],
                    change['change_type'],
                    change.get('old_value', ''),
                    change.get('new_value', ''),
                    change.get('old_formula', ''),
                    change.get('new_formula', ''),
                    current_author
                ])
        
        print(f"📝 {len(changes)} 項變更已記錄到 CSV")
        
    except (OSError, csv.Error) as e:
        logging.error(f"記錄有意義的變-更到 CSV 時發生錯誤: {e}")

# 輔助函數
def set_current_event_number(event_number):
    # 這個函數可能不再需要，但暫時保留
    pass
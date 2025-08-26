import os
import time
import hashlib
import shutil
import logging
import re
import config.settings as settings

_MAX_WIN_FILENAME = 240  # conservative cap to avoid MAX_PATH issues
_HASH_LEN = 16
_PREFIX_SEP = '_'

def _is_in_cache(path: str) -> bool:
    try:
        cache_root = os.path.abspath(settings.CACHE_FOLDER)
        p = os.path.abspath(path)
        return os.path.commonpath([p, cache_root]) == cache_root
    except Exception:
        return False

_def_invalid = re.compile(r'[\\/:*?"<>|]')

def _safe_cache_basename(src_path: str) -> str:
    """Build a safe cache file name: <md5[:16]>_<sanitized-and-trimmed-basename>"""
    base = os.path.basename(src_path)
    base = _def_invalid.sub('_', base)
    name, ext = os.path.splitext(base)
    prefix = hashlib.md5(src_path.encode('utf-8')).hexdigest()[:_HASH_LEN] + _PREFIX_SEP
    # compute allowed length for name part
    allowed = _MAX_WIN_FILENAME - len(prefix) - len(ext)
    if allowed < 8:
        allowed = 8
    if len(name) > allowed:
        name = name[:allowed]
    return f"{prefix}{name}{ext}"

def copy_to_cache(network_path, silent=False):
    if not settings.USE_LOCAL_CACHE:
        return network_path

    try:
        os.makedirs(settings.CACHE_FOLDER, exist_ok=True)

        # If the source already under cache root, return as-is to avoid prefix duplication
        if _is_in_cache(network_path):
            return network_path

        if not os.path.exists(network_path):
            raise FileNotFoundError(f"網絡檔案不存在: {network_path}")
        if not os.access(network_path, os.R_OK):
            raise PermissionError(f"無法讀取網絡檔案: {network_path}")

        cache_file = os.path.join(settings.CACHE_FOLDER, _safe_cache_basename(network_path))

        if os.path.exists(cache_file):
            try:
                if os.path.getmtime(cache_file) >= os.path.getmtime(network_path):
                    return cache_file
            except OSError as e:
                logging.warning(f"獲取緩存檔案時間失敗: {e}")

        network_size = os.path.getsize(network_path)
        if not silent:
            print(f"   📥 複製到緩存: {os.path.basename(network_path)} ({network_size/(1024*1024):.1f} MB)")

        copy_start = time.time()
        shutil.copy2(network_path, cache_file)

        if not silent:
            print(f"      複製完成，耗時 {time.time() - copy_start:.1f} 秒")

        return cache_file

    except FileNotFoundError as e:
        logging.error(f"緩存失敗 - 檔案未找到: {e}")
        if not silent:
            print(f"   ❌ 緩存失敗: {e}")
        return network_path
    except PermissionError as e:
        logging.error(f"緩存失敗 - 權限不足: {e}")
        if not silent:
            print(f"   ❌ 緩存失敗: {e}")
        return network_path
    except OSError as e:
        logging.error(f"緩存失敗 - 複製緩存檔案時發生 I/O 錯誤: {e}")
        if not silent:
            print(f"   ❌ 緩存失敗: {e}")
        return network_path
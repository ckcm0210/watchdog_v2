import os
import time
import hashlib
import shutil
import logging
import re
import io
import csv
from datetime import datetime
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

def _chunked_copy(src: str, dst: str, chunk_mb: int = 4):
    """Optional chunked copy to avoid long single-handle operations (best-effort)."""
    chunk_size = max(1, int(chunk_mb)) * 1024 * 1024
    with open(src, 'rb', buffering=1024 * 1024) as fsrc, open(dst, 'wb', buffering=1024 * 1024) as fdst:
        while True:
            buf = fsrc.read(chunk_size)
            if not buf:
                break
            fdst.write(buf)
    try:
        shutil.copystat(src, dst)
    except Exception:
        pass


def _ops_log_copy_failure(network_path: str, error: Exception, attempts: int, strict_mode: bool):
    try:
        base_dir = os.path.join(settings.LOG_FOLDER, 'ops_log')
        os.makedirs(base_dir, exist_ok=True)
        fname = f"copy_failures_{datetime.now():%Y%m%d}.csv"
        fpath = os.path.join(base_dir, fname)
        new_file = not os.path.exists(fpath)
        with open(fpath, 'a', encoding='utf-8', newline='') as f:
            w = csv.writer(f)
            if new_file:
                w.writerow(['Timestamp','Path','Error','Attempts','STRICT_NO_ORIGINAL_READ','COPY_CHUNK_SIZE_MB','BACKOFF_SEC'])
            w.writerow([
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                network_path,
                str(error),
                attempts,
                bool(getattr(settings, 'STRICT_NO_ORIGINAL_READ', False)),
                int(getattr(settings, 'COPY_CHUNK_SIZE_MB', 0)),
                float(getattr(settings, 'COPY_RETRY_BACKOFF_SEC', 0.0)),
            ])
    except Exception:
        pass

def _ops_log_copy_success(network_path: str, duration: float, attempts: int, engine: str, chunk_mb: int):
    try:
        base_dir = os.path.join(settings.LOG_FOLDER, 'ops_log')
        os.makedirs(base_dir, exist_ok=True)
        fname = f"copy_success_{datetime.now():%Y%m%d}.csv"
        fpath = os.path.join(base_dir, fname)
        new_file = not os.path.exists(fpath)
        with open(fpath, 'a', encoding='utf-8', newline='') as f:
            w = csv.writer(f)
            if new_file:
                w.writerow(['Timestamp','Path','SizeMB','DurationSec','Attempts','Engine','ChunkMB','StabilityChecks','StabilityInterval','StabilityMaxWait','STRICT_NO_ORIGINAL_READ'])
            size_mb = os.path.getsize(network_path)/(1024*1024) if os.path.exists(network_path) else ''
            w.writerow([
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                network_path,
                f"{(size_mb or 0):.2f}" if size_mb != '' else '',
                f"{duration:.2f}",
                attempts,
                engine,
                int(chunk_mb),
                int(getattr(settings, 'COPY_STABILITY_CHECKS', 0)),
                float(getattr(settings, 'COPY_STABILITY_INTERVAL_SEC', 0.0)),
                float(getattr(settings, 'COPY_STABILITY_MAX_WAIT_SEC', 0.0)),
                bool(getattr(settings, 'STRICT_NO_ORIGINAL_READ', False)),
            ])
    except Exception:
        pass


def _wait_for_stable_mtime(path: str, checks: int, interval: float, max_wait: float) -> bool:
    try:
        if checks <= 1:
            return True
        last = None
        same = 0
        start = time.time()
        while True:
            try:
                cur = os.path.getmtime(path)
            except Exception:
                return False
            if last is None:
                last = cur
                same = 1
            else:
                if cur == last:
                    same += 1
                else:
                    same = 1
                    last = cur
            if same >= checks:
                return True
            if max_wait is not None and (time.time() - start) >= max_wait:
                return False
            time.sleep(max(0.0, interval))
    except Exception:
        return False


def _run_subprocess_copy(src: str, dst: str, engine: str = 'robocopy'):
    """Run copy via subprocess engines (robocopy or powershell). dst is full file path."""
    import subprocess, shlex
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    if engine == 'robocopy':
        # robocopy 需要目標目錄 + 檔名分開處理
        src_dir = os.path.dirname(src)
        src_name = os.path.basename(src)
        dst_dir = os.path.dirname(dst)
        # /COPY:DAT 保留日期/屬性/時間；/NJH /NJS /NFL /NDL /NP 降噪；/R:2 /W:1 重試策略
        cmd = [
            'robocopy', src_dir, dst_dir, src_name,
            '/COPY:DAT', '/R:2', '/W:1', '/NJH', '/NJS', '/NFL', '/NDL', '/NP'
        ]
        # robocopy 返回碼 0-7 視為成功
        rc = subprocess.call(cmd, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        if rc > 7:
            raise OSError(f"robocopy rc={rc}")
    elif engine == 'powershell':
        # 使用 PowerShell Copy-Item
        ps_cmd = f"Copy-Item -LiteralPath '{src}' -Destination '{dst}' -Force"
        cmd = ['powershell', '-NoProfile', '-Command', ps_cmd]
        rc = subprocess.call(cmd, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        if rc != 0:
            raise OSError(f"powershell Copy-Item rc={rc}")
    else:
        raise ValueError(f"Unknown subprocess copy engine: {engine}")


def copy_to_cache(network_path, silent=False):
    # 嚴格模式下，如果不使用本地快取，直接返回 None（不讀原檔）
    if not settings.USE_LOCAL_CACHE:
        if getattr(settings, 'STRICT_NO_ORIGINAL_READ', False):
            if not silent:
                print("   ⚠️ 嚴格模式啟用且未啟用本地快取：跳過讀取原檔。")
            return None
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

        # 若快取已新於來源，直接用快取檔
        if os.path.exists(cache_file):
            try:
                if os.path.getmtime(cache_file) >= os.path.getmtime(network_path):
                    return cache_file
            except OSError as e:
                logging.warning(f"獲取緩存檔案時間失敗: {e}")

        network_size = None
        try:
            network_size = os.path.getsize(network_path)
        except Exception:
            pass
        if not silent:
            sz = f" ({network_size/(1024*1024):.1f} MB)" if network_size else ""
            print(f"   📥 複製到緩存: {os.path.basename(network_path)}{sz}")

        retry = max(1, int(getattr(settings, 'COPY_RETRY_COUNT', 3)))
        backoff = max(0.0, float(getattr(settings, 'COPY_RETRY_BACKOFF_SEC', 0.5)))
        chunk_mb = max(0, int(getattr(settings, 'COPY_CHUNK_SIZE_MB', 0)))

        last_err = None
        for attempt in range(1, retry + 1):
            # 複製前穩定性預檢
            st_checks = max(1, int(getattr(settings, 'COPY_STABILITY_CHECKS', 2)))
            st_interval = max(0.0, float(getattr(settings, 'COPY_STABILITY_INTERVAL_SEC', 1.0)))
            st_maxwait = float(getattr(settings, 'COPY_STABILITY_MAX_WAIT_SEC', 3.0))
            if st_checks > 1:
                stable_ok = _wait_for_stable_mtime(network_path, st_checks, st_interval, st_maxwait)
                if not stable_ok:
                    if not silent:
                        print(f"      ⏳ 源檔案仍在變動，延後複製（第 {attempt}/{retry} 次）")
                    time.sleep(backoff * attempt)
                    continue

            copy_start = time.time()
            try:
                # 子程序複製策略：.xlsm 或設定指定時優先
                use_sub = False
                sub_engine = getattr(settings, 'COPY_ENGINE', 'python')
                prefer_xlsm = bool(getattr(settings, 'PREFER_SUBPROCESS_FOR_XLSM', False))
                if sub_engine in ('robocopy', 'powershell'):
                    use_sub = True
                elif prefer_xlsm and str(network_path).lower().endswith('.xlsm'):
                    sub_engine = getattr(settings, 'SUBPROCESS_ENGINE_FOR_XLSM', 'robocopy')
                    use_sub = True

                if use_sub:
                    _run_subprocess_copy(network_path, cache_file, engine=sub_engine)
                else:
                    if chunk_mb > 0:
                        _chunked_copy(network_path, cache_file, chunk_mb=chunk_mb)
                    else:
                        shutil.copy2(network_path, cache_file)
                # 短暫等待，給檔案系統穩定
                time.sleep(getattr(settings, 'COPY_POST_SLEEP_SEC', 0.2))
                duration = time.time() - copy_start
                if not silent:
                    print(f"      複製完成，耗時 {duration:.1f} 秒（第 {attempt}/{retry} 次嘗試）")
                try:
                    _ops_log_copy_success(network_path, duration, attempt, engine='python', chunk_mb=chunk_mb)
                except Exception:
                    pass
                return cache_file
            except (PermissionError, OSError) as e:
                last_err = e
                if not silent:
                    print(f"      ↻ 第 {attempt}/{retry} 次複製失敗：{e}")
                if attempt < retry:
                    time.sleep(backoff * attempt)
                else:
                    break

        # 若最終複製失敗
        if getattr(settings, 'STRICT_NO_ORIGINAL_READ', False):
            logging.error(f"嚴格模式：無法複製到緩存，跳過原檔讀取：{last_err}")
            try:
                _ops_log_copy_failure(network_path, last_err, attempt, True)
            except Exception:
                pass
            if not silent:
                print("   ❌ 複製到快取失敗（嚴格模式：不讀原檔），略過。")
            return None
        else:
            logging.error(f"緩存失敗 - 將回退為直接使用原檔（非嚴格模式）：{last_err}")
            try:
                _ops_log_copy_failure(network_path, last_err, attempt, False)
            except Exception:
                pass
            if not silent:
                print("   ⚠️ 緩存失敗：回退為直接讀原檔（非嚴格模式）")
            return network_path

    except FileNotFoundError as e:
        logging.error(f"緩存失敗 - 檔案未找到: {e}")
        if not silent:
            print(f"   ❌ 緩存失敗: {e}")
        return None if getattr(settings, 'STRICT_NO_ORIGINAL_READ', False) else network_path
    except PermissionError as e:
        logging.error(f"緩存失敗 - 權限不足: {e}")
        if not silent:
            print(f"   ❌ 緩存失敗: {e}")
        return None if getattr(settings, 'STRICT_NO_ORIGINAL_READ', False) else network_path
    except OSError as e:
        logging.error(f"緩存失敗 - 複製緩存檔案時發生 I/O 錯誤: {e}")
        if not silent:
            print(f"   ❌ 緩存失敗: {e}")
        return None if getattr(settings, 'STRICT_NO_ORIGINAL_READ', False) else network_path
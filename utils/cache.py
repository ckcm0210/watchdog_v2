import os
import time
import hashlib
import shutil
import logging
import config.settings as settings

def copy_to_cache(network_path, silent=False):
    if not settings.USE_LOCAL_CACHE:
        return network_path

    try:
        os.makedirs(settings.CACHE_FOLDER, exist_ok=True)
        if not os.path.exists(network_path):
            raise FileNotFoundError(f"網絡檔案不存在: {network_path}")
        if not os.access(network_path, os.R_OK):
            raise PermissionError(f"無法讀取網絡檔案: {network_path}")

        file_hash = hashlib.md5(network_path.encode('utf-8')).hexdigest()[:16]
        cache_file = os.path.join(settings.CACHE_FOLDER, f"{file_hash}_{os.path.basename(network_path)}")

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
# Watchdog 行動方案與技術備忘 (2025-08-28)

作者: Rovo Dev

本文件整合最近討論結論、問題成因、已執行設定調整、以及分階段的落地計劃（Phase 1/2）。本次決策：採用 A 案（Phase 1 全套改動）。

---

## 1) Executive Summary
- 現況：唯一會直接接觸原始 Excel 檔的程式路徑為 `utils/cache.py::copy_to_cache`。其餘分析都在快取副本上完成（`openpyxl`/`zipfile`）。
- 痛點：在 Excel 安全儲存最後階段（臨時檔→覆蓋/rename→metadata）期間，如果輪巡見到 mtime 變就立刻比較（會觸發 copy），容易與 SMB/leasing/oplock、AV/同步掃描產生爭用，造成使用者在 Excel 內 Save 失敗；重啟 A 機後恢復正常。
- 已做設定優化（降低碰撞機率）：
  - COPY_STABILITY_CHECKS=5、COPY_STABILITY_MAX_WAIT_SEC=12.0
  - COPY_RETRY_COUNT=10
  - DEBOUNCE_INTERVAL_SEC=4
  - SPARSE_POLLING_INTERVAL_SEC=60
- 重要提議（已批准進行 Phase 1）：
  1) 快速跳過（mtime/size 不變即跳過 copy/compare）
  2) 穩定窗口先比較（等穩定 N 次/T 秒才進行 copy/compare）
  3) per-file 冷靜期（成功比較後 15–30 秒內不再複製）
  4) 檢測 `~$` 暫存鎖檔（存在則延後觸碰原檔）
  5) 成功複製也寫 ops CSV（現在只記錄失敗）
  6) UI 持久化修復（runtime JSON 合併保存、空值處理、normpath）
- 最後防線（Phase 2，可後續評估）：`COPY_ENGINE=python|powershell|robocopy`，以子程序複製將把手釋放綁定到子進程退出；.xlsm/UNC 可優先使用 robocopy。

---

## 2) 成因與觀察
- `copy_to_cache` 會短時 read-only 打開原檔；我方關檔後，第三方（AV/EDR、OneDrive/SharePoint、Windows Search）仍可能跟進掃描，短時間握住把手。
- SMB/leasing/oplock 在頻繁開關檔案時會切換，偶發殘留（stale handle）或延遲恢復。
- 現輪巡邏輯「見 mtime 變就比較」會在 Excel 保存尾段多次觸發 `copy_to_cache`，放大爭用；.xlsm（含 VBA/簽章/外部連結）保存更長，更易中招。
- 重啟 A 機有效，因為會清掉 A 端的進程/把手/連線與伺服器端租約狀態，第三方 filter 佇列亦重置。

---

## 3) 已執行設定變更（config/settings.py）
- COPY_STABILITY_CHECKS: 2 → 5 （複製前穩定性檢查次數）
- COPY_STABILITY_MAX_WAIT_SEC: 3.0 → 12.0 （穩定等待的最大秒數）
- COPY_RETRY_COUNT: 8 → 10
- DEBOUNCE_INTERVAL_SEC: 2 → 4 （防抖動）
- SPARSE_POLLING_INTERVAL_SEC: 15 → 60
注意：若有 `runtime_settings.json` 或 UI 覆寫，啟動時以 runtime 值為準。

---

## 4) 快速跳過（mtime/size）設計
- 在 baseline JSON 新增：
  - `source_mtime`（float，允許 ±1–2 秒容差）
  - `source_size`（int，bytes）
- 建立/更新 baseline 時（`core/baseline.py`）：
  - 儲存當時 `os.path.getmtime(file)` 與 `os.path.getsize(file)`。
- 比較前（`core/comparison.py`）：
  - 先 `stat` 原檔；若與 baseline 的 `mtime/size` 一致（容差內），直接判定「無變更」，跳過 `copy_to_cache + 讀內容 + 計 hash`。
- 效果：大幅減少重讀，降低在保存尾段觸碰原檔的機會。

---

## 5) Phase 1 變更（已選定方案 A）

1) 穩定窗口先比較（watcher）
- 變更檔：`core/watcher.py::ActivePollingHandler._poll_for_stability`
- 新行為：mtime/size 有變只記錄時間點；只有連續 N 次或 T 秒完全穩定才觸發首次比較（進而觸發 `copy_to_cache`）。
- Console 新增：
  - `[輪巡] 檢測到變動，等待穩定窗口 N/T…`
  - `[輪巡] 已穩定，開始比較…`

2) per-file 冷靜期（cooldown）
- 變更檔：`core/watcher.py`
- 行為：每次成功比較/複製後，對該檔案設置 15–30 秒冷靜期；期間不再觸發新複製/比較。
- Console 新增：`[cooldown] file.xlsx 冷靜期中（xxs），略過本次。`

3) 檢測 `~$` 暫存鎖檔
- 變更檔：`core/watcher.py`
- 行為：輪巡/比較前，如見到 `~$filename` 存在，延長穩定窗口，暫不觸碰原檔。
- Console 新增：`[鎖檔] 偵測到 ~$xxx.xlsx，延後複製/比較。`

4) 快速跳過（mtime/size）
- 變更檔：`core/baseline.py`（寫入 `source_mtime/source_size`）、`core/comparison.py`（比較前快速判斷）
- Console 新增：`[快速通過] mtime/size 未變，略過讀取。`

5) 成功複製也寫 ops CSV
- 變更檔：`utils/cache.py`（新增 `_ops_log_copy_success`）
- 欄位：Timestamp、Path、SizeMB、Duration、Attempts、ChunkMB、StabilityParams、STRICT、Engine、Thread。

6) UI 持久化修復
- 變更檔：`ui/settings_ui.py`
- 內容：
  - 保存時改為「合併保存」：先 load runtime JSON，再用畫面值覆蓋，未出現的鍵保持原值。
  - Path 欄位統一 `os.path.normpath`；空字串不覆蓋既有值或保存為 `None`。
  - 啟動時優先採用 runtime 值（現有邏輯已做，將加強檢查）。
- 效果：例如 `CACHE_FOLDER` 不會在重開 UI 後變回空白。

---

## 6) Phase 2（可選的最後防線：子程序複製引擎）
- 新增設定：`COPY_ENGINE = 'python' | 'powershell' | 'robocopy'`
- `utils/cache.py::copy_to_cache` 支援三種引擎：
  - python：沿用現狀（`open`/`shutil.copy2`/`_chunked_copy`）
  - powershell：`powershell -NoProfile -Command "Copy-Item -LiteralPath 'src' -Destination 'dst' -Force"`
  - robocopy：`robocopy "src_dir" "dst_dir" "filename" /R:2 /W:1 /COPY:DAT /NJH /NJS /NFL /NDL /NP`
    - 返回碼處理：0–7 視為成功，>7 為失敗
- 可選策略：對 `.xlsm` 或 UNC 路徑優先用 robocopy；多次失敗自動切換引擎。
- 效果：把手釋放與子進程生命週期綁定，降低殘留風險；在網絡路徑常更穩定。

---

## 7) 測試與接受準則
- `.xlsm` 壓力測試：連續快速 Save 5–10 次
  - 預期：出現比較輸出次數下降；Console 顯示「等待穩定」「冷靜期」「快速通過」等訊息；`ops_log` 有成功/失敗 copy 記錄
  - B 機 Save 失敗機率顯著下降，不需重啟 A 機即可恢復
- `UI` 重開驗證：`CACHE_FOLDER` 等欄位能準確帶出上次值

---

## 8) 風險與回退
- 快速跳過使用 mtime/size：極少數情境下可能 mtime 被保留或對齊；雙判斷已降低風險，如有疑慮可臨時關閉快速跳過開關（將提供 flag）。
- watcher 穩定窗口：如使用者期望「極即時」輸出，等待窗口會稍延遲首次表格；可通過 UI 可調參數折衝。
- 子程序複製：依賴外部工具（Windows 內置），須處理返回碼；保留 python 引擎作回退。

---

## 9) 待確認事項
- 是否將 `COPY_RETRY_BACKOFF_SEC` 從 1.0 升至 1.5/2.0？
- Phase 2 是否預設 `COPY_ENGINE=robocopy`（Windows 環境）？或先保持 python，引 `.xlsm` 自動升級？
- AV/同步（如 OneDrive/SharePoint）能否對 WATCH_FOLDERS 設排除即時掃描？

---

## 10) 後續執行清單（Phase 1）
- [ ] baseline JSON 寫入 `source_mtime/source_size`
- [ ] comparison 進入前快速判斷 `mtime/size` 是否一致（容差）
- [ ] watcher：改為「穩定窗口先比較」
- [ ] watcher：加入 per-file 冷靜期（預設 15–30 秒，可調）
- [ ] watcher：偵測 `~$` 暫存鎖檔，延後觸碰
- [ ] cache：新增 `_ops_log_copy_success`（CSV）
- [ ] UI：保存時合併 JSON、空值處理、normpath，加強載入邏輯
- [ ] Console：新增關鍵訊息（等待穩定/冷靜期/快速通過/引擎）

---

## 11) 參考與相關檔
- 設定檔：`config/settings.py`、`config/runtime_settings.json`
- 複製與快取：`utils/cache.py`
- Excel 解析：`core/excel_parser.py`
- 比對流程：`core/comparison.py`
- 基準線：`core/baseline.py`
- 監看與輪巡：`core/watcher.py`
- UI 設定：`ui/settings_ui.py`
- 調查記錄：`docs/Excel_File_Lock_Analysis_2025-08-28.md`（第 10–12 節）

---

## 12) 實施變更記錄（Phase 1 本次提交）
- 設定新增：QUICK_SKIP_BY_STAT、MTIME_TOLERANCE_SEC、POLLING_STABLE_CHECKS、POLLING_COOLDOWN_SEC、SKIP_WHEN_TEMP_LOCK_PRESENT
- 設定調整（早前）：COPY_STABILITY_CHECKS=5、COPY_STABILITY_MAX_WAIT_SEC=12.0、COPY_RETRY_COUNT=10、DEBOUNCE_INTERVAL_SEC=4、SPARSE_POLLING_INTERVAL_SEC=60
- baseline：在 baseline JSON 寫入 source_mtime/source_size（core/baseline.py）、比較後更新同步寫入（core/comparison.py）
- comparison：比較前快速判斷 mtime/size 未變則跳過讀取（core/comparison.py），Console 輸出「[快速通過] …」
- watcher：改為穩定窗口先比較 + 冷靜期 + 偵測 ~$ 暫存鎖檔（core/watcher.py），Console 輸出「[輪巡] 等待穩定」「[輪巡] 已穩定」「[cooldown]」「[鎖檔]」
- cache：新增成功複製 CSV 記錄（utils/cache.py::_ops_log_copy_success），Console 顯示耗時
- UI：儲存時合併 runtime JSON、空值不覆蓋、normpath（ui/settings_ui.py）

## 13) 程式碼審視與可改善點（避免鎖檔為首要）
- watcher._is_log_ignored 重複殘段：清理函數尾部多餘的 _is_cache_ignored 片段，提升可讀性。
- stop() 的狀態清理：目前 stop() 只清空 polling_tasks，建議同步清空 ActivePollingHandler.state 以免長期殘留。
- ENABLE_OPS_LOG 開關：在 _ops_log_copy_success/_ops_log_copy_failure 寫 CSV 前加判斷，當關閉時完全不寫檔。
- runtime 設定型別校驗：runtime_settings.json 出現過整數欄位存成路徑字串的錯型（例如 CONSOLE_INITIAL_TOPMOST_DURATION），建議 UI 存檔前做基本型別驗證；錯型自動回復預設或提示修正。
- 快速跳過容差優化：目前用 abs(mtime 差) <= MTIME_TOLERANCE_SEC 並 size 相同即可跳過。可選擇把 mtime 先四捨五入至 1–2 秒精度再比對，進一步降低跨檔案系統精度差造成的誤判（非必須）。
- UI 清空欄位的體驗：合併保存策略下，空值不覆蓋舊值；如用戶真的想清空某個欄位（例如排除資料夾），可加「清空/重設」按鈕以明確處理。

進一步避免鎖檔的可行加強（在現有基礎上）：
- 在 copy 前的穩定窗口同時檢查 size 穩定（可選），尤其對 .xlsm。
- 在「活動激烈期」動態放寬輪巡（延長 interval、提高 POLLING_STABLE_CHECKS），讓編輯/保存最後階段遠離複製。
- 大規模 refresh（大量 EXTERNAL_REF_UPDATE）時印/寫摘要行，減輕表格輸出壓力（搭配 MAX_CHANGES_TO_DISPLAY）。
- 監控停止時的「安全降落」：stop() 後短暫 sleep 等待子進程複製結束，確保不留把手。

最重要的落地清單（可立即採用/驗證）：
- 保持 STRICT_NO_ORIGINAL_READ=True，並使用 QUICK_SKIP_BY_STAT=True。
- 穩定窗口先比較 + 冷靜期 + 檢測 ~$ 暫存檔（已落地）。
- .xlsm 優先子程序複製；如仍有問題，可全局設 COPY_ENGINE='robocopy'（可在 UI 設定）。
- 視情況把 COPY_RETRY_BACKOFF_SEC 提升至 1.5–2.0（更溫和的退避）。
- 與 IT 協作：為 WATCH_FOLDERS 設即時掃描排除，保存高峰期暫停 OneDrive/SharePoint。

## 14) 不中斷監控地動態增刪目錄（可行性與做法）
可行性：可以。watchdog 允許在 Observer 運行期間 schedule/unschedule 新的監控路徑，無需停止整個 observer。

設計方案（WatcherManager，線上增刪）：
- 保持現有 Observer 與 ExcelFileEventHandler 不變，額外新增一個管理器（可放在 core/watcher.py 或 main.py），負責：
  - 保存 path → ObservedWatch 的映射（observer.schedule 的回傳值）。
  - 提供 add_watch(path) 與 remove_watch(path) 方法，內部做 thread-safe（加鎖）。
  - 新增時：
    1) 正規化與檢查（存在且可讀；不在被忽略的 CACHE_FOLDER/LOG_FOLDER 底下）。
    2) 調用 observer.schedule(event_handler, path, recursive=True) 取得 watch，寫入映射；console 印「[ADD WATCH]」。
  - 移除時：
    1) 從映射取出 watch，調用 observer.unschedule(watch) 移除；console 印「[REMOVE WATCH]」。
  - 避免重覆註冊：新增前檢查此 path 是否已存在於映射。
- UI/設定互動（兩種方式）：
  1) UI 新增「運行期間更新」小頁：提供「新增監控目錄」「移除監控目錄」控制；按下即調用 WatcherManager。
  2) 設定檔（runtime_settings.json）輪詢法：每隔 N 秒讀取一個 runtime_updates.json（或重用 runtime_settings.json）鍵值，發現 WATCH_FOLDERS 有變更時，對差集做 add/remove（此法簡單，但有輕微輪詢開銷）。
- 無停機：整個過程 observer 一直運行；schedule/unschedule 只影響增刪的目錄。
- 注意點：
  - 移除深層子路徑時，如其父路徑仍被監控，會繼續收到事件；必要時設計更細粒度的忽略規則。
  - Windows 權限/鎖檔導致 schedule 失敗時須友善提示/重試。

落地步驟（最少代碼）：
1) 在 main.py 保存 observer 與一個新的 WatcherManager（保存 path→watch 字典）。
2) 把原本 for folder in watch_roots 的 schedule 改為調用 manager.add_watch(folder)。
3) 新增一個簡單指令/函式可在運行中調用 manager.add_watch/remove_watch（先用函式；日後再接 UI）。
4) 如採用設定檔輪詢法，可在背景 thread 每 10–30s 比對 WATCH_FOLDERS 差集，對新增/移除路徑呼叫 manager 方法。

## 15) Handover / 待辦清單（明日續辦）
- 清理 watcher._is_log_ignored 尾部重複殘段；stop() 同步清空 ActivePollingHandler.state。
- 在 ops CSV 寫入處加 ENABLE_OPS_LOG 判斷，允許一鍵關閉寫檔。
- copy 前 size 穩定窗口（可選）開關與實作；或只在 .xlsm 啟用。
- comparison 快速跳過的 mtime 取整策略（非必須，可觀察後再做）。
- 大規模 refresh 概要輸出（console 與 CSV），避免過長輸出影響體驗。
- UI 數值型欄位的型別驗證（整數/小數），避免錯型寫入 runtime。
- 動態增刪監控目錄：
  - 設計/實作 WatcherManager，保留 path→ObservedWatch 映射；
  - main.py 改為透過 manager.add_watch() 註冊；
  - 選擇 UI 即時控制 或 設定檔輪詢法；
  - 測試在不停機狀態下新增/移除路徑能即時生效。

本文件會在落地過程中持續更新（加入實作進度、測試結果與調整記錄）。

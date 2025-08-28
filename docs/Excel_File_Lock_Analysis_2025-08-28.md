# Excel 原生檔「被鎖/無法儲存」調查筆記與行動方案 (2025-08-28)

- 建立時間 (UTC): 2025-08-28 09:11:06
- 建立者: ckcm0210
- 相關程式：watchdog_v2（Excel 變更監測工具）

## 1) 現象描述
- 在監控程式運行期間，部分 Excel 原生檔（特別是 .xlsm，有時 .xlsx）偶發出現「使用者在 Excel 內 Save 失敗」。
- 停止監控程式未必即時恢復；有時需要重啟電腦，使用者才可正常儲存。
- 監控程式的邏輯是：先複製原檔到本機 CACHE_FOLDER，再用 openpyxl/zip 於快取檔上做解析與比較；功能正常，亦能偵測與輸出差異。

## 2) 目前設計（與可能影響點）
- 複製：utils/cache.copy_to_cache 會對「原生檔」做 read-only 開檔後複製至快取，再於快取檔上讀取、比較。
  - 即使是 read-only，當 Excel 正在進行「安全儲存」（多階段寫入/覆蓋/重命名）時，網路/本機檔案系統的共享模式仍可能出現爭用。
  - 程式已加入「mtime 穩定性檢查」「重試/退避」「分塊複製」與每日 ops log（copy_failures_YYYYMMDD.csv）。
- 讀取：core/excel_parser.dump_excel_cells_with_timeout 只在「快取檔」上 openpyxl.load_workbook(..., read_only=True)，並確保 wb.close()/del wb。
- 作者：get_excel_last_author 先用 zipfile 讀快取檔 docProps/core.xml，而非讀原檔；失敗才以 openpyxl 讀「快取檔」。
- 輪巡：偵測到變更後，ActivePollingHandler 會按檔案大小做定期檢查；當偵測到 mtime 變更時，現行實作會即刻再跑一次比較（會觸發複製）。

## 3) 觀察與推論
- 「mtime 變更 ≠ 已完全釋放檔案鎖」。在 Windows/SMB 環境，Excel 的安全儲存可能在 mtime 變更後，仍持續握住/切換鎖或使用暫存檔 → 見到 mtime 穩定一陣，並不代表鎖一定已釋放。
- 輪巡期內，如頻繁嘗試複製大 .xlsm，容易撞正 Excel 正在進行最後階段保存，形成共享違規（我方讀原檔 vs Excel 寫原檔/覆蓋）。
- 即使停止監控程式，仍需重啟電腦才可儲存，可能原因：
  1) 第三方 filter/防毒/同步工具受我們的頻繁讀取觸發，殘留把手/延遲釋放。
  2) 網路驅動/SMB oplock 狀態異常（stale handle/oplock break），導致共享狀態不能即刻恢復。
  3) 監控程式未完全結束（殭屍 thread/timer），仍有殘留把手（需檢查 signal/stop 流程與 Timer 清理）。
- .xlsm 較常中招：檔案大、含 VBA/簽章/外部連結，保存流程較複雜和耗時 → 更易與頻密複製衝突。

## 4) 我們需要的證據/紀錄（方便根因定位）
- LOG_FOLDER/ops_log/copy_failures_*.csv：觀察是否在使用者 Save 失敗時段，copy 嘗試密集失敗/重試。
- 用 Sysinternals Handle/Process Explorer 鎖定追蹤：看到底是 python.exe、Excel、還是 AV/同步客戶端握住把手。
- Windows 事件檢視器（系統、應用程式）與檔案伺服器端的 SMB 記錄（如適用）。
- Excel 版本與儲存選項（是否啟用自動回復、OneDrive/SharePoint 同步、受信任位置等）。

## 5) 立即可做的「無改碼」緩解（建議設定）
- 保證永不直接以原檔做重讀：
  - USE_LOCAL_CACHE=True；STRICT_NO_ORIGINAL_READ=True（如快取失敗即跳過，不讀原檔）。
  - 確保 CACHE_FOLDER 在監察範圍外，並 IGNORE_CACHE_FOLDER=True。
- 放寬保存穩定窗口與減少輪巡干擾：
  - COPY_STABILITY_CHECKS=5、COPY_STABILITY_INTERVAL_SEC=1.0~1.5、COPY_STABILITY_MAX_WAIT_SEC=10~15。
  - COPY_RETRY_COUNT=8~10、COPY_RETRY_BACKOFF_SEC=1.0~2.0、COPY_CHUNK_SIZE_MB=4。
  - 調高 SPARSE_POLLING_INTERVAL_SEC（例如 60s）、DEBOUNCE_INTERVAL_SEC（3~5s），減少在使用者活躍保存期的碰撞。
  - 針對 .xlsm 個別路徑，可加白名單延長 quiescent window 或暫時只監察 metadata（先不比較）。

## 6) 建議的中期改碼（降低保存期的觸碰機率）
- Polling 由「mtime 一變就比較」改為「先等穩定」：只有當 mtime 連續 N 秒無變（穩定窗口）才執行 copy/compare；變更期間只更新 last_mtime，不做比較。
- 比較後即時更新 baseline，避免在輪巡期重覆對同一批變更做多次完整比較。
- 一致性的關閉流程：確保 Ctrl+C/信號處理會停止所有 Timer/thread 並清理（避免殘留把手令使用者需重啟電腦）。
- 在 UI/設定檔正式暴露並對齊：USE_LOCAL_CACHE、STRICT_NO_ORIGINAL_READ、COPY_* 相關鍵值及預設。

## 7) 驗證計畫
- 以大容量 .xlsm（含 VBA/外部連結）做測試：
  - A 組（現狀/高頻輪巡） vs B 組（穩定窗口策略 + 上述設定）。
  - 操作腳本：持續編輯/存檔 5~10 次，觀察 Save 成功率、ops_log 失敗次數、是否需要重啟。
- 用 Handle/Process Explorer 記錄在問題時段每 5s snapshot，確認握把手的進程是否來自監控程式或第三方。

## 8) 開放問題
- 停監控仍需重啟電腦的案例，是否與網絡驅動/防毒 filter/同步客戶端有關？需聯同 IT 安全部門交叉檢。
- 部分環境對 read-only 打開亦會造成 Excel Save 失敗？（特定檔案伺服器/權限/檔案屬性）

## 9) 參考代碼路徑
- 複製與穩定檢查：utils/cache.py → copy_to_cache、_wait_for_stable_mtime
- 解析與關閉：core/excel_parser.py → dump_excel_cells_with_timeout、safe_load_workbook、get_excel_last_author
- 比較流程：core/comparison.py → compare_excel_changes
- 監看與輪巡：core/watcher.py → ActivePollingHandler._poll_for_stability、ExcelFileEventHandler

---
此文件為調查筆記，後續如有新的證據或修正，請繼續補充。
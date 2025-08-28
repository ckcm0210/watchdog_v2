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


好，我哋用現有檢索到的代碼片段，集中喺 core 同 utils 做逐行級分析，標出會影響「鎖檔／Save 失敗」的關鍵觸點。注意：以下代碼摘錄源自快速檢索，可能不完整；建議用 GitHub 介面瀏覽全部內容以交叉核對：

## 10) 用家補充與回應（2025-08-28）

A. 只係 copy 先會接觸到原檔？
- 對。現時設計入面，唯一會直接打開「原始 Excel 檔」嘅步驟，就係 utils/cache.py 的 copy_to_cache（用 open/shutil.copy2 或分塊讀）。之後所有重讀都係針對「快取檔」做（openpyxl、zipfile），唔會再摸原檔。

B. 「copy 完之後仍然鎖住檔案」點解會發生？
- 我方程式碼喺複製階段用 with 開檔，理論上複製完成即釋放把手，並有短暫 sleep（COPY_POST_SLEEP_SEC）。但仍可能出現以下情況：
  1) 正撞正 Excel 安全儲存最後階段（覆蓋/rename/flush），SMB/oplock/AV filter 有延遲釋放或重試機制，導致使用者感覺「copy 完仍然被鎖」。
  2) 同期第三方（防毒、同步客戶端、索引器）因為我方複製行為而跟入掃描，握住把手未放。
  3) 我方輪巡期內再次嘗試 copy，對原檔造成連續讀取壓力（見下一節）。

C. 最後一道屏障（Last resort）— 如何「確保 copy 完一定唔會再鎖住原檔」？
- 可行做法（由低風險到高改動）：
  1) 加長並嚴格化「穩定窗口」：
     - 在 copy 前要求 mtime/size 連續 N 次（例如 N=5，間隔 1.0–1.5s）完全不變，且若偵測到暫存鎖檔（~$xxx.xlsx）存在則延長等待。
  2) copy 後的清理儀式：
     - 以 with 確保 close；顯式 del 變數、gc.collect()；保留 200–500ms sleep（已有 COPY_POST_SLEEP_SEC）。
  3) 冷靜期（cooldown）：
     - 在同一檔案成功複製後，設定至少 8–15 秒的 per-file cooldown；冷靜期內即使 mtime 輕微波動都唔再觸發新一次 copy。
  4) 「穩定窗口先比較」策略（建議實作）：
     - 輪巡時唔係見到 mtime 就即刻比較，而係要連續 T 秒或連續 N 次檢查「完全無變」先執行 copy/compare，明顯降低與 Excel 最後保存階段的碰撞。
  5) 子程序複製（可選）：
     - 提供一個選項用 Windows 原生 robocopy/Copy-Item（子進程）做複製；子進程結束即代表所有把手已釋放，可作為極端環境的後備策略（代價：需依賴系統工具）。
  6) 明確拒絕「回退直讀原檔」：
     - 保持 STRICT_NO_ORIGINAL_READ=True（現已如此）。若複製失敗就跳過，永不在原檔上用 openpyxl 讀。

D. 「如果 copy 期間啱啱撞正 user save 檔案會出事」是否正確？
- 可以咁理解：在 Windows/SMB/AV filter 環境，Excel 安全儲存係多步驟操作。當我方喺最後幾秒反覆打開來源檔進行 copy，就容易產生共享衝突或延遲釋放，把手錯覺為被鎖。透過上面 C(1)(4) 對應策略可大幅減少碰撞概率。

E. 什麼叫「輪巡期內過於頻繁嘗試複製」？（解釋）
- 目前輪巡邏輯（core/watcher.py::ActivePollingHandler._poll_for_stability）係：
  - 只要偵測到 mtime 變咗，就即刻走一次 compare。compare 又會觸發 copy_to_cache → 讀原檔。
  - 如果使用者連續幾次快速 Save，或者 Excel 在最後階段仍多次更新 mtime，就會形成「幾十秒內多次嘗試 copy」的情況。
  - 這種高頻觸碰會放大與 Excel/SMB/AV 的共享爭用機率，亦可能造成使用者 Save 失敗或需重啟先恢復。
- 改善：改為「穩定窗口先比較」：近期內 mtime 有變只記錄最後一次時間戳，等到連續 N 次檢查（或 T 秒）都無變先做第一次比較/複製。

F. 建議把這些策略落實到設定/代碼
- 設定層（即刻可調）：
  - COPY_STABILITY_CHECKS=5、COPY_STABILITY_INTERVAL_SEC=1.0–1.5、COPY_STABILITY_MAX_WAIT_SEC=10–15
  - COPY_RETRY_COUNT=8–10、COPY_RETRY_BACKOFF_SEC=1.0–2.0、COPY_CHUNK_SIZE_MB=4
  - SPARSE_POLLING_INTERVAL_SEC=60、DEBOUNCE_INTERVAL_SEC=3–5
  - 保持 STRICT_NO_ORIGINAL_READ=True、IGNORE_CACHE_FOLDER=True
- 代碼層（短期）：
  - 在 copy 成功亦寫入 ops CSV（目前只記錄失敗），方便對齊 Save 失敗時段分析。
  - watcher 改成「穩定窗口先比較」；對偵測到 ~$ 文件存在時延長窗口；複製成功後加入 per-file 冷靜期。
- 代碼層（可選）：
  - 子程序複製選項（robocopy / Copy-Item / xcopy），作為極端環境的後備路徑，確保子進程退出即不留把手。

## 11) 點解「copy 完仍然鎖住，要 reboot 先復原」？（深入解釋）
- 正常情況：我方複製步驟用 with 開檔，複製完成即釋放把手；理論上 Excel 應可繼續 Save。
- 實務上仍可能被鎖的原因：
  1) 第三方模組延遲掃描：防毒/EDR、OneDrive/SharePoint 同步客戶端、Windows Search 索引器，見到我方剛讀過原檔就跟進掃描，短時間握住把手。
  2) SMB/網絡 leasing/oplock 狀態：頻繁的打開/關閉會令 client/server 之間的租約/鎖狀態切換，偶發殘留（stale handle），需要等超時或重連才完全釋放。
  3) Excel 安全儲存尾段多步驟：臨時檔寫入→覆蓋/rename→寫入屬性/metadata。呢段時間 mtime/size 會多次跳動，若我方輪巡「見動就 copy」，會在尾段反覆觸碰來源檔，放大爭用。
  4) 我方 timer/thread 未完全停止（極少見）：若程式非優雅關閉，殘留計時器仍可能觸發 copy_to_cache（現版本已有 stop/cancel 保護）。
- 點解 reboot 有效：重啟會重置進程、檔案系統快取、SMB 連線、oplock/leasing 狀態、以及第三方 filter driver 內部佇列；所有把手被強制清走，Excel 自然可正常 Save。
- 回答你的直覺問題：「copy 完咗，點解仲鎖？」：因為「鎖」唔一定係我方程式；更多時候係被我方行為觸發咗其他系統/第三方模組去掃描，佢哋握住把手。避免方法係減少尾段高頻觸碰，並在 copy 後給足夠冷靜期讓其他模組自然釋放。

## 12) 「軟重啟」級對策（不重啟電腦，但盡量達到同等效果）
- 在應用層可即刻做：
  - 穩定窗口先比較：輪巡期只在連續 N 次或 T 秒完全穩定先觸發首次 copy/compare，避開 Excel Save 尾段。
  - per-file 冷靜期：每次成功複製之後，對該檔案強制 15–30 秒冷靜期；期間即使 mtime 微動都唔再觸發 copy。
  - 嚴格模式：保持 STRICT_NO_ORIGINAL_READ=True；複製失敗就 skip，永不直讀原檔。
  - 偵測 ~$.xls*：如見到 Office 暫存鎖檔，延長所有等待時間，暫不觸碰原檔。
  - 清理儀式：copy 後顯式關閉/刪除檔案物件、gc.collect()，並保留少量 sleep（COPY_POST_SLEEP_SEC）。
  - 子程序複製（可選）：改用 robocopy/Copy-Item 作為子進程，子進程退出即代表系統層把手已收回（在某些環境比 Python open 更少干擾）。
  - 觀察與退避：若短時間內多次 copy 失敗或比較連續觸發，將該檔案退避到 1–5 分鐘再試。
- 在監控層可做：
  - 提供「暫停監控此檔/此目錄 X 分鐘」功能，等候外部模組自然釋放。
  - 停止前確保 observer/timer 全部 cancel + join，避免殘留 thread。
- 進階（Windows 專用，可選）：
  - 若允許依賴 pywin32，可用 CreateFile 設定更寬鬆的共享旗標（FILE_SHARE_READ|WRITE|DELETE）讀取，以減少與 Excel 的共享衝突（需充分測試）。
- 系統/IT 協作：
  - 用 Handle/Process Explorer 確認握把手的進程（Excel/AV/同步客戶端等），必要時暫停對該路徑的即時掃描或同步。
  - 重置 SMB 連線（例如 net use \\server\share /delete），或暫停/重啟 OneDrive/Office Upload Center。

core 目錄搜索結果：https://github.com/ckcm0210/watchdog_v2/search?q=path%3A%5Ecore%5C%2F.*%24
utils 目錄搜索結果：https://github.com/ckcm0210/watchdog_v2/search?q=path%3A%5Eutils%5C%2F.*%24
一、事件入口與輪巡（core/watcher.py）

檔案與連結：core/watcher.py
https://github.com/ckcm0210/watchdog_v2/blob/main/core/watcher.py
重點片段（已檢索到的部分）：初始化後，以檔案大小決定輪巡間隔，並以 getmtime 作為變更判斷：
start_polling()：計算檔案大小，決定 interval；讀取初始 mtime；啟動計時器。
_start_adaptive_polling()：用 threading.Timer 按 interval 呼叫 _poll_for_stability。
_poll_for_stability()：以 os.path.getmtime 比較 current_mtime vs last_mtime 決定 has_changes，然後延長或結束輪巡。
對鎖檔的含意：
這層本身只做 mtime/stat，不會直接鎖原檔；真正會「開檔」的是之後的比對流程（舊實作 watcher-Copy1 可見直接呼叫 compare_excel_changes）。
舊版參考（有助理解流程）（core/watcher-Copy1.py）
連結：https://github.com/ckcm0210/watchdog_v2/blob/main/core/watcher-Copy1.py
_poll_dense() 會呼叫 compare_excel_changes(...)，而 compare_excel_changes 會間接觸發 copy_to_cache → 複製原檔到快取。
二、複製與原檔讀取的唯一入口（utils/cache.py）

檔案與連結：utils/cache.py
https://github.com/ckcm0210/watchdog_v2/blob/main/utils/cache.py
關鍵函數：copy_to_cache(network_path, silent=False)
嚴格模式與快取開關：
如果 USE_LOCAL_CACHE=False 且 STRICT_NO_ORIGINAL_READ=True → 直接返回 None（完全不讀原檔）。
如果 USE_LOCAL_CACHE=True（推薦），後續所有重讀都在快取檔上做。
快取命名與去重：_safe_cache_basename，確保檔名安全，避免 cache 路徑重複。
快取新舊判斷：如果 cache_file 的 mtime >= network_path 的 mtime → 直接用快取（避免讀原檔）。
複製重試與穩定性預檢：
讀取 COPY_RETRY_COUNT、COPY_RETRY_BACKOFF_SEC、COPY_CHUNK_SIZE_MB。
在每次複製前，_wait_for_stable_mtime(network_path, checks, interval, max_wait) 檢查 mtime 是否在一段時間內保持穩定；否則延後重試。
真正觸碰原檔的地方：
若 chunk_mb > 0 → _chunked_copy(network_path, cache_file, chunk_mb)：以 open(src,'rb') 分塊讀取，長時間持有 read handle。
否則 → shutil.copy2(network_path, cache_file)：同樣會 open 原檔進行讀取。
這裡就係最可能與 Excel Save 爭用的點：Excel 在安全儲存的最後階段（覆蓋/rename/flush）時，對來源檔的共享模式可能與我們的 read-only 打開衝突（尤其是網絡 SMB/AV filter 同步下）。
其他相關輔助：
_wait_for_stable_mtime(...)：確保連續多次 mtime 不變才放行複製（檢索到開頭，實作邏輯為輪詢 mtime，達到 checks 次相同或超時）。
_chunked_copy(...)：分塊讀寫，最後 copystat；優點是減少單一長時間操作，但仍然持續讀取來源。
三、Excel 解析與作者資訊（core/excel_parser.py）

檔案與連結：core/excel_parser.py
https://github.com/ckcm0210/watchdog_v2/blob/main/core/excel_parser.py
已見重點：
import copy_to_cache：所有需要讀 Excel 的流程，應該先拿到本地快取路徑，再對快取檔操作。
extract_external_refs(xlsx_path)：直接用 zipfile.ZipFile 在 xlsx 上讀 externalLinks 與 rels；預期 xlsx_path 應為快取檔。
_normalize_path()：處理 file:// 與 UNC 路徑格式，避免誤解路徑。
推論（基於命名與常見模式）：dump_excel_cells_with_timeout / safe_load_workbook / get_excel_last_author 會對「快取檔」使用 openpyxl.load_workbook(..., read_only=True) 或 zipfile 讀核心屬性；不會再碰原檔。
四、比對流程與基準線存取（core/comparison.py, core/baseline.py）

檔案與連結：
core/comparison.py：https://github.com/ckcm0210/watchdog_v2/blob/main/core/comparison.py
core/baseline.py：https://github.com/ckcm0210/watchdog_v2/blob/main/core/baseline.py
compare_excel_changes() 會：
讀 baseline（壓縮格式支援在 utils/compression），
呼叫 excel_parser 產出當前快取檔內容、最後作者等，
然後做差異列印。
這兩個模組不直接開「原檔」，只會驅動 excel_parser → cache 的路徑。因此「輪巡頻率」與「是否頻繁觸發 copy_to_cache」才是爭用風險來源。
五、關於「停程式都要重啟電腦才可 Save」的角度

watcher 使用 threading.Timer 啟動多個定時工作，並以 stop_event 控制；需要確保所有 timer 都被 cancel、無殘留 background thread 正在嘗試 copy_to_cache（即仍持有來源檔讀把手）。
即使本程式已停，先前的頻繁讀取可能讓 AV/同步/索引器接手掃描該檔，導致把手延後釋放；或 SMB oplock 狀態未即時恢復。這些會令使用者感覺要重啟才回復正常。
六、逐行核對重點清單（建議按此路徑過數）

utils/cache.py
copy_to_cache：逐行確認所有可能 open 原檔的地方（shutil.copy2、_chunked_copy），以及遇錯誤時的重試頻率與 backoff。確定 STRICT_NO_ORIGINAL_READ/USE_LOCAL_CACHE 行為符合預期。
_wait_for_stable_mtime：確認 checks/interval/max_wait 的實作是否可避免 Save 過程中最後數秒的覆蓋／rename。
_chunked_copy：確認 open/close 時機、copystat 是否可能再觸發存取。
core/excel_parser.py
所有 load_workbook/zipfile.ZipFile 的入參是否「必定來自 cache」，避免任何路徑分支導致直讀原檔。
wb.close()/檔案 handle 是否可靠釋放（finally 區塊）。
core/watcher.py
_poll_for_stability 內部是否仍有任何「立即比較」的邏輯（當 mtime 改變時）會過度頻繁地觸發 copy_to_cache。
stop 流程是否會把所有 Timer cancel 與 join（避免殭屍 timer 再觸發）。
core/comparison.py / core/baseline.py
確認它們不會自己打開 network_path，而一律透過 excel_parser → cache。
七、直接可查閱的檔案連結（方便比對）

watcher（新）：https://github.com/ckcm0210/watchdog_v2/blob/main/core/watcher.py
watcher（舊參考）：https://github.com/ckcm0210/watchdog_v2/blob/main/core/watcher-Copy1.py
cache：https://github.com/ckcm0210/watchdog_v2/blob/main/utils/cache.py
excel_parser：https://github.com/ckcm0210/watchdog_v2/blob/main/core/excel_parser.py
comparison：https://github.com/ckcm0210/watchdog_v2/blob/main/core/comparison.py
baseline：https://github.com/ckcm0210/watchdog_v2/blob/main/core/baseline.py

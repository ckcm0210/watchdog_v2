# Watchdog Runbook (2025-08-27)

Audience: Operators and non-developer users who need to run and observe Excel Watchdog.

Quick start
1) Install requirements (Python 3.10+ recommended)
   - pip install -r requirements.txt
2) Start the tool
   - python main.py
3) Settings UI (first run)
   - Configure:
     - WATCH_FOLDERS (roots to monitor)
     - SCAN_TARGET_FOLDERS (optional subset for startup baseline)
     - MONITOR_ONLY_FOLDERS (monitor changes first, create baseline on first change)
     - CACHE_FOLDER (must be outside watched roots to avoid noise)
     - LOG_FOLDER (where baselines and logs go)
     - Compression format (lz4/zstd/gzip) and levels
     - Timeout/Memory monitoring options
   - Save and start.

What you will see
- A black console window with timestamped messages and aligned tables for diffs.
- Startup prints:
  - Supported compression formats and chosen default
  - Watched roots, monitor-only roots, and supported extensions
  - Scan summary and baseline creation progress (if enabled)

Normal operations
- When an Excel file changes under WATCH_FOLDERS:
  - A silent preview compare runs; if there are changes, a banner appears and a table is printed.
  - If AUTO_UPDATE_BASELINE_AFTER_COMPARE is enabled, the baseline is updated after the first visible compare.
- Under MONITOR_ONLY_FOLDERS:
  - On first detected change: we log info and create a baseline (no compare yet). Subsequent changes will compare against this baseline.
- Adaptive polling:
  - After a change, the system checks periodically (by file size-driven interval). When changes are detected, it prints a new before/after table and immediately updates the baseline so the next table compares against the latest state. Once mtime stabilizes and no diff remains, polling stops.

Best practices to avoid file locks
- Let the tool copy files to a local CACHE_FOLDER before reading (default behavior). Do not place CACHE_FOLDER under a watched root.
- If a file is actively being saved by Excel, copies may fail briefly; the tool retries automatically with backoff (default BACKOFF=1.0s). You can increase retries in Settings.
- The tool reads last modified author from the cached .xlsx by parsing core properties (docProps/core.xml). It does not need to open the original.
- All workbook reads use read_only mode and close resources immediately.
- In STRICT_NO_ORIGINAL_READ mode, if copying ultimately fails, the file is skipped (original file is never opened).

Console text log
- You can choose to write only change-related messages to the text file (banners and comparison tables). In Settings UI, enable "將 Console 輸出寫入文字檔" and optionally enable "Console 文字檔僅記錄「變更」".
- All console messages are appended to a UTF-8 text file when enabled. Configure path in "Console 文字檔路徑".

Daily ops log
- Copy failures are recorded to LOG_FOLDER/ops_log/copy_failures_YYYYMMDD.csv with timestamp, file path, error, attempts, and current strict/chunk/backoff settings.
- Review daily to spot repeated failures (network issues, Excel save conflicts, or access problems).

Troubleshooting
- No events observed:
  - Confirm WATCH_FOLDERS paths are correct; ensure SCAN_TARGET_FOLDERS is a subset of watched roots.
  - Check exclusions: WATCH_EXCLUDE_FOLDERS and MONITOR_ONLY_EXCLUDE_FOLDERS.
- Cache noise (too many events):
  - Move CACHE_FOLDER outside of watched roots or enable an ignore rule (if available in your build).
- CSV logs not produced:
  - Only meaningful changes are logged; ensure TRACK_* settings match your needs.
- Memory/Timeout alerts:
  - The tool will GC and throttle. Reduce watched scope or adjust limits in Settings UI.

Operational checklists
- Daily start:
  - Confirm settings (paths, formats) and that the console opens.
  - Verify that baseline creation (if enabled) completes without errors.
- During changes:
  - Confirm that the banner and diff table show up when editing a watched file.
  - For monitor-only roots, expect the first change to only create a baseline and print info.
- Prior to shutdown:
  - Use Ctrl+C once to request graceful stop; press again to force.

FAQ (short)
- Q: Will the tool lock my Excel files?
  A: It operates on cached copies for heavy reads and parses the author from the cached ZIP’s core properties. It does not need to open your original file after copying.
- Q: Can I watch entire drives?
  A: Yes via MONITOR_ONLY_FOLDERS for large roots; define exclusions to avoid system/hidden directories.
- Q: What file types are supported?
  A: .xlsx and .xlsm by default; others are out of scope for now.

End of runbook.

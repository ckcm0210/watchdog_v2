# Excel Watchdog Project Summary
Generated at: 2025-08-27 00:00 UTC

This document provides a detailed architecture map, end-to-end flow (monitor → compare → baseline update), module relationships, and operational guardrails (with a focus on avoiding Excel file locks). It incorporates findings up to 2025-08-27.

1) System overview
- Purpose: Monitor folders (including network shares) for Excel file changes (.xlsx/.xlsm). Build baseline snapshots of cell values/formulas, detect differences, display aligned diffs in console/UI, and log meaningful changes.
- Key capabilities:
  - File monitoring with debounce (watchdog) and adaptive polling
  - Excel parsing of cell values/formulas (openpyxl, read-only)
  - Difference computation and classification (formula change, direct value change, external ref update)
  - Compressed baseline management (lz4/zstd/gzip) and migration
  - Local caching of source files to avoid locking originals
  - Progress save/resume, timeout protection, memory monitoring
  - Black console UI for real-time, CJK-safe output

2) Detailed architecture map (modules and roles)

```
+----------------------+        +--------------------+        +------------------+
|      UI / Console    |<------>|   utils.logging    |<------>|  core.comparison |
|  (ui.console)        |        | (print hook + CJK) |        |  (diff & classify)
+----------^-----------+        +--------------------+        +---------^--------+
           |                                                          |
           | messages (timestamped)                                     | loads baseline / prints table
           |                                                          |
+----------+-----------+        +--------------------+        +---------+--------+
|      main.py         |------->|  watchdog Observer |------->| core.watcher    |
| - init UI/logging    |        | (filesystem events)|        |  (event handler)|
| - startup scan       |        +--------------------+        +---------+--------+
| - schedule watchers  |                                                  |
+----------+-----------+                                                  |
           |                                                              | compare / baseline
           v                                                              v
+----------------------+        +--------------------+        +------------------+
|  core.baseline       |<------>| utils.compression  |        | core.excel_parser|
| - save/load/migrate  |        | (gzip/lz4/zstd)    |        | - dump cells     |
| - batch create       |        +--------------------+        | - last author    |
| - archive old        |                                        | - extract refs   |
+----------^-----------+                                        +---------^--------+
           |                                                                |
           | base_key from path hash                                        | uses cached files
           |                                                                |
           v                                                                v
+----------------------+                                        +------------------+
|  utils.helpers       |                                        |  utils.cache     |
| - baseline key       |                                        | - copy_to_cache  |
| - mtime/size         |                                        | - safe filenames |
| - progress save/load |                                        +------------------+
+----------------------+
```

3) End-to-end flow (monitor → compare → baseline update)
- Startup
  1. main.py initializes logging (timestamped print), optional Settings UI, timeout/memory monitoring, and console.
  2. If SCAN_ALL_MODE=True, collects Excel files (WATCH_FOLDERS or SCAN_TARGET_FOLDERS) and calls core.baseline.create_baseline_for_files_robust.
  3. Schedules both WATCH_FOLDERS and MONITOR_ONLY_FOLDERS in the observer (WATCH_FOLDERS takes priority over monitor-only).

- On file created
  4. core.watcher.ExcelFileEventHandler.on_created: If supported Excel and not temp (~$), triggers baseline creation for that file.

- On file modified
  5. Debounce by path; ignore ~$. Preview compare silently to avoid noise; if changes detected, print banner.
  6. If under MONITOR_ONLY_FOLDERS and no baseline exists yet: log info (mtime/author), create first baseline, and return (no compare yet). Next changes will compare.
  7. Otherwise, run compare_excel_changes(file):
     - core.excel_parser.dump_excel_cells_with_timeout() reads from a cached copy, in read_only mode; extracts ref_map once per workbook; prettifies formulas.
     - core.baseline.load_baseline(base_key) loads compressed baseline.
     - Differences are rendered with CJK-aware aligned columns; meaningful changes can be logged to CSV.
  8. If changes and AUTO_UPDATE_BASELINE_AFTER_COMPARE=True: save updated baseline (compressed), with new timestamp/author.
  9. ActivePollingHandler starts adaptive polling by file size, using mtime stability to decide when to stop.

4) File-lock avoidance strategy (current and recommended)
- Always operate on cached copies
  - utils.cache.copy_to_cache() creates a local copy (safe filename, freshness check). All heavy reads use the cached file, not the original.
- Read-only Excel access and explicit release
  - openpyxl.load_workbook(..., read_only=True) to minimize resource usage.
  - Ensure wb.close() and delete references (del wb) in finally blocks. Current code does this in both happy and exception paths.
- Last author retrieval without opening the original
  - Prefer parsing docProps/core.xml from the cached .xlsx via zipfile/ElementTree (no openpyxl, no original file lock).
  - If ZIP parse fails (corrupt/legacy), fall back to reading the cached copy with openpyxl, still avoiding the original.
- Timeout and in-flight markers
  - settings.current_processing_file and processing_start_time guard a per-file timeout; on timeout, markers are cleared to avoid dangling state.
- Periodic cleanup
  - After each file in batch baseline: perform gc.collect(), drop large dicts (del cell_data), and release old_baseline references.
  - Memory checks via utils.memory.enable + get_memory_usage() to throttle when high.

Recommended operational hygiene (no code required to adopt as practice)
- Do not place CACHE_FOLDER within watched roots; or configure an ignore rule for the cache path.
- Avoid falling back to reading the original file if caching fails; prefer retrying the copy (operational runbook decision).
- If needed, schedule a periodic “janitor” step every N files/events: gc.collect(), clear temp variables, and verify no lingering wb/ZipFile handles.

5) Architecture responsibilities (who does what)
- main.py: orchestrates startup, scheduling, and graceful shutdown; prints capability info (compression formats and settings).
- core.watcher: debounced event handling, preview compare to reduce noise, monitor-only first-seen handling, and adaptive polling by mtime.
- core.comparison: loads baseline, generates aligned diff tables, classifies changes (FORMULA_CHANGE, DIRECT_VALUE_CHANGE, EXTERNAL_REF_UPDATE, INDIRECT_CHANGE), and logs meaningful ones.
- core.baseline: saves/loads compressed baselines; batch baseline build with resume/timeout/memory protection; archives old baselines.
- core.excel_parser: safe workbook dump; ref_map extraction; pretty_formula normalization; last author via core.xml.
- utils.cache: local copy creation for safe reading; avoids original locks; safe filenames; freshness checks.
- utils.compression: compression abstraction (gzip/lz4/zstd), stats, migrations.
- utils.helpers: baseline key generation from path hash; mtime formatting; resume progress I/O; timeout thread.
- ui.console + utils.logging: timestamped print hooked to black console with CJK width handling.

6) Detailed flow diagram (text)

```
[File event] --(watchdog)--> ExcelFileEventHandler.on_modified
  -> debounce/filter temp files
  -> silent preview compare (no banner if no diff)
  -> if MONITOR_ONLY and no baseline: log info + create baseline + return
  -> compare_excel_changes(file)
      -> dump_excel_cells_with_timeout(file)
          -> path' := copy_to_cache(file)
          -> wb := load_workbook(path', read_only=True)
          -> ref_map := extract_external_refs(path') once
          -> iterate cells: {formula:=pretty_formula(f, ref_map), value:=serialize(v)}
          -> wb.close(); del wb
      -> old := load_baseline(base_key)
      -> diff := render aligned CJK table; classify changes
      -> if !is_polling and AUTO_UPDATE_BASELINE: save_baseline(base_key, current_data)
  -> ActivePollingHandler.start_polling(file): use mtime stability to stop polling
```

7) Update log (2025-08-27)
- Config defaults (more conservative): COPY_RETRY_COUNT=8, BACKOFF=1.0s, CHUNK=4MB; daily ops log for copy failures.
- Documentation: Added detailed architecture map, end-to-end flow, and file-lock avoidance/cleanup practices.
- Excel last author: documented a non-locking approach (parse core.xml from cached copy first; fallback to reading cached workbook), matching the tool’s overall strategy of operating on cached files.
- Next steps (doc-level recommendations):
  - Consider adding a runtime toggle to avoid falling back to original file reads if caching fails.
  - Consider an ignore rule for CACHE_FOLDER in watcher configuration to eliminate cache-noise events.
  - Optionally add a periodic janitor step in long-running sessions.

8) Runbook checklist (ops)
- Verify Settings UI values applied; SUPPORTED_EXTS shows .xlsx/.xlsm.
- Ensure WATCH_FOLDERS and MONITOR_ONLY_FOLDERS are scheduled; exclusions applied.
- Confirm baseline creation and comparison flows work end-to-end on a sample set (with external links).
- Watch for cache placement under watched roots; move out or ignore if noisy.
- Observe memory and timeout prints; adjust limits if needed.

9) Detailed call flow and data structures

9.1 Call flow (startup baseline)
```
main.main()
  -> ui.settings_ui.show_settings_ui() [optional]
  -> (if SCAN_ALL_MODE) utils.helpers.get_all_excel_files()
  -> core.baseline.create_baseline_for_files_robust(files)
       for each file F:
         - base_key := utils.helpers._baseline_key_for_path(F)
         - core.excel_parser.dump_excel_cells_with_timeout(F)
             - F' := utils.cache.copy_to_cache(F)
             - wb := openpyxl.load_workbook(F', read_only=True)
             - ref_map := extract_external_refs(F')
             - iterate ws/rows/cells -> {formula:=pretty_formula(f, ref_map), value:=serialize(v)}
             - wb.close(); del wb
         - hash := hash_excel_content(cells)
         - last_author := get_excel_last_author(F)  // uses cached copy and core.xml
         - core.baseline.save_baseline(base_key, {last_author, content_hash, cells, timestamp})
```

9.2 Call flow (modify event → compare → optional baseline update)
```
watchdog -> core.watcher.ExcelFileEventHandler.on_modified(event)
  -> debounce + skip '~$'
  -> preview := core.comparison.compare_excel_changes(F, silent=True)
  -> if preview: print change banner (file, event #, author)
  -> if monitor-only and no baseline: log info + create baseline + return
  -> has_changes := core.comparison.compare_excel_changes(F, silent=False)
       - base_key := _baseline_key_for_path(F)
       - current := dump_excel_cells_with_timeout(F)  // cached copy
       - baseline := core.baseline.load_baseline(base_key)
       - per-sheet compare; print aligned diff
       - meaningful := analyze_meaningful_changes(...); if not polling -> CSV log
       - if has_changes and AUTO_UPDATE_BASELINE_AFTER_COMPARE and not polling:
           core.baseline.save_baseline(base_key, updated_baseline)
  -> ActivePollingHandler.start_polling(F)
       - _poll_for_stability: check mtime; if changed -> compare(is_polling=True)
       - stop when mtime stable
```

9.3 Baseline JSON structure (compressed on disk as .baseline.json.{lz4|zst|gz})
```json
{
  "last_author": "USER01",
  "content_hash": "2b1a0c0e3f...",
  "timestamp": "2025-08-27T10:23:45",
  "cells": {
    "Sheet1": {
      "A1": {"formula": "=SUM(B1:B10)", "value": null},
      "B2": {"formula": null, "value": 123.0},
      "C3": {"formula": "='\\\\server\\share\\file.xlsx'!Ref!A1", "value": 456}
    },
    "Links": {
      "A1": {"formula": "='C:\\data\\source.xlsx'!Ref!A1", "value": 42}
    }
  }
}
```

9.4 Pretty formula normalization example
- Raw: `=[1]SheetX!A1`
- With ref_map: `[1] -> file:///C:/data/source.xlsx`
- Normalized pretty: `='C:\\data\\source.xlsx'!SheetX!A1`

9.5 Meaningful change record (CSV row fields)
```csv
Timestamp,Filename,Worksheet,Cell,Change_Type,Old_Value,New_Value,Old_Formula,New_Formula,Last_Author
2025-08-27 10:25:10,Book1.xlsx,Sheet1,A1,EXTERNAL_REF_UPDATE,,42,="='C:\\data\\source.xlsx'!Ref!A1",="='C:\\data\\source.xlsx'!Ref!A1",USER01
```

9.6 Diff printing inputs
- display_old: dict[address -> cell_repr] only for differing addresses
- display_new: dict[address -> cell_repr] only for differing addresses
- Each cell_repr is `{formula?: string, value?: scalar}`; printing layer formats as `=...` for formulas, `repr(value)` for values, marks `[MOD]/[ADD]/[DEL]`.

End of document.

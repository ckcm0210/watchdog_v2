# Excel Watchdog Project Summary
Generated at: 2025-08-26 00:00 UTC

This document summarizes the current codebase, provides per-file roles, counts (lines/defs/classes), a dependency map, and recommendations. It also includes proposals addressing five requested enhancements.

1) System overview
- Purpose: Monitor folders (including network drives) for Excel file changes (.xlsx/.xlsm). Create baseline snapshots of cell values/formulas, detect differences on change, display aligned diffs in console/UI, and log meaningful changes.
- Key capabilities:
  - File monitoring with debounce (watchdog)
  - Excel parsing of cell values/formulas (openpyxl)
  - Difference computation and classification (formula change, direct value change, external ref update, etc.)
  - Compression management of baselines (lz4/zstd/gzip) and migration
  - Local caching for network files
  - Progress save/resume, timeout protection, memory monitoring
  - Black console UI (Tkinter) for real-time output

2) Per-file summary, stats, and imports
Note: Counts were measured via an analyzer; numbers are approximate and for orientation.

- main.py
  - Role: Entry point. Initializes logging/console/timeout, checks compression, scans Excel, builds baselines, starts watchdog, handles graceful shutdown.
  - Stats: 143 lines, 2 defs, 0 classes
  - Imports: config.settings, utils.logging, utils.memory, utils.helpers, utils.compression, ui.console, core.baseline, core.watcher, core.comparison, watchdog.observers, etc.

- config/settings.py
  - Role: Central runtime configuration (folders, formats, cache paths, timeouts, options).
  - Stats: 81 lines

- core/baseline.py
  - Role: Load/save baseline (compressed), archive/migrate formats, batch baseline creation (with resume/timeout/memory safeguards and caching).
  - Key funcs: load_baseline, save_baseline, get_baseline_file_with_extension, create_baseline_for_files_robust, archive_old_baselines.
  - Stats: 314 lines, 6 defs
  - Imports: utils.compression, utils.helpers, utils.memory, core.excel_parser, config.settings.

- core/comparison.py
  - Role: Compare current Excel content vs baseline; aligned CJK-safe rendering; classify meaningful changes; write compressed CSV logs; optional auto-update baseline.
  - Key funcs: print_aligned_console_diff, compare_excel_changes, analyze_meaningful_changes, classify_change_type, has_external_reference, log_meaningful_changes_to_csv.
  - Stats: 352 lines, 11 defs
  - Imports: core.baseline, core.excel_parser, utils.helpers, utils.logging._get_display_width, config.settings, gzip/csv, etc.

- core/excel_parser.py
  - Role: Safe-loading Excel; dump per-sheet cell values/formulas (array formula aware) to dict; compute hash; extract external refs via ZIP/XML; get lastModifiedBy.
  - Key funcs: dump_excel_cells_with_timeout, get_excel_last_author, pretty_formula, get_cell_formula, serialize_cell_value, hash_excel_content, extract_external_refs, safe_load_workbook.
  - Stats: 218 lines, 9 defs
  - Imports: openpyxl, utils.cache, config.settings, zipfile/XML, etc.

- core/watcher.py
  - Role: Watchdog handlers for create/modify. Adaptive polling by file size; avoids repeated immediate checks; debounce; triggers compare/baseline.
  - Classes/funcs: ActivePollingHandler (start_polling/_poll_for_stability/stop), ExcelFileEventHandler (on_created/on_modified).
  - Stats: 181 lines, 10 defs, 1 class

- core/watcher-Copy1.py
  - Role: Legacy watcher (dense/sparse poll variants). Not used by main. Consider moving to legacy/ or removing.
  - Stats: 204 lines, 13 defs, 1 class

- ui/console.py
  - Role: Tkinter black console window; message queue; popup on comparison; short topmost; runs in separate thread.
  - Key: BlackConsoleWindow (create_window, popup_window, check_messages, add_message, ...), init_console.
  - Stats: 202 lines, 15 defs

- utils/cache.py
  - Role: Copy network file to local cache with freshness check and progress output; robust error handling.
  - Stats: 55 lines, 1 def

- utils/compression.py
  - Role: Unified compression/IO across gzip, lz4, zstd; detection, validation, stats, migration, test support output.
  - Key funcs: compress_data, decompress_data, save_compressed_file, load_compressed_file, get_compression_stats, migrate_baseline_format, test_compression_support.
  - Stats: 345 lines, 11 defs

- utils/helpers.py
  - Role: File mtime formatting, human size, scanning Excel files, forced baseline detection, progress save/load, timeout handler thread.
  - Stats: 117 lines, 7 defs

- utils/logging.py
  - Role: Hook builtins.print to add timestamps and forward to black console; CJK width handling.
  - Stats: 86 lines, 4 defs

- utils/memory.py
  - Role: Process memory usage and limit enforcement with GC.
  - Stats: 41 lines, 2 defs

3) High-level dependency map
- main → config.settings, utils.logging, ui.console, utils.helpers, utils.compression, core.baseline, core.watcher (Observer)
- core.watcher → core.baseline (on_created), core.comparison (compare), core.excel_parser (last author)
- core.comparison → core.baseline (load), core.excel_parser (pretty/refs/author), utils.helpers (mtime), utils.logging (width)
- core.baseline → utils.compression, core.excel_parser (dump/hash/author), utils.memory, utils.helpers
- core.excel_parser → utils.cache, openpyxl, zipfile/XML
- ui/console ↔ utils.logging (print hook forwards messages)

4) Cleanup and refactor suggestions
- Move core/watcher-Copy1.py to legacy/ or remove; clean unused imports in main.py (set_current_event_number/check_memory_limit not used there).
- If FORMULA_ONLY_MODE/WHITELIST_USERS intended to affect logic, wire them into comparison filters.
- Convert compression module startup prints to logging and/or gated by a debug flag.

5) Roadmap for requested enhancements

Enhancement 1: Convert config/settings to a startup UI for parameters
- Approach A (quick win):
  - Create a Tkinter settings dialog on startup (reusing UI stack) to gather key parameters: WATCH_FOLDERS (multi-select), DEFAULT_COMPRESSION_FORMAT (lz4/zstd/gzip), USE_LOCAL_CACHE/CACHE_FOLDER, ENABLE_TIMEOUT/FILE_TIMEOUT_SECONDS, ENABLE_MEMORY_MONITOR/MEMORY_LIMIT_MB, POLLING_* values, SCAN_ALL_MODE, AUTO_UPDATE_BASELINE_AFTER_COMPARE, etc.
  - Save to config/runtime_settings.json and apply to process by assigning into config.settings (module attributes can be changed at runtime).
- Approach B (better architecture):
  - Add a thread-safe ConfigManager (singleton/global) with from_json/to_json, get/set, and change callbacks. Consumers read via ConfigManager instead of reading config.settings directly, enabling live updates.
- Difficulty: Low-to-Medium for A, Medium for B.
- Estimated effort: A ~0.5–1.5 days; B ~2–4 days (including incremental adoption and tests).

Enhancement 2: Monitor entire (large) network drive without pre-creating baselines; on change, fetch last modified and last author; optionally baseline on-demand
- Plan:
  - Run watchdog recursively on the drive root, but set SCAN_ALL_MODE=False and skip initial baseline creation.
  - Modify ExcelFileEventHandler to support a "monitor-only" mode: on_created/on_modified, do not automatic baseline; instead:
    - Record event with path, file size, get_file_mtime(file_path), and get_excel_last_author(file_path) (openpyxl wb.properties.lastModifiedBy).
    - Optionally, on-demand baseline creation per file (e.g., a flag, command, or threshold-based rule).
  - Consider filtering by SUPPORTED_EXTS to reduce load.
- Caveats: Determining the Windows/SMB "last modified by" at filesystem level is non-trivial. Using Excel lastModifiedBy is reliable if file is saved by Excel. For non-Excel files, see Enhancement 5.
- Difficulty: Low (changes are localized to watcher + main flags)
- Effort: ~0.5 day.

Enhancement 3: Normalize external reference paths in baseline (replace [1]/[2] placeholders, clean %20, and UNC/backslash normalization)
- Current: extract_external_refs() can map [n] to paths; pretty_formula() can rewrite formula text with source path hint.
- Proposal:
  - During dump_excel_cells_with_timeout, once per workbook, call extract_external_refs(xlsx_path) and pass its ref_map to pretty_formula for every formula encountered.
  - Decode percent-encoded paths using urllib.parse.unquote and normalize to a consistent UNC/backslash style (e.g., \\
etwork\share\path) via os.path.normpath; consider storing both raw_formula and pretty_formula, or replace formula with the prettified version for readability.
  - Ensure comparison uses prettified formulas consistently in both baseline and current to avoid false diffs.
- Difficulty: Medium (need careful integration at serialization); Effort: ~0.5–1 day.

Enhancement 4: Change detection coverage for (a) direct values, (b) indirect values with unchanged formula, (c) refresh-only external links
- Current behavior:
  - DIRECT_VALUE_CHANGE: when no formula and value changes.
  - EXTERNAL_REF_UPDATE: when formula unchanged and values change, and has_external_reference(formula) is true (heuristic).
  - INDIRECT_CHANGE: when formula unchanged, values change, and no explicit external ref signature.
- Gaps and improvements:
  - Strengthen has_external_reference by using real token patterns or leveraging extract_external_refs mapping; detect external refs reliably beyond simple "['" heuristics.
  - Optionally log external dependency file mtimes alongside change to indicate probable cause (refreshed source).
  - Provide toggles to include/exclude INDIRECT_CHANGE and EXTERNAL_REF_UPDATE via settings/UI.
- Difficulty: Medium; Effort: ~0.5–1.5 days.

Enhancement 5: Extend monitoring to other file types (e.g., Word) capturing last modified/user without before/after content
- Plan:
  - Generalize SUPPORTED_EXTS and add a GenericFileEventHandler branch for non-Excel.
  - For .docx/.pptx: parse ZIP docProps/core.xml to get lastModifiedBy (similar approach as Excel via OpenXML core properties).
  - For arbitrary files: capture path, event time, size, and mtime; "last author" may not be universally available.
  - Output to CSV log (new lightweight schema) with type field.
- Difficulty: Medium (docx support straightforward; generic files limited metadata); Effort: ~1–2 days.

6) Suggested incremental plan
- Phase 1: Settings dialog (Approach A) + monitor-only mode for Excel + improved external ref detection (baseline-time prettify). Clean legacy watcher.
- Phase 2: ConfigManager for live updates of low-risk params; ActivePollingHandler apply-config; optional drive-wide watcher update function.
- Phase 3: Extend to Word/docx and generic non-Excel metadata logging; refine change classification and reporting.

7) Notes on runtime live updates
- Immediately adjustable (low risk): display limits, debug flags, compression format for future saves, cache usage (applies to next reads).
- Requires rescheduling (medium): polling intervals/debounce, timeout/memory monitor threads, UI on/off.
- Requires re-watching (higher): WATCH_FOLDERS changes (unschedule/schedule on the fly).

8) Quick wins
- Add monitor-only mode flag in settings and adjust watcher to skip baseline until on-demand.
- Apply extract_external_refs + unquote/normalize paths in formula serialization.
- Improve external ref detection logic in classify_change_type with ref_map awareness.
- Move legacy watcher-Copy1.py to legacy/ and remove unused imports.

9) Detailed note on change classification (刷新連結但公式不變)
- 需求目標：只記錄「有意義」變更：
  - 直接值（非公式）變更 → 必須記錄。
  - 公式字串變更 → 必須記錄。
  - 公式不變、結果因工作簿內其他儲存格變更而改變 → 預設忽略（IGNORE_INDIRECT_CHANGES=True）。
  - 公式不變、結果因外部連結刷新 → 記錄為 EXTERNAL_REF_UPDATE。
- 改善：
  - 使用 extract_external_refs(file) 取得 ref_map，強化 has_external_reference 判斷（不僅靠字串）。
  - 新增 DISPLAY_ONLY_MEANINGFUL_CHANGES：畫面/CSV 僅顯示/記錄 meaningful_changes（避免非重點變更噪音）。
  - 可選：對外部來源做 mtime 對照以提升準確性。
- 設定建議：
  - TRACK_DIRECT_VALUE_CHANGES=True, TRACK_FORMULA_CHANGES=True, TRACK_EXTERNAL_REFERENCES=True,
  - IGNORE_INDIRECT_CHANGES=True, DISPLAY_ONLY_MEANINGFUL_CHANGES=True（新增）。

End of document.

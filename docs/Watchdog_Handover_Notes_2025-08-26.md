# Watchdog Handover Notes (2025-08-26)

Note: See also Watchdog_Handover_Notes_2025-08-27.md for the latest handover with additional guidance on avoiding Excel file locks and detailed flow diagrams.

Purpose: Persist the working context so we can resume quickly tomorrow.

Summary of what changed today
- Settings UI: near 1:1 mapping of config.settings with clear help, priorities, path pickers, automatic fill from defaults + runtime JSON, wrap long help. New items:
  - MONITOR_ONLY_FOLDERS (+ MONITOR_ONLY_EXCLUDE_FOLDERS)
  - WATCH_EXCLUDE_FOLDERS
  - SCAN_TARGET_FOLDERS (subset to scan on startup)
  - AUTO_SYNC_SCAN_TARGETS (auto-copy WATCH_FOLDERS to SCAN_TARGET_FOLDERS on UI load)
- Runtime coercion: SUPPORTED_EXTS normalization; accept comma/semicolon; ignore empty overrides. Fixed JSON errors.
- Scheduling: Now schedules both WATCH_FOLDERS and MONITOR_ONLY_FOLDERS. Priority: WATCH_FOLDERS overrides monitor-only.
- Baseline identity: Use base_key = original_filename + "__" + path_sha1_8; but console shows original filename only.
- Cache filenames: sanitize invalid chars, cap length, avoid repeated prefixes.
- Polling: stability detection via mtime; reduce false "still changing".
- Watcher: silent preview compare before printing change banner to reduce noise.
- Progress save: RESUME_LOG_FILE fallback to LOG_FOLDER/resume_log if invalid; fixed exceptions.

Key files touched
- ui/settings_ui.py: big update for UI; new spec entries; path and paths controls; AUTO_SYNC_SCAN_TARGETS; help wraps.
- main.py: schedules both WATCH_FOLDERS + MONITOR_ONLY_FOLDERS; supports SCAN_TARGET_FOLDERS for startup scan; prints monitor-only list.
- core/watcher.py: priority resolution, monitor-only first-seen, mtime-based polling, silent preview compare.
- core/baseline.py: prints display_name (no hash); base_key for storage; compression stats.
- core/comparison.py: base_key usage; removed legacy base_name in prints; errors fixed.
- utils/cache.py: safe cache filename builder; avoid double-prefix; length cap.
- utils/helpers.py: _baseline_key_for_path cleans stacked md5 prefixes; truncates; appends hash8; save_progress fallback.
- config/runtime.py: robust list/tuple coercion; SUPPORTED_EXTS clean.
- config/settings.py: new settings (monitor-only and exclusions; scan subset).
- docs/Watchdog_Project_Summary_2025-08-26.md: extended with Update log.

Open items (ordered)
1) External reference normalization in formulas (priority: high)
   - Integrate extract_external_refs into dump/pretty; normalize with unquote + os.path.normpath; ensure both baseline and current normalized.
2) Change classification refinement (priority: medium)
   - Use ref_map to detect EXTERNAL_REF_UPDATE reliably; optional CSV logging for monitor-only first-seen events.
3) UI live-sync: WATCH_FOLDERS → SCAN_TARGET_FOLDERS (merge mode) when AUTO_SYNC_SCAN_TARGETS enabled
   - Current: sync on UI load only; TODO: intercept add/remove/clear buttons to maintain SCAN_TARGET_FOLDERS.
4) Optional: ignore events under CACHE_FOLDER entirely (global checkbox), to eliminate cache noise.
5) Optional: Manual "Sync from WATCH_FOLDERS" button for SCAN_TARGET_FOLDERS (if not using auto sync).

Risks / gotchas
- If CACHE_FOLDER is under a watched root, cache events may appear; consider ignoring.
- SCAN_TARGET_FOLDERS should be subset of WATCH_FOLDERS; otherwise startup builds baselines for paths that are not observed.
- MONITOR_ONLY_FOLDERS for large roots needs MONITOR_ONLY_EXCLUDE_FOLDERS to avoid system directories.
- Windows path length: we capped names but avoid extremely deep LOG_FOLDER.

What to verify on resume
- Settings UI opens; no empty path fields; SUPPORTED_EXTS prints ['.xlsx', '.xlsm']; startup scan finds expected files.
- Events: WATCH_FOLDERS immediate compare; MONITOR_ONLY_FOLDERS first-seen logs info + creates baseline, second change compares.
- Exclusions filtering works for both modes.
- Polling relies on mtime and stabilizes properly—no phantom "still changing".

Decision points awaiting user confirmation
- Live-sync behavior: Hard sync vs Merge sync for SCAN_TARGET_FOLDERS when WATCH_FOLDERS changes (user prefers merge sync).
- Exclude CACHE_FOLDER by default?
- CSV logging for monitor-only first-seen events; define schema.
- External reference normalization details; sample files.

Next concrete steps checklist
- [ ] Implement external reference normalization in core/excel_parser.py (dump + pretty_formula) and ensure compare uses normalized values.
- [ ] Implement merge sync for SCAN_TARGET_FOLDERS.
- [ ] Add optional "Ignore cache folder" setting and apply in watcher scheduling.
- [ ] Add manual "Sync from WATCH_FOLDERS" to SCAN_TARGET_FOLDERS block.
- [ ] (Optional) Add CSV logging for monitor-only first-seen events.

How to reload this context tomorrow
- Read this note and the Update log in docs/Watchdog_Project_Summary_2025-08-26.md.
- Open config/settings.py + ui/settings_ui.py to see the new fields.
- If user reported issues: check console logs around watcher -> comparison flow; focus on mtime polling behavior.

Contact points / quick references
- Baseline key logic: utils/helpers.py::_baseline_key_for_path
- Monitor-only logic: core/watcher.py::_is_monitor_only and on_modified
- Startup scan subset: main.py (SCAN_TARGET_FOLDERS handling)
- Runtime coercion: config/runtime.py::_coerce_type


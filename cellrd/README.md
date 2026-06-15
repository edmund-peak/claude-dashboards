# CellRD Lab Manager — File-Watcher Edition

Single-program version of CellRD Lab Manager v6 that **auto-populates 12 of the 14 chambers from Neware filenames** in one or more watched folders, while keeping the entire RPT Tracker tab and the remaining two chambers user-editable.

## What changed vs v6

| | v6 (artifact) | This version |
|---|---|---|
| Storage | `window.storage` (browser) | `data.json` on the server |
| Cell IDs in Chamber_1–9, HighTemp, CoinCell_1, CoinCell_2 | manually edited | **auto-detected from Neware filenames** |
| Metadata on those chambers (analysis group, format, duration) | manually edited | still manually edited (preserved across watcher updates) |
| UHPC, Bio-logic | manually edited | manually edited (unchanged) |
| RPT Tracker tab | manually edited | manually edited (unchanged) |
| Source machines | n/a | **multiple watch folders** (local + network shares) |
| Multi-user | one browser at a time | shared across everyone on the network |

## Auto-sync vs manual

**Auto-sync chambers** (driven by file watcher):
`Chamber_1`, `Chamber_2`, `Chamber_3`, `Chamber_4`, `Chamber_5`, `Chamber_6`,
`Chamber_7`, `Chamber_8`, `Chamber_9`, `HighTemp`, `CoinCell_1`, `CoinCell_2`

**Manual chambers** (fully user-editable like v6):
`UHPC`, `Bio-logic`

The **RPT Tracker tab** stays fully manual regardless — even though Chamber_9 cells in the Scheduler view get auto-detected, the Tracker tab still relies on you to record RPT events, assignments, and storage SOC.

## Quick start

```bash
pip install -r requirements.txt
python cellrd_app.py
```

Then open **http://localhost:5000** — or, from another machine on the same network, replace `localhost` with the host computer's IP. Everyone sees the same data.

On first run the app builds `data.json` (the lab state) and `metadata.json` (your analysis-group/format/duration assignments for the auto-sync chambers) from the seed baked into the app.

## How auto-sync works

The 6 Standard Testing chambers (`Chamber_1` through `Chamber_6`) are now driven by file scanning:

1. Every 30 seconds (configurable), the app scans every folder in `watch_folders` for `.xlsx` files matching the Neware naming pattern
2. It finds the most recently modified file across all folders
3. **Gap-based batch detection**: walks backwards through file mtimes and stops at the first gap larger than `batch_gap_minutes` (default 120). Everything before that gap is treated as a previous session (cells already pulled). `active_window_hours` is a hard outer cap that prevents stale files from sneaking in.
4. For each channel, the most recent file in the active batch wins
5. The parsed `cellId`, test type, and start date (from file creation time) overwrite those fields on the matching channel

For these 6 chambers, `cellId`, `currentTest`, and `startDate` are **read-only** in the UI — they're driven by what's actually running. Chambers carry an `⟲ Auto` badge in the header so it's obvious. The Edit button is renamed "Meta" and the locked fields are grayed out in the modal with a banner explaining what's editable.

`analysisGroup`, `cellFormat`, `durationWeeks`, and `purpose` stay editable on auto-sync chambers — they're stored separately in `metadata.json` keyed by `channelId` and re-applied after every scan. Set them once and they survive forever.

The header shows a small live pill like `⟲ 23 live · 12s` — that's the watcher status. Click it to force an immediate rescan.

## Setup

### 1. Edit `config.json`

```json
{
  "watch_folders": [
    "C:\\ProgramData\\Neware\\Data",
    "\\\\PE-HZ83664\\Neware Data"
  ],
  "poll_interval_seconds": 30,
  "active_window_hours": 24,
  "rack_mapping": {
    "3":   "R03",
    "4":   "R04",
    "5":   "R05",
    "165": "R02",
    "166": "R01"
  }
}
```

**`watch_folders`** — a list of one or more folders to scan. Each folder is scanned independently every poll cycle; files from all folders are pooled and routed to the right chamber by `channelId`. Use a single-element list if you only have one source machine. Network paths work — just remember every `\` needs to be doubled to `\\` in JSON, so `\\PE-HZ83664\Neware Data` becomes `"\\\\PE-HZ83664\\Neware Data"`. If a folder is unreachable (machine off, share missing), the watcher logs a warning and keeps going with whatever folders *are* reachable.

(Back-compat: the old `"watch_folder": "single\\path"` form still works if you have a leftover config.)

**`rack_mapping`** — maps the Neware rack number (the number that appears between the `BTSnn` segment and the tester number in the filename) to the rack code used in your scheduler (`R01`–`R05`). This is a single global mapping that covers all watch folders — so as long as rack numbers don't collide across machines, one map handles everything.

> Look at one of your filenames: `CRD_12_CID018_..._127.0.0.1-BTS85-`**`166`**`-10-8-399_2.xlsx`
> The bold number is the Neware rack. The next number (`10`) is the tester, and `8` is the channel.

**Filename formats supported:**
- Original: `CRD_XX_CIDYYY_TestInfo_127.0.0.1-BTSnn-Rack-Tester-Channel-Seq.xlsx`
- Newer (with timestamp prefix and underscore IP): `<timestamp>_CRD_XX_CIDYYY_TestInfo_127_0_0_1-BTSnn-Rack-Tester-Channel-Seq.xlsx`

The list of channels per chamber is **not** in `config.json` — it comes from the chamber/channel grid baked into the app (lifted directly from your v6 schedule). If the watcher sees a `channelId` that doesn't match any chamber, that file is reported under "unmapped" in `/api/watcher` but doesn't break anything.

**`poll_interval_seconds`** — how often the watcher rescans the folders.

**`active_window_hours`** — outer hard cap. Even with the batch gap logic below, no file older than this many hours from the newest is ever included.

**`batch_gap_minutes`** — gap-based batch detection. The watcher finds the newest file across all folders, walks backwards through mtimes, and stops when it hits a gap larger than this many minutes between consecutive files. Everything before the gap is treated as a previous batch (cells already pulled). Default: 120 minutes.

Example: yesterday at 5pm you pulled cells off; today at 9am you started new ones. Today's files have mtimes 09:00–09:15; the most recent file in yesterday's batch is at 17:03. The gap between today's oldest active file (09:00) and yesterday's newest (17:03) is ~16 hours, far larger than 120 minutes — so yesterday's batch gets dropped and the dashboard correctly shows only today's cells.

### 2. Initial data

The first time you run, the app seeds itself from the embedded v6 snapshot — your existing chamber layout, analysis groups, and RPT cell list. From then on it persists to `data.json`.

If you ever want to nuke and re-seed, delete `data.json` and `metadata.json` and restart. Or POST to `/api/reset`.

## File parsing

```
CRD_12_CID018_Po_CYLT-1P-RPT_60C_127.0.0.1-BTS85-166-10-8-399_2.xlsx
|____________| |_______________| |__| |___________|  |_______|
   Cell ID       Test Info       Temp    Machine     Rack-Tester-Channel
```

Test type detection looks for keywords in the test-info segment in this order: `RPT`, `CYLT`/`CYC` → Cycle Life, `CLD`/`CAL` → Calendar Life, `FORM` → Formation, `HPPC`, `EIS`, `OCV`.

## API endpoints

For curiosity / debugging:

- `GET /api/data` — full lab data
- `POST /api/data` — accepts a full data blob (used by the UI). Auto-sync chamber edits to `cellId`/`currentTest`/`startDate` are silently ignored; everything else persists.
- `GET /api/watcher` — watcher status: file counts, last scan time, unmapped channels
- `POST /api/watcher/scan` — force an immediate rescan
- `POST /api/reset` — wipe `data.json` and `metadata.json`, re-seed, rescan

## Folder structure

```
cellrd/
├── cellrd_app.py      # Flask backend + watcher (run this)
├── config.json        # watch folder + rack mapping
├── seed.json          # initial chambers, RPT cells, test library
├── requirements.txt
├── static/
│   └── index.html     # the React UI (single file, no build step)
├── data.json          # created on first run — current lab state
└── metadata.json      # created on first run — user metadata for auto-sync channels
```

## Troubleshooting

**`⟲ 0 live` in the header but I have running tests**
- Check that every path in `watch_folders` (in `config.json`) is correct and that backslashes are doubled (`\\\\`) for Windows/network paths
- Hit `http://localhost:5000/api/watcher` and look at `folder_stats` — each folder reports `exists`, `file_count`, and any `error`. If a network share shows `exists: false`, the path is wrong or the share isn't reachable from this machine.
- Make sure files match the Neware naming convention. Both formats are supported: with or without a leading timestamp, with `127.0.0.1` or `127_0_0_1` IP.

**Channels show under "unmapped" in `/api/watcher`**
- The Neware rack number isn't in `rack_mapping`. Add it.
- Or that channelId is genuinely outside Chamber_1–6 and isn't supposed to be auto-synced.

**Auto-sync overwrote my analysis group**
- It shouldn't — `metadata.json` stores user fields and re-applies them on every scan. If this happens, check that `metadata.json` exists and is writable.

**I want to retire a chamber from auto-sync**
- Edit `AUTO_SYNC_CHAMBERS` near the top of `cellrd_app.py` and the matching set in `static/index.html` (search for `AUTO_SYNC_CHAMBERS = new Set`).

**Multi-user write conflicts**
- The backend uses last-write-wins. Two people editing the same chamber metadata at the same time is rare and the auto-refresh (10s) means stale state self-corrects.

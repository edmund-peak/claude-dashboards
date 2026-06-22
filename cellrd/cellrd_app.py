"""
CellRD Lab Manager (file-watcher edition)
=========================================
Single-program version of CellRD Lab Manager v6 that auto-populates the
6 Standard Testing chambers from Neware filenames in a watched folder.

For Chamber_1 through Chamber_6:
  - cellId, currentTest, startDate are owned by the file watcher.
  - User cannot edit them via the UI.
  - analysisGroup, cellFormat, durationWeeks remain user-editable
    (they're stored as "metadata" keyed by channelId and re-applied
    on every scan so the watcher's overwrites don't blow them away).

For all other chambers (RPT, Formation, Misc, CoinCell, etc.) and the
RPT Tracker tab, behavior is identical to the original v6 — fully
user-editable through the UI.

Usage:
    pip install -r requirements.txt
    python cellrd_app.py

Then open http://localhost:5000
"""

import io
import json
import os
import re
import shutil
import sys
import threading
import time
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS

ROOT = Path(__file__).parent.resolve()
CONFIG_PATH = ROOT / "config.json"
SEED_PATH   = ROOT / "seed.json"
DATA_PATH        = ROOT / "data.json"
DATA_STABLE_PATH = ROOT / "data_stable.json"
META_PATH        = ROOT / "metadata.json"  # user metadata for auto-sync channels
STORAGE_SEED_PATH  = ROOT / "storage_seed.json"   # living layout memory (updated on edits)
STORAGE_EXCEL_PATH = ROOT / "storage_excel.json"  # immutable Excel reference for re-import
RPT_SCHEDULE_SEED_PATH = ROOT / "rpt_schedule_seed.json"  # one-time next-RPT-start seed (Excel 'RPT Planner' col J)
STATIC_DIR       = ROOT / "static"

AUTO_SYNC_CHAMBERS = {
    "Chamber_1", "Chamber_2", "Chamber_3",
    "Chamber_4", "Chamber_5", "Chamber_6",
    "Chamber_7", "Chamber_8", "Chamber_9",
    "HighTemp",  "CoinCell_1", "CoinCell_2",
}

# Chambers that physically hold cells in storage between RPTs. RPT cells are
# only ever auto-assigned to these (Chamber_1 / Chamber_7 / Chamber_8 have no
# shelf space).
STORAGE_CHAMBERS = {
    "Chamber_2", "Chamber_3",
    "Chamber_4", "Chamber_5", "Chamber_6", "Chamber_9",
}

# Known non-default shelf layouts. Every shelf is a uniform 4-column × 3-row
# (12-slot) grid, so there are no special-case chambers; shelf COUNT per chamber
# comes from storageConfig (integer N => N such shelves).
DEFAULT_STORAGE_LAYOUT = {
}


# -----------------------------------------------------------------
# Config + first-run setup
# -----------------------------------------------------------------

def load_config():
    if not CONFIG_PATH.exists():
        print("[FATAL] config.json missing — see README", file=sys.stderr)
        sys.exit(1)
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


def load_seed():
    if not SEED_PATH.exists():
        print("[FATAL] seed.json missing — re-run the build "
              "(extract_seed.py)", file=sys.stderr)
        sys.exit(1)
    with open(SEED_PATH, encoding="utf-8") as f:
        return json.load(f)


CONFIG = load_config()

# Support both `watch_folders` (new, list) and `watch_folder` (legacy, single)
if "watch_folders" in CONFIG:
    WATCH_FOLDERS = list(CONFIG["watch_folders"])
elif "watch_folder" in CONFIG:
    WATCH_FOLDERS = [CONFIG["watch_folder"]]
else:
    print("[FATAL] config.json must contain 'watch_folders' (list) or "
          "'watch_folder' (string)", file=sys.stderr)
    sys.exit(1)

POLL_INTERVAL        = CONFIG.get("poll_interval_seconds", 30)
RACK_MAP             = CONFIG.get("rack_mapping", {})
FOLDER_LABELS        = {str(Path(k)).lower(): v for k, v in CONFIG.get("folder_labels", {}).items()}
BATCH_GAP_MINUTES    = CONFIG.get("batch_gap_minutes", 120)
UPDATING_THRESHOLD_MINUTES = CONFIG.get("updating_threshold_minutes", 20)
STABLE_STALE_HOURS   = CONFIG.get("stable_stale_hours", 12)


# -----------------------------------------------------------------
# Filename parsing — same logic as watcher.py
# -----------------------------------------------------------------

FILENAME_RE = re.compile(
    r"(CRD_\d+_CID\d+)"    # cell ID,  e.g. CRD_05_CID0027
    r"_(.+?)"               # test info, e.g. Po_CYLT-1P-RPT_45C
    r"_127\.0\.0\.1"        # Neware always runs on localhost
    r"-BTS\d+"              # BTS device number (not needed beyond parsing)
    r"-(\d+)-(\d+)-(\d+)"  # rack - tester - channel
    r"-(\d+(?:_\d+)?)"     # serial[_rollover], e.g. 189 or 189_4
    r"\.xlsx$",
    re.IGNORECASE,
)
TEST_TYPE_KEYWORDS = [
    ("RPT",   "RPT"),
    ("CYLT",  "Cycle Life"),
    ("CYC",   "Cycle Life"),
    ("CLD",   "Calendar Life"),   # new format: CLD-2M-100SOC
    ("CAL",   "Calendar Life"),
    ("FORM",  "Formation"),
    ("HPPC",  "HPPC"),
    ("EIS",   "EIS"),
    ("OCV",   "OCV"),
]
# Storage condition embedded in the test-info, e.g. "..._45C_..." -> 45°C.
TEMP_RE = re.compile(r"(\d{1,3})\s*C(?:\b|_|$)", re.IGNORECASE)
# State of charge, e.g. "100SOC", "0%SOC", "50 SOC" -> 1.0 / 0.0 / 0.5
SOC_RE = re.compile(r"(\d{1,3})\s*%?\s*SOC", re.IGNORECASE)


def parse_filename(filename):
    m = FILENAME_RE.search(filename)
    if not m:
        return None
    cell_id     = m.group(1)
    test_info   = m.group(2)
    rack_num    = m.group(3)
    tester_num  = m.group(4)
    channel_num = m.group(5)
    seq         = m.group(6)

    rack_code  = RACK_MAP.get(rack_num)
    prefix     = rack_code if rack_code else "?" + rack_num
    channel_id = "{}-T{:02d}-CH{:02d}".format(prefix, int(tester_num), int(channel_num))

    # Strip Neware rollover suffix (_1, _2 …) so all files for the same
    # ongoing test share one key for start-date tracking.
    base_seq = seq.split("_", 1)[0]
    test_key = "{}|{}|{}".format(cell_id, channel_id, base_seq)

    test_type = "Unknown"
    for keyword, label in TEST_TYPE_KEYWORDS:
        if keyword in test_info.upper():
            test_type = label
            break

    # Storage temp (e.g. "45C" -> "45°C"). SOC strips out first so its digits
    # aren't mistaken for a temperature.
    info_wo_soc = SOC_RE.sub("", test_info)
    temp_m = TEMP_RE.search(info_wo_soc)
    temp = temp_m.group(1) + "°C" if temp_m else None

    soc_m = SOC_RE.search(test_info)
    soc = round(int(soc_m.group(1)) / 100.0, 2) if soc_m else None

    # True only for Calendar Life RPT files (contain "CLD" or "CAL" in test_info).
    # Regular one-time RPTs on cycle-life cells have "RPT" but not "CLD"/"CAL".
    info_up = test_info.upper()
    is_calendar_life = ("CLD" in info_up) or ("CAL" in info_up)

    return {
        "cell_id":          cell_id,
        "test_type":        test_type,
        "temp":             temp,
        "soc":              soc,
        "channel_id":       channel_id,
        "test_key":         test_key,
        "filename":         filename,
        "is_calendar_life": is_calendar_life,
    }


def _iso_add(date_str, days):
    """Add `days` to an ISO date string, returning a new ISO string (or None)."""
    if not date_str:
        return None
    try:
        d = datetime.strptime(date_str, "%Y-%m-%d").date() + timedelta(days=days)
        return d.strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return None


# -----------------------------------------------------------------
# Persistence
# -----------------------------------------------------------------

_data_lock = threading.RLock()


def _seed_clone():
    seed = load_seed()
    return json.loads(json.dumps(seed))


def load_data():
    """Read data.json, or seed it if missing."""
    with _data_lock:
        if not DATA_PATH.exists():
            data = _seed_clone()
            with open(DATA_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            return data
        with open(DATA_PATH, encoding="utf-8") as f:
            data = json.load(f)
        # back-compat: ensure rptCells / rptSettings exist
        if "rptCells" not in data:
            data["rptCells"] = []
        if "rptSettings" not in data:
            data["rptSettings"] = {"durationDays": 4, "restDays": 28}
        if "storageConfig" not in data:
            data["storageConfig"] = {}  # chamber -> {shelves:[{cap,cols}]}
        for ch, layout in DEFAULT_STORAGE_LAYOUT.items():
            data["storageConfig"].setdefault(ch, layout)
        if "removedCells" not in data:
            data["removedCells"] = []   # cellIds pulled from calendar-life testing
        # One-time migration to the auto-scheduling model: wipe the old hand-
        # seeded RPT cells so the scheduler repopulates them from Neware files.
        if data.get("rptSchema") != "auto-v1":
            data["rptCells"] = []
            data["rptSchema"] = "auto-v1"
            with open(DATA_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        return data


def save_data(data, path=None):
    """Atomic write."""
    target = path if path is not None else DATA_PATH
    with _data_lock:
        tmp = target.with_suffix(".json.tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        os.replace(tmp, target)


def load_stable_data():
    """Read data_stable.json if it exists; else fall back to live data."""
    if DATA_STABLE_PATH.exists():
        with open(DATA_STABLE_PATH, encoding="utf-8") as f:
            return json.load(f)
    return load_data()


def load_storage_seed():
    """Persistent layout memory: { cellId: {chamber, shelf, slot} }.
    Seeded from the lab's Excel, then kept in sync with the user's edits."""
    try:
        with open(STORAGE_SEED_PATH, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}


def load_rpt_schedule_seed():
    """One-time next-RPT-start seed: { cellId: "YYYY-MM-DD" }, exported from the
    lab Excel 'RPT Planner' sheet column J (Next RPT Start)."""
    try:
        with open(RPT_SCHEDULE_SEED_PATH, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}


def parse_storage_from_xlsx(xlsx_bytes):
    """Pull {cellId: {chamber, shelf, slot}} from a lab Excel's 'RPT Planner'
    sheet, using stdlib only (no openpyxl). Locates the 'Cell ID' and
    'Assigned Storage Chamber' columns by header text, so column order can
    shift between spreadsheet versions. Returns {} if it can't find them."""
    z = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
    names = z.namelist()
    ss = []
    if "xl/sharedStrings.xml" in names:
        ss = re.findall(r"<t[^>]*>(.*?)</t>",
                        z.read("xl/sharedStrings.xml").decode("utf-8", "replace"), re.S)
    relmap = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"',
                             z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")))
    sheets = re.findall(r'<sheet[^>]*name="([^"]*)"[^>]*r:id="([^"]+)"',
                        z.read("xl/workbook.xml").decode("utf-8", "replace"))
    target = None
    for name, rid in sheets:
        if name.strip().lower() == "rpt planner":
            target = relmap.get(rid)
            break
    if not target and sheets:
        target = relmap.get(sheets[0][1])
    if not target:
        return {}
    target = target.lstrip("/")
    if not target.startswith("xl/"):
        target = "xl/" + target
    xml = z.read(target).decode("utf-8", "replace")
    colletter = lambda ref: re.match(r"([A-Z]+)", ref).group(1)

    def rowvals(row):
        out = {}
        for ref, t, v in re.findall(
                r'<c r="([^"]+)"(?:[^>]*t="([^"]*)")?[^>]*>(?:<v>(.*?)</v>)?', row):
            if not v:
                continue
            out[colletter(ref)] = ss[int(v)] if t == "s" else v
        return out

    rows = [rowvals(r) for r in re.findall(r"<row[^>]*>(.*?)</row>", xml, re.S)]
    cell_col = chamber_col = None
    for rv in rows:
        by_text = {str(val).strip().lower(): col for col, val in rv.items()}
        if "cell id" in by_text and "assigned storage chamber" in by_text:
            cell_col = by_text["cell id"]
            chamber_col = by_text["assigned storage chamber"]
            break
    if not cell_col or not chamber_col:
        return {}
    # Every Cell ID row is a "listed" cell (it's in the sheet, hence active).
    # chamber is None when its storage cell is blank (e.g. currently on test).
    mapping = {}
    for rv in rows:
        cid = rv.get(cell_col)
        if not (cid and re.match(r"CRD_\d+_CID\d+$", str(cid))):
            continue
        ch = rv.get(chamber_col)
        chamber = ch if (ch and str(ch).startswith("Chamber_")) else None
        mapping[cid] = {"chamber": chamber, "shelf": None, "slot": None}
    return mapping


def save_storage_seed(data):
    """Re-export the seed from the current layout so it always reflects the
    latest arrangement. Called whenever the UI saves a layout change; that way
    a cell returns to where the user last put it if it's ever re-detected or
    the live data is rebuilt."""
    seed = load_storage_seed()
    seed = dict(seed) if isinstance(seed, dict) else {}
    for c in data.get("rptCells", []):
        cid = c.get("cellId")
        if cid and c.get("assignedStorage"):
            seed[cid] = {
                "chamber": c.get("assignedStorage"),
                "shelf":   c.get("storageShelf"),
                "slot":    c.get("storageSlot"),
            }
    try:
        tmp = STORAGE_SEED_PATH.with_suffix(".json.tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(seed, f, indent=0, sort_keys=True)
        os.replace(tmp, STORAGE_SEED_PATH)
    except OSError:
        pass


def load_metadata():
    """User metadata (project, analysisGroup, cellFormat, durationWeeks,
    purpose) keyed by cellId — so it follows the cell, not the channel slot.

    Schema versions:
      - (unset / old)  : top-level keys are channelIds (legacy)
      - "cell"         : by_cell dict, analysisGroup is the only campaign tag
      - "cell-v2"      : by_cell dict, both `project` and `analysisGroup`
                         (project is the primary tag; analysisGroup is an
                         optional sub-tag within the project)
    """
    if not META_PATH.exists():
        return {"version": "cell-v2", "by_cell": {}}
    with open(META_PATH, encoding="utf-8") as f:
        raw = json.load(f)

    # Already in the latest format
    if isinstance(raw, dict) and raw.get("version") == "cell-v2" and "by_cell" in raw:
        return raw

    # Migration path 1: "cell" → "cell-v2"
    # The old "analysisGroup" semantically becomes "project" (the primary
    # campaign tag). The new "analysisGroup" becomes an optional sub-tag,
    # which is empty for all migrated entries.
    if isinstance(raw, dict) and raw.get("version") == "cell" and "by_cell" in raw:
        print("[MIGRATE] Promoting metadata.json analysisGroup -> project (cell-v2)")
        new_by_cell = {}
        for cell_id, m in raw["by_cell"].items():
            new_by_cell[cell_id] = {
                "project":       m.get("analysisGroup"),
                "analysisGroup": None,
                "cellFormat":    m.get("cellFormat"),
                "durationWeeks": m.get("durationWeeks"),
                "purpose":       m.get("purpose"),
            }
        migrated = {"version": "cell-v2", "by_cell": new_by_cell}
        tmp = META_PATH.with_suffix(".json.tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(migrated, f, indent=2)
        os.replace(tmp, META_PATH)
        return migrated

    # Migration path 2 (legacy): channel-keyed → cell-keyed v2.
    # Used when someone is jumping from the very first version to now.
    # Look up what cellId was on each channel in data.json (if it exists)
    # to map channelId-keyed entries to cellId-keyed entries.
    print("[MIGRATE] Converting metadata.json from channel-keyed to cell-keyed v2")
    by_cell = {}
    if DATA_PATH.exists():
        try:
            with open(DATA_PATH, encoding="utf-8") as f:
                data = json.load(f)
            for c in data.get("chambers", []):
                for ch in c.get("channels", []):
                    cid_channel = ch.get("channelId")
                    cid_cell    = ch.get("cellId")
                    if not cid_channel or not cid_cell:
                        continue
                    old_meta = raw.get(cid_channel)
                    if not old_meta:
                        continue
                    # Skip empty entries
                    if not any(old_meta.get(k) for k in
                               ("analysisGroup", "cellFormat", "durationWeeks")):
                        continue
                    by_cell[cid_cell] = {
                        "project":       old_meta.get("analysisGroup"),
                        "analysisGroup": None,
                        "cellFormat":    old_meta.get("cellFormat"),
                        "durationWeeks": old_meta.get("durationWeeks"),
                        "purpose":       old_meta.get("purpose"),
                    }
            print("[MIGRATE] Mapped {} entries from channels to cells".format(
                len(by_cell)))
        except Exception as e:
            print("[MIGRATE] Could not read data.json for migration: " + str(e))

    migrated = {"version": "cell-v2", "by_cell": by_cell}
    tmp = META_PATH.with_suffix(".json.tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(migrated, f, indent=2)
    os.replace(tmp, META_PATH)
    return migrated


def save_metadata(meta):
    with _data_lock:
        tmp = META_PATH.with_suffix(".json.tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(meta, f, indent=2)
        os.replace(tmp, META_PATH)


# On first run, populate metadata.json from the seed's auto-sync chambers
# (keyed by cellId so it follows the cell).
def _bootstrap_metadata():
    if META_PATH.exists():
        return
    seed = _seed_clone()
    by_cell = {}
    for c in seed["chambers"]:
        if c["name"] not in AUTO_SYNC_CHAMBERS:
            continue
        for ch in c["channels"]:
            cell_id = ch.get("cellId")
            if not cell_id:
                continue
            # Only bootstrap if there's actually metadata to preserve
            if not any(ch.get(k) for k in
                       ("analysisGroup", "cellFormat", "durationWeeks")):
                continue
            # The seed used "analysisGroup" as the primary tag; in v2 that
            # becomes "project". analysisGroup (now a sub-tag) starts empty.
            by_cell[cell_id] = {
                "project":       ch.get("analysisGroup"),
                "analysisGroup": None,
                "cellFormat":    ch.get("cellFormat"),
                "durationWeeks": ch.get("durationWeeks"),
                "purpose":       ch.get("purpose"),
            }
    save_metadata({"version": "cell-v2", "by_cell": by_cell})
    print("[INIT] Bootstrapped metadata.json from seed for "
          + str(len(by_cell)) + " cells")


# -----------------------------------------------------------------
# Scanner
# -----------------------------------------------------------------

class NewareScanner:
    def __init__(self, folders):
        self.folders = [Path(f) for f in folders]
        self.last_scan = None
        self.newest_file_time = None
        self.cutoff_time = None
        self.total_files = 0
        self.active_count = 0
        self.stable_active_count = 0
        self.any_folder_updating = False
        self.stable_frozen = False
        self.stable_last_refreshed = None
        self.unmapped_channels = []  # files whose channel_id doesn't match any chamber
        self.folder_stats = []       # per-folder: {folder, exists, file_count}
        self.lock = threading.Lock()

    def scan(self):
        """Scan all configured folders, compute live + stable batches,
        and write both to data.json and data_stable.json."""
        # Pass 1: parse all files with their mtimes from every folder
        all_files = []
        folder_stats = []
        for folder in self.folders:
            stat = {"folder": str(folder), "label": FOLDER_LABELS.get(str(folder).lower()),
                    "exists": False, "file_count": 0, "error": None}
            try:
                if not folder.exists():
                    print("[WARN] Watch folder does not exist: " + str(folder))
                    folder_stats.append(stat)
                    continue
                stat["exists"] = True
                count = 0
                for f in folder.glob("*.xlsx"):
                    parsed = parse_filename(f.name)
                    if not parsed:
                        continue
                    try:
                        st = f.stat()
                        mtime = datetime.fromtimestamp(st.st_mtime)
                        # File creation time: prefer st_birthtime (Mac, BSD, Win
                        # on Python 3.12+); else st_ctime, which IS creation
                        # time on Windows. On Linux, st_ctime is inode-change
                        # time — close enough to creation for our purposes
                        # since these files don't have inode changes happening.
                        if hasattr(st, "st_birthtime"):
                            ctime = datetime.fromtimestamp(st.st_birthtime)
                        else:
                            ctime = datetime.fromtimestamp(st.st_ctime)
                    except OSError:
                        continue
                    parsed["file_modified"] = mtime
                    parsed["file_created"]  = ctime
                    parsed["source_folder"] = str(folder)
                    all_files.append(parsed)
                    count += 1
                stat["file_count"] = count
            except OSError as e:
                # e.g. network share unreachable
                stat["error"] = str(e)
                print("[WARN] Could not scan " + str(folder) + ": " + str(e))
            folder_stats.append(stat)

        # Pass 2: per-folder batch detection — but now extract TWO batches
        # per folder (live = first, stable = first or second depending on
        # whether the folder is currently being updated).
        gap = timedelta(minutes=BATCH_GAP_MINUTES)
        updating_threshold = timedelta(minutes=UPDATING_THRESHOLD_MINUTES)
        now = datetime.now()

        live_active = []
        per_folder_newest = {}
        per_folder_live_cutoff = {}
        per_folder_updating = {}

        # Group files by source folder
        files_by_folder = {}
        for f in all_files:
            files_by_folder.setdefault(f["source_folder"], []).append(f)

        for folder_str, folder_files in files_by_folder.items():
            sorted_files = sorted(
                folder_files, key=lambda f: f["file_modified"], reverse=True)
            f_newest = sorted_files[0]["file_modified"]
            per_folder_newest[folder_str] = f_newest

            # Is this folder currently being updated?
            is_updating = (now - f_newest) <= updating_threshold
            per_folder_updating[folder_str] = is_updating

            # Walk newest-first; include files as long as consecutive mtimes
            # are within the batch-gap threshold. This is the "live" batch:
            # everything Neware is currently writing or recently touched.
            live_batch = [sorted_files[0]]
            for i in range(1, len(sorted_files)):
                prev_mtime = sorted_files[i - 1]["file_modified"]
                this_mtime = sorted_files[i]["file_modified"]
                if (prev_mtime - this_mtime) > gap:
                    break
                live_batch.append(sorted_files[i])

            per_folder_live_cutoff[folder_str] = live_batch[-1]["file_modified"]
            live_active.extend(live_batch)

        # For top-level "newest" / "cutoff" reporting, use globals across folders
        if live_active:
            newest = max(f["file_modified"] for f in live_active)
            cutoff = min(f["file_modified"] for f in live_active)
        else:
            newest, cutoff = None, None

        any_updating = any(per_folder_updating.values())

        # Annotate folder_stats with per-folder batch info for the API
        for fs in folder_stats:
            folder = fs["folder"]
            fs["newest"] = per_folder_newest.get(folder)
            fs["batch_cutoff"] = per_folder_live_cutoff.get(folder)
            fs["updating"] = per_folder_updating.get(folder, False)
            fs["active_files"] = sum(
                1 for f in live_active if f["source_folder"] == folder)
            if fs["newest"]:
                fs["newest"] = fs["newest"].isoformat()
            if fs["batch_cutoff"]:
                fs["batch_cutoff"] = fs["batch_cutoff"].isoformat()

        # Build a min-ctime lookup keyed by test_key. Neware rolls over to
        # a new file (..._1.xlsx, ..._2.xlsx, ...) when the current one gets
        # too big. All rollover files share the same test_key, and the
        # *original* file's creation time is the real test start. We compute
        # this across ALL parseable files (not just the active batch) — the
        # original file may be many months old and outside the gap window,
        # but it's still the start of the ongoing test.
        min_ctime_by_test_key = {}
        for f in all_files:
            tk = f["test_key"]
            ct = f["file_created"]
            cur = min_ctime_by_test_key.get(tk)
            if cur is None or ct < cur:
                min_ctime_by_test_key[tk] = ct

        # Pass 3+4: apply the live active set to data.json (always).
        # active_test_keys = the runs Neware is writing right now; the archive
        # excludes these so it only ever shows *finished* tests.
        active_test_keys = {f["test_key"] for f in live_active}
        live_unmapped = self._apply_active(
            live_active, min_ctime_by_test_key, all_files, active_test_keys)

        # Stable view: freeze during updates so the UI has a calm picture
        # while Neware is actively writing. Advance when quiet, or after
        # STABLE_STALE_HOURS as a safeguard against long uninterrupted runs.
        try:
            stable_age = (now - datetime.fromtimestamp(DATA_STABLE_PATH.stat().st_mtime)
                          if DATA_STABLE_PATH.exists() else None)
        except OSError:
            stable_age = None

        stable_is_stale = stable_age is not None and stable_age > timedelta(hours=STABLE_STALE_HOURS)
        if not any_updating or not DATA_STABLE_PATH.exists() or stable_is_stale:
            with _data_lock:
                save_data(json.loads(DATA_PATH.read_text(encoding="utf-8")), path=DATA_STABLE_PATH)
            stable_refreshed = True
        else:
            stable_refreshed = False

        # Count assigned cells in both live and stable data files.
        # Both use the same methodology (cellId present in auto-sync chambers)
        # so the two numbers are directly comparable in the UI.
        def _count_cells(path):
            count = 0
            try:
                with open(path, encoding="utf-8") as f:
                    d = json.load(f)
                for c in d.get("chambers", []):
                    if c.get("name") not in AUTO_SYNC_CHAMBERS:
                        continue
                    for ch in c.get("channels", []):
                        if ch.get("cellId"):
                            count += 1
            except (OSError, json.JSONDecodeError):
                pass
            return count

        live_active_count   = _count_cells(DATA_PATH)
        stable_active_count = _count_cells(DATA_STABLE_PATH)

        with self.lock:
            self.last_scan = datetime.now().isoformat()
            self.newest_file_time = newest.isoformat() if newest else None
            self.cutoff_time = cutoff.isoformat() if cutoff else None
            self.total_files = len(all_files)
            self.active_count = live_active_count
            self.stable_active_count = stable_active_count
            self.unmapped_channels = live_unmapped
            self.folder_stats = folder_stats
            self.any_folder_updating = any_updating
            self.stable_frozen = any_updating and not stable_refreshed
            # Time of last stable refresh in iso (for the UI to show "frozen since")
            if stable_refreshed:
                self.stable_last_refreshed = datetime.now().isoformat()

        print("[SCAN] {} files across {} folders, live={} stable={} mapped, "
              "{} unmapped, updating={}, stable_{} @ {}".format(
                  len(all_files), len(self.folders),
                  self.active_count, self.stable_active_count,
                  len(live_unmapped),
                  sum(1 for v in per_folder_updating.values() if v),
                  "refreshed" if stable_refreshed else "frozen",
                  datetime.now().strftime("%H:%M:%S")))

    def _apply_active(self, active, min_ctime_by_test_key, all_files, active_test_keys):
        """Merge the active file batch into data.json. Returns unmapped channels."""
        # Pick the most recently-modified file per channel
        by_channel = {}
        for f in active:
            cid = f["channel_id"]
            if cid not in by_channel or f["file_modified"] > by_channel[cid]["file_modified"]:
                by_channel[cid] = f

        with _data_lock:
            data = load_data()
            meta = load_metadata()
            by_cell_meta = meta.get("by_cell", {})

            # Build a quick lookup: channelId -> (chamber_idx, channel_idx)
            channel_lookup = {}
            for ci, c in enumerate(data["chambers"]):
                if c["name"] not in AUTO_SYNC_CHAMBERS:
                    continue
                for chi, ch in enumerate(c["channels"]):
                    if ch.get("channelId"):
                        channel_lookup[ch["channelId"]] = (ci, chi)

            # Blank out all auto-sync channels first.
            for ci, c in enumerate(data["chambers"]):
                if c["name"] not in AUTO_SYNC_CHAMBERS:
                    continue
                for chi, ch in enumerate(c["channels"]):
                    ch["cellId"]        = None
                    ch["currentTest"]   = None
                    ch["startDate"]     = None
                    ch["filename"]      = None
                    ch["project"]       = None
                    ch["analysisGroup"] = None
                    ch["cellFormat"]    = None
                    ch["durationWeeks"] = None

            unmapped = []
            for cid, info in by_channel.items():
                if cid not in channel_lookup:
                    unmapped.append({
                        "channel_id":    cid,
                        "cell_id":       info["cell_id"],
                        "filename":      info["filename"],
                        "source_folder": info.get("source_folder"),
                    })
                    continue
                ci, chi = channel_lookup[cid]
                ch = data["chambers"][ci]["channels"][chi]
                cell_id = info["cell_id"]
                ch["cellId"]      = cell_id
                ch["currentTest"] = info["test_type"]
                ch["filename"]    = info.get("filename")
                start_ct = min_ctime_by_test_key.get(
                    info["test_key"], info["file_created"])
                ch["startDate"]   = start_ct.strftime("%Y-%m-%d")
                ch["startTs"]     = start_ct.isoformat() if start_ct else None
                cm = by_cell_meta.get(cell_id, {})
                ch["project"]       = cm.get("project")
                ch["analysisGroup"] = cm.get("analysisGroup")
                ch["cellFormat"]    = cm.get("cellFormat")
                ch["durationWeeks"] = cm.get("durationWeeks")

            self._update_archive(all_files, active_test_keys, data)
            self._seed_rpt_schedule(data)
            self._sync_rpt_cells(by_channel, min_ctime_by_test_key, data)
            self._seed_storage(data)
            save_data(data)
            return unmapped

    def _storage_capacity(self, data, chamber):
        """Total slots in a chamber, from storageConfig (mirrors the client's
        shelvesOf): dict shelves sum their caps, a bare int N => N shelves of 12,
        otherwise one 12-slot shelf."""
        cfg = (data.get("storageConfig") or {}).get(chamber)
        if isinstance(cfg, dict) and isinstance(cfg.get("shelves"), list) and cfg["shelves"]:
            return sum(s.get("cap", 12) for s in cfg["shelves"])
        if isinstance(cfg, (int, float)) and cfg > 0:
            return int(cfg) * 12
        return 12

    def _pack_chamber(self, data, temp):
        """Pick a storage chamber for a new cell at `temp` by PACKING: the
        lowest-numbered matching chamber that still has a free slot (keeps cells
        concentrated rather than spread). Returns None if no chamber matches."""
        if not temp:
            return None
        def num(n):
            m = re.search(r"(\d+)", n)
            return int(m.group(1)) if m else 999
        candidates = sorted(
            [c["name"] for c in data.get("chambers", [])
             if c.get("temp") == temp and c["name"] in STORAGE_CHAMBERS],
            key=num)
        if not candidates:
            return None
        counts = {}
        for c in data.get("rptCells", []):
            s = c.get("assignedStorage")
            if s in candidates:
                counts[s] = counts.get(s, 0) + 1
        for ch in candidates:
            if counts.get(ch, 0) < self._storage_capacity(data, ch):
                return ch
        return candidates[-1]  # all full → overflow into the last

    def _update_archive(self, all_files, active_test_keys, data):
        """Rebuild a compact ledger of every cell ever run and the tests it ran,
        swept fresh from *all* files on disk (as far back as the folders go).
        Lives in data['archive'] keyed by cellId:

            { "<cellId>": {
                "tests": { "<testType>": {"first": iso, "last": iso, "file": name} },
                "lastSeen": iso } }

        One small entry per cell, two dates per distinct test — deliberately
        tiny. Currently-running tests (active_test_keys) are excluded, so the
        archive only ever shows *finished* work. Rebuilt each scan, so it's
        idempotent and self-heals if files are added/removed.
        """
        # Aggregate first/last per (cellId, testType), skipping in-progress runs.
        per = {}
        for f in all_files:
            if f.get("test_key") in active_test_keys:
                continue
            cid = f.get("cell_id")
            if not cid:
                continue
            tt = f.get("test_type") or "Unknown"
            fc = f.get("file_created")
            fm = f.get("file_modified")
            if fc is None and fm is None:
                continue
            first = fc or fm
            last = fm or fc
            fn = f.get("filename")
            e = per.get((cid, tt))
            if e is None:
                # [first, last, file-of-latest]
                per[(cid, tt)] = [first, last, fn]
            else:
                if first < e[0]:
                    e[0] = first
                if last > e[1]:
                    e[1] = last
                    e[2] = fn  # keep the filename from the most recent run

        arch = {}
        for (cid, tt), (first, last, fn) in per.items():
            rec = arch.setdefault(cid, {"tests": {}, "lastSeen": None})
            first_s = first.strftime("%Y-%m-%d")
            last_s = last.strftime("%Y-%m-%d")
            rec["tests"][tt] = {"first": first_s, "last": last_s, "file": fn}
            if not rec["lastSeen"] or last_s > rec["lastSeen"]:
                rec["lastSeen"] = last_s
        data["archive"] = arch

    def _seed_storage(self, data):
        """One-time bulk initialization of where existing cells are stored, from
        storage_seed.json (initially exported from the lab's Excel). Applies the
        saved chamber/shelf/slot to matching cells. Runs once (guarded by the
        'storageSeeded' flag); after that, the seed is kept in sync with the
        user's layout edits (see save_storage_seed) and used to restore a cell's
        spot when it's re-detected (see _sync_rpt_cells).
        """
        if data.get("storageSeeded"):
            return
        seed = load_storage_seed()
        if not seed:
            return  # no seed file yet — don't set the flag, retry next scan
        applied = 0
        for c in data.get("rptCells", []):
            spot = seed.get(c.get("cellId"))
            if spot and spot.get("chamber"):
                c["assignedStorage"] = spot["chamber"]
                c["storageShelf"] = spot.get("shelf")
                c["storageSlot"] = spot.get("slot")   # may be None -> client packs
                applied += 1
        data["storageSeeded"] = True
        print("[SEED] storage initialized: {} cells placed".format(applied))

    def _seed_rpt_schedule(self, data):
        """One-time seed of each cell's NEXT RPT start from the lab Excel
        (rpt_schedule_seed.json, exported from 'RPT Planner' col J). Sets
        `nextRptOverride` per cell so the upcoming RPT is pinned to the lab's
        planned date instead of the return+rest estimate (which, for historical
        cells, was computed off test-end and under-counts storage time).

        The override is honoured by the schedule-derive step below until the
        cell actually starts that RPT run (a new event is appended), at which
        point it is cleared and scheduling reverts to the normal
        physical-return + rest cadence. Runs once (guarded by
        'rptScheduleSeeded'); if the seed file is absent the flag is not set so
        it retries on a later scan.
        """
        if data.get("rptScheduleSeeded"):
            return
        seed = load_rpt_schedule_seed()
        if not seed:
            return
        by_id = {c.get("cellId"): c for c in data.get("rptCells", [])}
        applied = skipped_on_test = 0
        for cid, start in seed.items():
            c = by_id.get(cid)
            if not (c and start):
                continue
            # Only pin RESTING cells (last event closed). A cell currently on
            # test / awaiting return is left to recompute its next start from
            # its actual physical return date (returned + rest), not the Excel.
            evs = c.get("events", [])
            last = evs[-1] if evs else None
            if last and last.get("returned") is None:
                skipped_on_test += 1
                continue
            c["nextRptOverride"] = start
            applied += 1
        data["rptScheduleSeeded"] = True
        print("[SEED] RPT schedule seeded: {} resting cells pinned to Excel next-start ({} on-test skipped)".format(applied, skipped_on_test))

    def _sync_rpt_cells(self, by_channel, min_ctime_by_test_key, data):
        """Auto-manage data['rptCells'] from the running file record.

        The schedule is event-driven, mirroring the lab's spreadsheet: each
        detected RPT run is an event {start, offTest, returned}. `offTest` is the
        auto-detected end of cycling (last file write); `returned` is the
        user-logged date the cell physically went back to storage. The 28-day
        storage clock — and thus the next RPT — keys off `returned`, never
        offTest. The watcher revises the schedule as cells run:

        - First RPT seen for a cell  -> create the cell (storage temp + SOC
          from the filename, temp-matched least-full storage chamber) with an
          open event.
        - A new RPT run after the last one returned -> append a new open event.
        - `running` is True while the RPT file is being actively written.
        - When a run stops but its event is still open, the cell is "awaiting
          return": the UI prompts for the date it went back to storage, which
          closes the event and schedules the next RPT (return + rest days).

        Storage location and event return-dates are owned by the user; the
        watcher never overwrites them. Cells the user removed from calendar-
        life testing (data['removedCells']) are never recreated.
        """
        cells = data.setdefault("rptCells", [])
        removed = set(data.get("removedCells", []))
        by_id = {c.get("cellId"): c for c in cells}
        rs = data.get("rptSettings") or {}
        duration = rs.get("durationDays", 4)
        rest = rs.get("restDays", 28)
        now = datetime.now()
        updating = timedelta(minutes=UPDATING_THRESHOLD_MINUTES)

        # Cells with a Calendar Life RPT file in the active batch, and whether
        # it's "live" (touched within the updating threshold => actively cycling).
        # Only "CLD"/"CAL" files qualify — not one-time RPTs on cycle-life cells.
        running_now = {}
        for info in by_channel.values():
            if not info.get("is_calendar_life"):
                continue
            mod = info.get("file_modified")
            start = min_ctime_by_test_key.get(info["test_key"], info.get("file_created"))
            running_now[info["cell_id"]] = {
                "live":    mod is not None and (now - mod) <= updating,
                "start":   start.strftime("%Y-%m-%d") if start else None,
                "startTs": start.isoformat() if start else None,   # full timestamp (file creation)
                "lastMod": mod.strftime("%Y-%m-%d") if mod else None,
                "channel": info["channel_id"],
                "temp":    info.get("temp"),
                "soc":     info.get("soc"),
                "file":    info.get("filename"),
            }

        # Restore a re-detected cell to its last-known spot; otherwise default
        # to temp-based packing for genuinely new cells.
        seed = load_storage_seed()
        # 1) Create new cells / open a new event when a fresh RPT run starts.
        for cell_id, r in running_now.items():
            if cell_id in removed:
                continue
            c = by_id.get(cell_id)
            if c is None:
                spot = seed.get(cell_id) or {}
                c = {
                    "cellId":          cell_id,
                    "storageTemp":     r["temp"],
                    "soc":             r["soc"],
                    "rptType":         "Calendar Life RPT",
                    "assignedChannel": r["channel"],
                    "assignedStorage": spot.get("chamber") or self._pack_chamber(data, r["temp"]),
                    "storageShelf":    spot.get("shelf"),
                    "storageSlot":     spot.get("slot"),   # client packs if None
                    "anchorDate":      r["start"],
                    "events":          [{"start": r["start"], "startTs": r.get("startTs"), "offTest": None, "returned": None, "file": r.get("file")}],
                }
                cells.append(c)
                by_id[cell_id] = c
            else:
                evs = c.setdefault("events", [])
                last = evs[-1] if evs else None
                # New cycle: only when the previous one is closed and this run
                # is newer than the last recorded start.
                if (last is None or last.get("returned") is not None) and r["start"] \
                        and (last is None or r["start"] > (last.get("start") or "")):
                    evs.append({"start": r["start"], "startTs": r.get("startTs"), "offTest": None, "returned": None, "file": r.get("file")})
                    # The seeded next-RPT start has begun — consume the override
                    # so future cycles schedule off the real physical return.
                    c["nextRptOverride"] = None
                elif last is not None and last.get("returned") is None:
                    # Same run still going — keep the latest filename and stamp the
                    # start timestamp on the open event (idempotent; same value).
                    if r.get("file"):
                        last["file"] = r["file"]
                    if r.get("startTs") and not last.get("startTs"):
                        last["startTs"] = r["startTs"]
                c["assignedChannel"] = r["channel"]

        # 2) Derive run-state + schedule for every cell (safe to overwrite).
        for c in cells:
            evs = c.get("events", [])
            last = evs[-1] if evs else None
            r = running_now.get(c.get("cellId"))
            open_event = bool(last and last.get("returned") is None)
            c["running"] = bool(r and r["live"] and open_event)
            # Data-estimated off-test date = last time the RPT file was written.
            # Surfaced as a pre-fill in the return picker so the user can confirm
            # "yes, I pulled it when I thought I did." Persist once and never
            # overwrite with None when the file ages out of the active batch.
            if open_event and r and r.get("lastMod"):
                c["estReturn"] = r["lastMod"]
                # Stamp the open event's own off-test date (when cycling actually
                # finished), kept distinct from the user-logged physical return.
                last["offTest"] = r["lastMod"]
            # Backfill offTest on every event missing it (idempotent). Closed
            # events: the recorded end ~= off-test, so reuse `returned`. Open
            # events with no file activity in the batch: fall back to the cell's
            # last estimate, else the nominal due date (start + duration).
            # offTest is display-only (it drives the shown test duration); the
            # schedule keys off `returned`, never off offTest.
            for e in evs:
                if e.get("offTest"):
                    continue
                if e.get("returned"):
                    e["offTest"] = e["returned"]
                else:
                    e["offTest"] = c.get("estReturn") or _iso_add(e.get("start"), duration)
            if open_event:
                c["nextRptStart"] = None          # on test / awaiting return
            elif c.get("nextRptOverride"):
                # Lab-Excel-pinned next start (until that RPT run begins).
                c["nextRptStart"] = c["nextRptOverride"]
            elif last and last.get("returned"):
                c["nextRptStart"] = _iso_add(last["returned"], rest)
            else:
                c["nextRptStart"] = c.get("anchorDate")
            c["nextRptMonth"] = str(sum(1 for e in evs if e.get("returned"))) + "M"
            # Expected return ("the day it was due") = run start + duration.
            c["dueReturn"] = _iso_add(last["start"], duration) if (last and last.get("start")) else None

    def status(self):
        with self.lock:
            return {
                "last_scan":         self.last_scan,
                "newest_file":       self.newest_file_time,
                "cutoff":            self.cutoff_time,
                "total_files":       self.total_files,
                "active_channels":   self.active_count,
                "stable_active_channels": self.stable_active_count,
                "any_folder_updating": self.any_folder_updating,
                "stable_frozen":     self.stable_frozen,
                "stable_last_refreshed": self.stable_last_refreshed,
                "unmapped":          self.unmapped_channels,
                "watch_folders":     [str(f) for f in self.folders],
                "folder_stats":      self.folder_stats,
                "batch_gap_min":     BATCH_GAP_MINUTES,
                "updating_threshold_min": UPDATING_THRESHOLD_MINUTES,
                "stable_stale_hours": STABLE_STALE_HOURS,
                "poll_interval_s":   POLL_INTERVAL,
                "auto_sync_chambers": sorted(AUTO_SYNC_CHAMBERS),
            }


SCANNER = NewareScanner(WATCH_FOLDERS)


# -----------------------------------------------------------------
# Flask app
# -----------------------------------------------------------------

app = Flask(__name__, static_folder=str(STATIC_DIR))
CORS(app)


@app.route("/")
def index():
    return send_from_directory(str(STATIC_DIR), "index.html")


@app.route("/api/data", methods=["GET"])
def api_data_get():
    view = request.args.get("view", "live")
    data = load_stable_data() if view == "stable" else load_data()
    # Per-cell notes live in metadata.json (durable across scanner rebuilds),
    # keyed by cellId. Surface them on the data blob the UI consumes.
    _meta = load_metadata()
    data["cellNotes"] = _meta.get("cell_notes", {})
    # Per-cell user metadata (project/analysisGroup/…) keyed by cellId, so the
    # drawer can show & edit Project for ANY cell — including calendar-life RPT
    # cells that aren't currently on a channel.
    data["cellMeta"] = _meta.get("by_cell", {})
    return jsonify(data)


@app.route("/api/data", methods=["POST"])
def api_data_post():
    """Accept a full data blob from the UI. For auto-sync chambers, only
    the metadata fields are persisted into metadata.json; the watcher
    fields are ignored. For all other chambers, the whole channel is
    persisted as-is."""
    try:
        new_data = request.get_json(force=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 400

    if not new_data or "chambers" not in new_data:
        return jsonify({"error": "missing chambers"}), 400

    with _data_lock:
        current = load_data()
        meta = load_metadata()

        # Overwrite non-auto-sync chambers wholesale; for auto-sync,
        # capture user-metadata fields and drop the rest.
        for new_c in new_data["chambers"]:
            # Find matching chamber in current
            target = None
            for c in current["chambers"]:
                if c["name"] == new_c["name"]:
                    target = c
                    break
            if target is None:
                continue

            if new_c["name"] in AUTO_SYNC_CHAMBERS:
                # Only update metadata; keep current channel watcher state.
                # Metadata is keyed by cellId, not channelId — so it follows
                # the cell wherever it goes.
                #
                # IMPORTANT: we take the cellId from the payload itself, not
                # from looking up the channel in server-side live data. This
                # is what makes edits work correctly in stable view: when the
                # user assigned a group while looking at Cell X, X is what
                # was in the payload, and X is what we save under — even if
                # live data shows a different cell on that channel right now.
                by_cell_meta = meta.setdefault("by_cell", {})
                for new_ch in new_c.get("channels", []):
                    payload_cell_id = new_ch.get("cellId")
                    if not payload_cell_id:
                        # No cell in this row — nothing to attach metadata to.
                        continue
                    by_cell_meta[payload_cell_id] = {
                        "project":       new_ch.get("project"),
                        "analysisGroup": new_ch.get("analysisGroup"),
                        "cellFormat":    new_ch.get("cellFormat"),
                        "durationWeeks": new_ch.get("durationWeeks"),
                        "purpose":       new_ch.get("purpose"),
                    }
                # Allow temp changes (the chip in the UI lets users edit it)
                if "temp" in new_c:
                    target["temp"] = new_c["temp"]
            else:
                # Full overwrite for user-controlled chambers
                target.update(new_c)

        # rptCells, rptSettings, testLibrary -- always full overwrite
        for key in ("rptCells", "rptSettings", "testLibrary", "storageConfig", "removedCells", "archivedCells"):
            if key in new_data:
                current[key] = new_data[key]

        # Per-cell notes are durable user metadata, keyed by cellId — store
        # them in metadata.json so the scanner's rebuilds never wipe them.
        if "cellNotes" in new_data and isinstance(new_data["cellNotes"], dict):
            meta["cell_notes"] = {k: v for k, v in new_data["cellNotes"].items() if v}

        # Project edits from the drawer (any cell, keyed by cellId). Merge into
        # by_cell so other metadata fields (analysisGroup/cellFormat/…) survive.
        if "cellProjects" in new_data and isinstance(new_data["cellProjects"], dict):
            by_cell_meta = meta.setdefault("by_cell", {})
            for cid, proj in new_data["cellProjects"].items():
                entry = by_cell_meta.setdefault(cid, {})
                entry["project"] = (proj or None)

        # Re-apply the now-updated metadata to live data so the immediate
        # response reflects the user's edit. Without this, the user wouldn't
        # see their change until the next background scan (30s).
        by_cell_meta = meta.get("by_cell", {})
        for c in current.get("chambers", []):
            if c["name"] not in AUTO_SYNC_CHAMBERS:
                continue
            for ch in c["channels"]:
                cell_id = ch.get("cellId")
                if cell_id:
                    cm = by_cell_meta.get(cell_id, {})
                    ch["project"]       = cm.get("project")
                    ch["analysisGroup"] = cm.get("analysisGroup")
                    ch["cellFormat"]    = cm.get("cellFormat")
                    ch["durationWeeks"] = cm.get("durationWeeks")

        save_data(current)
        save_metadata(meta)
        # Keep the storage seed in lock-step with the user's layout edits.
        save_storage_seed(current)

        # Same treatment for data_stable.json: propagate non-auto-sync edits
        # wholesale, and refresh auto-sync metadata against the cells that
        # are in stable view (which may differ from live).
        _propagate_to_stable(current, meta)

    return jsonify({"status": "ok"})


def _propagate_to_stable(live_data, meta):
    """Update data_stable.json to reflect a fresh write to data.json:
    - Copy over non-auto-sync chambers (full overwrite — these are user-edited)
    - Copy rptCells / rptSettings / testLibrary
    - For auto-sync chambers, re-apply metadata to the existing stable channels
      (preserving stable's own cellId/test/startDate, since those reflect a
      different point in time than live's).
    """
    if not DATA_STABLE_PATH.exists():
        return
    with _data_lock:
        with open(DATA_STABLE_PATH, encoding="utf-8") as f:
            stable = json.load(f)
        by_cell_meta = meta.get("by_cell", {})

        # Copy non-auto-sync chambers wholesale and refresh auto-sync metadata
        for live_c in live_data.get("chambers", []):
            stable_c = None
            for s in stable.get("chambers", []):
                if s["name"] == live_c["name"]:
                    stable_c = s
                    break
            if stable_c is None:
                continue
            if live_c["name"] in AUTO_SYNC_CHAMBERS:
                # Refresh metadata on each channel based on whatever cell is
                # *currently* in stable view (which may differ from live).
                for st_ch in stable_c["channels"]:
                    cell_id = st_ch.get("cellId")
                    if cell_id:
                        cm = by_cell_meta.get(cell_id, {})
                        st_ch["project"]       = cm.get("project")
                        st_ch["analysisGroup"] = cm.get("analysisGroup")
                        st_ch["cellFormat"]    = cm.get("cellFormat")
                        st_ch["durationWeeks"] = cm.get("durationWeeks")
                    else:
                        st_ch["project"]       = None
                        st_ch["analysisGroup"] = None
                        st_ch["cellFormat"]    = None
                        st_ch["durationWeeks"] = None
                # Temp can be edited via the chip
                if "temp" in live_c:
                    stable_c["temp"] = live_c["temp"]
            else:
                # User-edited chambers: stable should mirror live exactly
                stable_c.update(live_c)

        # Other top-level keys (rptCells, rptSettings, testLibrary) — copy
        for key in ("rptCells", "rptSettings", "testLibrary", "storageConfig", "removedCells", "archive", "archivedCells"):
            if key in live_data:
                stable[key] = live_data[key]

        save_data(stable, path=DATA_STABLE_PATH)


@app.route("/api/watcher", methods=["GET"])
def api_watcher_status():
    return jsonify(SCANNER.status())


@app.route("/api/watcher/scan", methods=["POST"])
def api_watcher_scan():
    SCANNER.scan()
    return jsonify({"status": "ok", "last_scan": SCANNER.last_scan})


def _win_focus_explorer(folder):
    """Bring the Explorer window for `folder` to the foreground.

    Windows suppresses foreground changes requested by a background process
    (our Flask server), so a freshly opened Explorer window just blinks in the
    taskbar. We work around it with the standard AttachThreadInput trick: attach
    our thread's input to the current foreground thread and the target window's
    thread, which lets SetForegroundWindow actually take. Best-effort and silent
    on any failure. Runs on a short delay so Explorer has time to create the
    window. ctypes only (stdlib) — HWND args are typed so 64-bit handles aren't
    truncated.
    """
    import ctypes, time
    from ctypes import wintypes
    try:
        u = ctypes.windll.user32
        k = ctypes.windll.kernel32
    except Exception:
        return
    u.GetForegroundWindow.restype = wintypes.HWND
    u.SetForegroundWindow.argtypes = [wintypes.HWND]
    u.BringWindowToTop.argtypes = [wintypes.HWND]
    u.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
    u.IsWindowVisible.argtypes = [wintypes.HWND]
    u.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    u.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    u.GetWindowTextLengthW.argtypes = [wintypes.HWND]
    u.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    u.GetWindowThreadProcessId.restype = wintypes.DWORD

    leaf = os.path.basename(os.path.normpath(folder)).lower()  # Explorer titles a window by its leaf folder name
    found = []
    PROTO = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def _cb(hwnd, _):
        if not u.IsWindowVisible(hwnd):
            return True
        cls = ctypes.create_unicode_buffer(64)
        u.GetClassNameW(hwnd, cls, 64)
        if cls.value != "CabinetWClass":   # the file-Explorer window class
            return True
        n = u.GetWindowTextLengthW(hwnd)
        tb = ctypes.create_unicode_buffer(n + 1)
        u.GetWindowTextW(hwnd, tb, n + 1)
        t = tb.value.lower()
        if leaf and (t == leaf or t.endswith(leaf)):
            found.append(hwnd)
            return False
        return True

    cb = PROTO(_cb)
    hwnd = None
    for _ in range(20):                      # ~3s of polling while Explorer opens
        found.clear()
        u.EnumWindows(cb, 0)
        if found:
            hwnd = found[0]
            break
        time.sleep(0.15)
    if not hwnd:
        return
    cur = k.GetCurrentThreadId()
    fg = u.GetForegroundWindow()
    fg_tid = u.GetWindowThreadProcessId(fg, None) if fg else 0
    tgt_tid = u.GetWindowThreadProcessId(hwnd, None)
    u.ShowWindow(hwnd, 9)                     # SW_RESTORE (un-minimize)
    try:
        if fg_tid:
            u.AttachThreadInput(fg_tid, cur, True)
        if tgt_tid:
            u.AttachThreadInput(tgt_tid, cur, True)
        u.BringWindowToTop(hwnd)
        u.SetForegroundWindow(hwnd)
    finally:
        if fg_tid:
            u.AttachThreadInput(fg_tid, cur, False)
        if tgt_tid:
            u.AttachThreadInput(tgt_tid, cur, False)


@app.route("/api/open-folder", methods=["POST"])
def api_open_folder():
    """Open one of the configured watch folders in the OS file browser.
    Restricted to the configured folders so we never open arbitrary paths."""
    import subprocess, threading
    try:
        body = request.get_json(force=True) or {}
    except Exception:
        body = {}
    folder = body.get("folder")
    norm = lambda p: os.path.normcase(os.path.normpath(str(p)))
    allowed = {norm(f) for f in SCANNER.folders}
    if not folder or norm(folder) not in allowed:
        return jsonify({"error": "unknown folder"}), 400
    try:
        if sys.platform.startswith("win"):
            subprocess.Popen(["explorer", folder])
            # Force the new window to the foreground (Windows blocks background
            # processes from doing this directly) — on a daemon thread so the
            # request returns immediately.
            threading.Thread(target=_win_focus_explorer, args=(folder,), daemon=True).start()
        elif sys.platform == "darwin":
            subprocess.Popen(["open", folder])
        else:
            subprocess.Popen(["xdg-open", folder])
        return jsonify({"status": "ok"})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/reimport-storage", methods=["POST"])
def api_reimport_storage():
    """Reset the storage layout from the lab Excel — and ONLY the layout: it
    rewrites assignedStorage + shelf/slot for the active RPT cells the sheet
    lists, then re-packs them. Archive, removedCells, schedules, events, and
    cells the sheet doesn't mention are left untouched.

    If an .xlsx is uploaded (multipart 'file'), it's parsed and also saved as
    the new baseline (storage_excel.json) so future resets use the latest. With
    no upload, the stored baseline is used."""
    upload = request.files.get("file")
    if upload is not None:
        try:
            excel = parse_storage_from_xlsx(upload.read())
        except Exception as e:
            return jsonify({"error": "could not read Excel: " + str(e)}), 400
        if not excel:
            return jsonify({"error": "no 'Cell ID' / 'Assigned Storage Chamber' columns found in an 'RPT Planner' sheet"}), 400
        try:
            with open(STORAGE_EXCEL_PATH, "w", encoding="utf-8") as f:
                json.dump(excel, f, indent=0, sort_keys=True)
        except OSError:
            pass
    else:
        try:
            with open(STORAGE_EXCEL_PATH, encoding="utf-8") as f:
                excel = json.load(f)
        except (OSError, json.JSONDecodeError):
            return jsonify({"error": "no Excel reference (storage_excel.json) found"}), 400
    with _data_lock:
        data = load_data()
        cells = data.get("rptCells", [])
        applied = 0
        for c in cells:
            spot = excel.get(c.get("cellId"))
            if spot and spot.get("chamber"):
                c["assignedStorage"] = spot["chamber"]
                c["storageShelf"] = None
                c["storageSlot"] = None
                applied += 1

        # Pack the just-cleared cells into free slots server-side (the client's
        # auto-pack only runs once per session, so we can't rely on it here).
        def shelves_of(chamber):
            cfg = (data.get("storageConfig") or {}).get(chamber)
            if isinstance(cfg, dict) and isinstance(cfg.get("shelves"), list) and cfg["shelves"]:
                return [s.get("cap", 12) for s in cfg["shelves"]]
            if isinstance(cfg, (int, float)) and cfg > 0:
                return [12] * int(cfg)
            return [12]

        def cell_num(cid):
            m = re.match(r"CRD_(\d+)_CID(\d+)", cid or "")
            return (int(m.group(1)), int(m.group(2))) if m else (10**9, 10**9)

        occ = {}  # chamber -> set of (shelf, slot)
        for c in cells:
            if c.get("assignedStorage") and c.get("storageSlot") is not None:
                occ.setdefault(c["assignedStorage"], set()).add((c.get("storageShelf") or 0, c["storageSlot"]))
        to_place = sorted(
            [c for c in cells if c.get("assignedStorage") and c.get("storageSlot") is None],
            key=lambda c: cell_num(c.get("cellId")))
        for c in to_place:
            taken = occ.setdefault(c["assignedStorage"], set())
            placed = False
            for sh, cap in enumerate(shelves_of(c["assignedStorage"])):
                for sl in range(cap):
                    if (sh, sl) not in taken:
                        c["storageShelf"], c["storageSlot"] = sh, sl
                        taken.add((sh, sl))
                        placed = True
                        break
                if placed:
                    break
            # chambers full -> leaves it unplaced (shows in the tray)

        # Flag (don't touch) active cells the sheet doesn't list — they may be
        # done and want archiving, but that's the user's manual call.
        listed = set(excel.keys())
        unlisted = sorted(c.get("cellId") for c in cells
                          if c.get("cellId") and c["cellId"] not in listed)

        save_data(data)
        # Reset the living seed to the placed cells from this Excel.
        placement = {k: v for k, v in excel.items() if v.get("chamber")}
        try:
            tmp = STORAGE_SEED_PATH.with_suffix(".json.tmp")
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(placement, f, indent=0, sort_keys=True)
            os.replace(tmp, STORAGE_SEED_PATH)
        except OSError:
            pass
        _propagate_to_stable(data, load_metadata())
    return jsonify({"status": "ok", "applied": applied, "unlisted": unlisted, "data": data})


@app.route("/api/unarchive", methods=["POST"])
def api_unarchive():
    """Restore an archived calendar-life cell back to the planner: drop it from
    removedCells (so the watcher manages it again), re-add it to rptCells with
    its saved history + last-known storage spot, and remove it from the archive.
    The next scan recomputes its run-state/schedule from the events."""
    try:
        body = request.get_json(force=True) or {}
    except Exception:
        body = {}
    cid = body.get("cellId")
    if not cid:
        return jsonify({"error": "missing cellId"}), 400
    with _data_lock:
        data = load_data()
        arch = data.get("archivedCells") or {}
        a = arch.get(cid)
        if not a:
            return jsonify({"error": "cell not in archive"}), 404
        data["removedCells"] = [x for x in data.get("removedCells", []) if x != cid]
        cells = data.setdefault("rptCells", [])
        if not any(c.get("cellId") == cid for c in cells):
            spot = load_storage_seed().get(cid) or {}
            evs = a.get("events") or []
            last = evs[-1] if evs else None
            rest = (data.get("rptSettings") or {}).get("restDays", 28)
            nrs = _iso_add(last["returned"], rest) if (last and last.get("returned")) else None
            cells.append({
                "cellId":          cid,
                "storageTemp":     a.get("storageTemp"),
                "soc":             a.get("soc"),
                "rptType":         a.get("rptType") or "Calendar Life RPT",
                "assignedChannel": None,
                "assignedStorage": spot.get("chamber"),
                "storageShelf":    spot.get("shelf"),
                "storageSlot":     spot.get("slot"),
                "anchorDate":      (evs[0].get("start") if evs else None),
                "events":          evs,
                "nextRptStart":    nrs,
            })
        arch.pop(cid, None)
        data["archivedCells"] = arch
        save_data(data)
        _propagate_to_stable(data, load_metadata())
    return jsonify({"status": "ok", "data": data})


@app.route("/api/reset", methods=["POST"])
def api_reset():
    """Hard reset: re-seed data.json and wipe metadata."""
    with _data_lock:
        if DATA_PATH.exists():
            DATA_PATH.unlink()
        if DATA_STABLE_PATH.exists():
            DATA_STABLE_PATH.unlink()
        if META_PATH.exists():
            META_PATH.unlink()
        load_data()  # re-seeds
        _bootstrap_metadata()
    SCANNER.scan()
    return jsonify({"status": "ok"})


# -----------------------------------------------------------------
# Background scanner thread
# -----------------------------------------------------------------

def background_scanner():
    while True:
        try:
            SCANNER.scan()
        except Exception as e:
            print("[ERROR] Scan failed: " + str(e))
        time.sleep(POLL_INTERVAL)


# -----------------------------------------------------------------
# Main
# -----------------------------------------------------------------

if __name__ == "__main__":
    print("")
    print("  CellRD Lab Manager — File-Watcher Edition")
    print("  ==========================================")
    print("  Watch folders:")
    for f in WATCH_FOLDERS:
        print("    " + f)
    print("  Poll:          {}s".format(POLL_INTERVAL))
    print("  Batch gap:     {}min (file gap that ends a batch)".format(BATCH_GAP_MINUTES))
    print("  Updating thresh:{}min (folder considered \"updating\" if any file newer than this)".format(UPDATING_THRESHOLD_MINUTES))
    print("  Stable stale:  {}h (force-advance stable view if frozen longer than this)".format(STABLE_STALE_HOURS))
    print("")
    print("  Rack mapping:")
    for k, v in RACK_MAP.items():
        print("    Neware {} -> {}".format(k, v))
    print("")
    print("  Auto-sync chambers: " + ", ".join(sorted(AUTO_SYNC_CHAMBERS)))
    print("")

    # First-run init
    load_data()
    _bootstrap_metadata()

    # Initial scan — always take a fresh stable snapshot on startup so
    # data_stable.json never carries stale state across server restarts.
    SCANNER.scan()
    with _data_lock:
        save_data(json.loads(DATA_PATH.read_text(encoding="utf-8")), path=DATA_STABLE_PATH)

    # Background thread
    t = threading.Thread(target=background_scanner, daemon=True)
    t.start()

    print("  Open http://localhost:5000")
    print("")

    app.run(host="0.0.0.0", port=5000, debug=False)

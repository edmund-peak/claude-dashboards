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

import json
import os
import re
import shutil
import sys
import threading
import time
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
STATIC_DIR       = ROOT / "static"

AUTO_SYNC_CHAMBERS = {
    "Chamber_1", "Chamber_2", "Chamber_3",
    "Chamber_4", "Chamber_5", "Chamber_6",
    "Chamber_7", "Chamber_8", "Chamber_9",
    "HighTemp",  "CoinCell_1", "CoinCell_2",
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
BATCH_GAP_MINUTES    = CONFIG.get("batch_gap_minutes", 120)
UPDATING_THRESHOLD_MINUTES = CONFIG.get("updating_threshold_minutes", 20)
STABLE_STALE_HOURS   = CONFIG.get("stable_stale_hours", 12)


# -----------------------------------------------------------------
# Filename parsing — same logic as watcher.py
# -----------------------------------------------------------------

FILENAME_RE = re.compile(
    # Optional leading timestamp like "1778605617301_"
    r"(?:\d+_)?"
    # CRD_XX_CIDYYY
    r"(CRD_\d+_CID\d+)"
    # Test info (greedy-back, captures everything up to the IP segment)
    r"_(.+?)"
    # IP: either 127.0.0.1 or 127_0_0_1
    r"_(\d+[._]\d+[._]\d+[._]\d+)"
    # BTSnn
    r"-BTS(\d+)"
    # Rack-Tester-Channel
    r"-(\d+)-(\d+)-(\d+)"
    # Trailing seq (e.g. "399_2" or "145")
    r"-(\d+(?:_\d+)?)"
    r"\.xlsx$",
    re.IGNORECASE,
)
TEMP_RE = re.compile(r"(\d+)C", re.IGNORECASE)
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


def parse_filename(filename):
    m = FILENAME_RE.search(filename)
    if not m:
        return None
    cell_id     = m.group(1)
    test_info   = m.group(2)
    ip_seg      = m.group(3)
    bts_num     = m.group(4)
    rack_num    = m.group(5)
    tester_num  = m.group(6)
    channel_num = m.group(7)
    seq         = m.group(8)

    # When Neware rolls a file over after it gets too large, it appends "_N"
    # to the trailing sequence (e.g. "...-185.xlsx" → "...-185_1.xlsx" →
    # "...-185_2.xlsx"). All these are the same test. Strip the rollover
    # suffix so we can group them by a stable "test_key".
    base_seq = seq.split("_", 1)[0]
    test_key = "{}|{}|{}|{}|{}|{}|{}|{}".format(
        cell_id, test_info, ip_seg, bts_num,
        rack_num, tester_num, channel_num, base_seq)

    rack_code = RACK_MAP.get(rack_num)
    if rack_code:
        channel_id = "{}-T{:02d}-CH{:02d}".format(
            rack_code, int(tester_num), int(channel_num))
    else:
        channel_id = "?{}-T{:02d}-CH{:02d}".format(
            rack_num, int(tester_num), int(channel_num))

    temp_m = TEMP_RE.search(test_info)
    temp = temp_m.group(1) + "C" if temp_m else None

    test_type = "Unknown"
    info_up = test_info.upper()
    for keyword, label in TEST_TYPE_KEYWORDS:
        if keyword in info_up:
            test_type = label
            break

    return {
        "cell_id":     cell_id,
        "test_info":   test_info,
        "test_type":   test_type,
        "temp":        temp,
        "channel_id":  channel_id,
        "test_key":    test_key,
        "filename":    filename,
    }


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
        # On the first scan after server startup, always rebuild stable from
        # live regardless of updating state. This catches stale data_stable.json
        # files left over from previous code versions (or manual resets) so the
        # user doesn't get permanently frozen on a wrong snapshot.
        self.has_bootstrapped = False

    def scan(self):
        """Scan all configured folders, compute live + stable batches,
        and write both to data.json and data_stable.json."""
        # Pass 1: parse all files with their mtimes from every folder
        all_files = []
        folder_stats = []
        for folder in self.folders:
            stat = {"folder": str(folder), "exists": False, "file_count": 0,
                    "error": None}
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
        live_unmapped = self._apply_active(
            live_active, min_ctime_by_test_key, DATA_PATH)

        # Stable view: keep frozen during updates so the user has a calm
        # picture to look at while Neware is rewriting files.
        # Update stable when ANY of the following is true:
        #   - data_stable.json does not exist yet (bootstrap)
        #   - no folder is currently updating (Neware is quiet — safe to advance)
        #   - stable has gone stale beyond STABLE_STALE_HOURS (safeguard against
        #     a runaway updating streak that never lets stable catch up)
        stale_threshold = timedelta(hours=STABLE_STALE_HOURS)
        stable_age = None
        if DATA_STABLE_PATH.exists():
            try:
                stable_age = now - datetime.fromtimestamp(
                    DATA_STABLE_PATH.stat().st_mtime)
            except OSError:
                stable_age = None
        is_first_stable = not DATA_STABLE_PATH.exists()
        is_stable_stale = (stable_age is not None
                           and stable_age > stale_threshold)
        is_first_scan_this_boot = not self.has_bootstrapped

        # Refresh stable from live when any of:
        #   - bootstrap (first scan since process started — guards against
        #     stale data_stable.json left over from a previous code version)
        #   - no data_stable.json exists yet
        #   - no folder is currently updating (Neware is quiet — safe to advance)
        #   - stable has gone stale beyond STABLE_STALE_HOURS
        if (is_first_scan_this_boot or is_first_stable
                or not any_updating or is_stable_stale):
            # Copy live → stable. We don't re-apply by walking files; we just
            # copy the freshly-written data.json content. This is exactly the
            # state we want: stable = "snapshot of live taken when quiet".
            with _data_lock:
                with open(DATA_PATH, encoding="utf-8") as f:
                    live_data = json.load(f)
                save_data(live_data, path=DATA_STABLE_PATH)
            stable_refreshed = True
            if is_first_scan_this_boot:
                self.has_bootstrapped = True
        else:
            stable_refreshed = False

        # Count what's in stable now (whether refreshed or frozen) — read it
        # back so the count is honest about what the user is seeing.
        stable_active_count = 0
        try:
            with open(DATA_STABLE_PATH, encoding="utf-8") as f:
                stable_data = json.load(f)
            for c in stable_data.get("chambers", []):
                if c.get("name") not in AUTO_SYNC_CHAMBERS:
                    continue
                for ch in c.get("channels", []):
                    if ch.get("cellId"):
                        stable_active_count += 1
        except (OSError, json.JSONDecodeError):
            pass

        with self.lock:
            self.last_scan = datetime.now().isoformat()
            self.newest_file_time = newest.isoformat() if newest else None
            self.cutoff_time = cutoff.isoformat() if cutoff else None
            self.total_files = len(all_files)
            live_channels = set(f["channel_id"] for f in live_active)
            self.active_count = len(live_channels)
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

    def _apply_active(self, active, min_ctime_by_test_key, target_path):
        """Apply a list of active files to a fresh copy of the seed data,
        and write it to target_path. Returns the list of unmapped channels."""
        # Pick the most recently-modified file per channel
        by_channel = {}
        for f in active:
            cid = f["channel_id"]
            if cid not in by_channel or f["file_modified"] > by_channel[cid]["file_modified"]:
                by_channel[cid] = f

        with _data_lock:
            # Load the existing target file (preserves user-edited non-auto-sync
            # chambers and RPT data). If it doesn't exist yet, fall back to
            # data.json (so the stable file starts from the same baseline).
            if target_path.exists():
                with open(target_path, encoding="utf-8") as f:
                    data = json.load(f)
            else:
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
                start_ct = min_ctime_by_test_key.get(
                    info["test_key"], info["file_created"])
                ch["startDate"]   = start_ct.strftime("%Y-%m-%d")
                cm = by_cell_meta.get(cell_id, {})
                ch["project"]       = cm.get("project")
                ch["analysisGroup"] = cm.get("analysisGroup")
                ch["cellFormat"]    = cm.get("cellFormat")
                ch["durationWeeks"] = cm.get("durationWeeks")

            save_data(data, path=target_path)
            return unmapped

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
    if view == "stable":
        return jsonify(load_stable_data())
    return jsonify(load_data())


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
        for key in ("rptCells", "rptSettings", "testLibrary"):
            if key in new_data:
                current[key] = new_data[key]

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
        for key in ("rptCells", "rptSettings", "testLibrary"):
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

    # Initial scan
    SCANNER.scan()

    # Background thread
    t = threading.Thread(target=background_scanner, daemon=True)
    t.start()

    print("  Open http://localhost:5000")
    print("")

    app.run(host="0.0.0.0", port=5000, debug=False)

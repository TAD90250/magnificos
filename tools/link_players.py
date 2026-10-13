#!/usr/bin/env python3
"""
Tidy the Players sheet after both platforms' lineups are loaded:

  1. Merge duplicates: a player who first appears on Sleeper (key "S<id>") but has an
     older Yahoo entry (key "Y<id>") under the same name + position becomes ONE player.
     His Weekly_Rosters rows are re-pointed to the Yahoo key and the duplicate row removed.
  2. Fill SleeperPlayerID on Yahoo-era players Sleeper can identify.
  3. Fill PrimaryPosition where it is blank -- from Sleeper's list, else from the lineup
     slot the player was most often started in (QB / RB / WR / TE / K / DST).

Safe to re-run. Also runs automatically at the end of sleeper_backfill.py.

Usage:  python link_players.py
"""
import json
import os
import re
import shutil
from collections import Counter, defaultdict

import schema_io
from ffl_extract import WORKBOOK_PATH

HERE = os.path.dirname(os.path.abspath(__file__))
PLAYERS_CACHE = os.path.join(HERE, "sleeper_players_cache.json")
FIXED_SLOTS = {"QB", "RB", "WR", "TE", "K", "DST"}


def _norm(name):
    s = re.sub(r"\b(jr|sr|ii|iii|iv|v)\b\.?", "", str(name).lower())
    return re.sub(r"[^a-z0-9]", "", s)


def link_players(wb, cache=None):
    if cache is None:
        cache = {}
        if os.path.exists(PLAYERS_CACHE):
            with open(PLAYERS_CACHE, encoding="utf-8") as f:
                cache = json.load(f)
    cache.pop("_missing", None)

    ws = wb["Players"]
    rows = {}
    for r in range(2, ws.max_row + 1):
        key = ws.cell(row=r, column=1).value
        if key:
            rows[key] = r

    # ---- 1. merge Sleeper-only duplicates into their Yahoo twin -----------------
    yahoo_free = defaultdict(list)       # normalized name -> Yahoo rows with no Sleeper id yet
    for key, r in rows.items():
        if str(key).startswith("Y") and not str(key).startswith("Y-") and not ws.cell(row=r, column=5).value:
            yahoo_free[_norm(ws.cell(row=r, column=2).value)].append(key)
    remap = {}                           # S key -> Y key
    for key, r in list(rows.items()):
        if not str(key).startswith("S") or str(key).startswith("S-"):
            continue
        name = _norm(ws.cell(row=r, column=2).value)
        pos = ws.cell(row=r, column=3).value
        cands = yahoo_free.get(name, [])
        if len(cands) > 1 and pos:
            cands = [c for c in cands if ws.cell(row=rows[c], column=3).value in (None, pos)] or cands
        if len(cands) == 1:
            remap[key] = cands[0]
            yahoo_free[name] = []        # each Yahoo row can be claimed once

    if remap:
        for skey, ykey in remap.items():
            ws.cell(row=rows[ykey], column=5).value = ws.cell(row=rows[skey], column=5).value
            if not ws.cell(row=rows[ykey], column=3).value:
                ws.cell(row=rows[ykey], column=3).value = ws.cell(row=rows[skey], column=3).value
        wr = wb["Weekly_Rosters"]
        for r in range(2, wr.max_row + 1):
            pk = wr.cell(row=r, column=5).value
            if pk in remap:
                new = remap[pk]
                wr.cell(row=r, column=5).value = new
                wr.cell(row=r, column=1).value = schema_io.roster_entry_id(
                    wr.cell(row=r, column=2).value, wr.cell(row=r, column=3).value,
                    wr.cell(row=r, column=4).value, new)
        for skey in sorted(remap, key=lambda k: -rows[k]):          # delete bottom-up
            ws.delete_rows(rows[skey])
        rows = {}
        for r in range(2, ws.max_row + 1):
            key = ws.cell(row=r, column=1).value
            if key:
                rows[key] = r
    print(f"  Players: merged {len(remap)} Sleeper-only duplicates into their Yahoo entries")

    # ---- 2. Sleeper ids for Yahoo players ---------------------------------------
    by_yahoo = {}
    for key, r in rows.items():
        y = ws.cell(row=r, column=4).value
        if y:
            by_yahoo[str(y)] = r
    linked = 0
    for sid, info in cache.items():
        yid = info[3]
        if yid and str(yid) in by_yahoo:
            r = by_yahoo[str(yid)]
            if not ws.cell(row=r, column=5).value:
                ws.cell(row=r, column=5).value = sid
                linked += 1
    print(f"  Players: {linked} more Yahoo-era players linked to a Sleeper id")

    # ---- 3. fill missing PrimaryPosition ----------------------------------------
    slot_use = defaultdict(Counter)
    wr = wb["Weekly_Rosters"]
    for row in wr.iter_rows(min_row=2, min_col=5, max_col=7, values_only=True):
        pk, slot, starter = row
        if pk and slot in FIXED_SLOTS:
            slot_use[pk][slot] += 1
    filled = 0
    for key, r in rows.items():
        if ws.cell(row=r, column=3).value:
            continue
        sid = ws.cell(row=r, column=5).value
        pos = None
        if sid and sid in cache and cache[sid][1]:
            pos = "DST" if cache[sid][1] == "DEF" else cache[sid][1]
        elif slot_use.get(key):
            pos = slot_use[key].most_common(1)[0][0]
        if pos:
            ws.cell(row=r, column=3).value = pos
            filled += 1
    blank = sum(1 for r in rows.values() if not ws.cell(row=r, column=3).value)
    print(f"  Players: filled {filled} blank positions ({blank} still blank -- never started in a fixed slot)")
    return wb


def main():
    if os.path.exists(WORKBOOK_PATH):
        bak = WORKBOOK_PATH.replace(".xlsx", "_backup_before_linking.xlsx")
        shutil.copy2(WORKBOOK_PATH, bak)
        print(f"Backup saved: {bak}")
    wb = schema_io.load_workbook(WORKBOOK_PATH)
    link_players(wb)
    wb.save(WORKBOOK_PATH)
    print(f"Saved {WORKBOOK_PATH}")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
FFL Magnificos -- load Yahoo weekly lineups (2016-2025, starters + bench + IR with points)
into FFL_Magnificos_Database.xlsx (Players + Weekly_Rosters), then link Yahoo players to
the Sleeper players already in the sheet (link_players.py).

Input: yahoo_lineups_magnificos.json (this folder), scraped from the Yahoo matchup pages:
  {"players": {pid: [name, " - POS"]},
   "games":   [[year, week, tidA, tidB, [[pid, slot, pts, benchFlag], ...], [...]], ...]}
  benchFlag 0 = starter. Team defenses are "DEF:<city-slug>".

Run load_yahoo_magnificos.py first (Matchups/Teams). Safe to re-run; backs up first.
"""
import json
import os
import shutil
from collections import Counter

import pandas as pd

import schema_io
from ffl_extract import load_owner_map
from league_config import HERE, WORKBOOK_PATH, YAHOO_LEAGUE_KEYS, assert_isolated
from load_yahoo_magnificos import TEAMS_CSV, owner_key
import link_players

LINEUPS_JSON = os.path.join(HERE, "yahoo_lineups_magnificos.json")
DRAFTS_JSON = os.path.join(HERE, "yahoo_drafts_magnificos.json")   # {year: [[round, pick, pid, name, team_name], ...]}
POSITIONS_JSON = os.path.join(HERE, "yahoo_positions_magnificos.json")  # {yahoo_id: "QB"/"RB"/...}, from Sleeper's player list
assert_isolated(LINEUPS_JSON, DRAFTS_JSON, POSITIONS_JSON)
REAL_POS = {"QB", "RB", "WR", "TE", "K"}
POSITIONS = json.load(open(POSITIONS_JSON, encoding="utf-8")) if os.path.exists(POSITIONS_JSON) else {}

SLOT_MAP = {"QB": "QB", "RB": "RB", "WR": "WR", "TE": "TE", "K": "K", "DEF": "DST", "D": "DST", "DST": "DST",
            "W/R/T": "W/R/T FLEX", "W/R": "W/R/T FLEX", "W/T": "W/R/T FLEX", "R/W/T": "W/R/T FLEX",
            "BN": "BN", "IR": "IR", "IR+": "IR", "Q/W/R/T": "SUPERFLEX"}


def team_maps(owner_map):
    teams = pd.read_csv(TEAMS_CSV)
    out = {}
    for year, lid in YAHOO_LEAGUE_KEYS.items():
        sub = teams[(teams["league_id"] == lid) & (teams["season"] == year)]
        out[year] = {int(t["team_id"]): owner_map["yahoo"][owner_key(t)]["owner_id"] for _, t in sub.iterrows()}
    return out


def load_drafts(wb, t2o):
    if not os.path.exists(DRAFTS_JSON):
        print("No yahoo_drafts_magnificos.json -- drafts skipped")
        return
    with open(DRAFTS_JSON, encoding="utf-8") as f:
        drafts = json.load(f)
    teams = pd.read_csv(TEAMS_CSV)
    pos_of = {}
    pws = wb["Players"]
    for r in range(2, pws.max_row + 1):
        if pws.cell(row=r, column=4).value:
            pos_of[str(pws.cell(row=r, column=4).value)] = pws.cell(row=r, column=3).value
    n = 0
    notes_by_id = {}
    for year, picks in drafts.items():
        year = int(year)
        lid = YAHOO_LEAGUE_KEYS[year]
        sub = teams[(teams["league_id"] == lid) & (teams["season"] == year)]
        name_to_tid = {str(t["name"]).strip(): int(t["team_id"]) for _, t in sub.iterrows()}
        for rnd, pick, pid, name, team_name in picks:
            tid = name_to_tid.get(str(team_name).strip())
            if tid is None:
                print(f"  WARNING: {year} R{rnd}P{pick}: team {team_name!r} not found")
                continue
            tsid = schema_io.team_season_id(year, t2o[year][tid])
            note = POSITIONS.get(str(pid)) or pos_of.get(str(pid)) or ("DST" if str(pid).startswith("DEF:") else "")
            did = schema_io.upsert_draft_pick(wb, year, int(rnd), int(pick), tsid, name, notes=note)
            notes_by_id[did] = note
            n += 1
    dws = wb["Draft"]
    for r in range(2, dws.max_row + 1):          # refresh position notes on picks loaded earlier
        did = dws.cell(row=r, column=1).value
        if did in notes_by_id and notes_by_id[did]:
            dws.cell(row=r, column=7, value=notes_by_id[did])
    print(f"Draft picks loaded: {n}")


def main():
    with open(LINEUPS_JSON, encoding="utf-8") as f:
        data = json.load(f)
    players, games = data["players"], data["games"]
    t2o = team_maps(load_owner_map())

    if os.path.exists(WORKBOOK_PATH):
        shutil.copy2(WORKBOOK_PATH, WORKBOOK_PATH.replace(".xlsx", "_backup_before_yahoo_lineups.xlsx"))
    wb = schema_io.load_workbook(WORKBOOK_PATH)

    # ---- Players ----
    pws = wb["Players"]
    have = {}
    dst_rows = []
    for r in range(2, pws.max_row + 1):
        k = pws.cell(row=r, column=1).value
        if k:
            have[k] = r
            if pws.cell(row=r, column=3).value == "DST":
                dst_rows.append((k, str(pws.cell(row=r, column=2).value or "")))

    def def_key(pid):
        """Reuse the Sleeper defense row (e.g. 'Baltimore Ravens') when the nickname or city matches."""
        name = players.get(pid, ["", ""])[0]
        slug = pid.split(":", 1)[1].replace("-", " ")
        for k, full in dst_rows:
            if name and full.lower().endswith(name.lower()):
                return k
        for k, full in dst_rows:
            if slug and full.lower().startswith(slug.lower()):
                return k
        return "Y-" + pid

    keymap = {}
    r = schema_io._next_row(pws)
    added_players = 0
    used = {e[0] for g in games for side in (g[4], g[5]) for e in side}
    for pid in sorted(used):
        key = def_key(pid) if pid.startswith("DEF:") else "Y" + pid
        keymap[pid] = key
        if key in have:
            cur = pws.cell(row=have[key], column=3).value
            if not pid.startswith("DEF:") and cur not in REAL_POS | {"DST"} and POSITIONS.get(pid):
                pws.cell(row=have[key], column=3, value=POSITIONS[pid])
            continue
        name, tp = players.get(pid, [pid, ""])
        pos = "DST" if pid.startswith("DEF:") else (tp.split("-")[-1].strip() if "-" in tp else "")
        if not pid.startswith("DEF:") and (pid in POSITIONS or pos not in REAL_POS):
            pos = POSITIONS.get(pid) or ""   # the scraped label is the lineup slot (BN/IR/W/R/T), not the position
        pws.cell(row=r, column=1, value=key)
        pws.cell(row=r, column=2, value=name)
        pws.cell(row=r, column=3, value=pos or None)
        if not pid.startswith("DEF:"):
            pws.cell(row=r, column=4, value=pid)
        have[key] = r
        r += 1
        added_players += 1
    print(f"Players added: {added_players}")

    # ---- Weekly_Rosters ----
    rws = wb["Weekly_Rosters"]
    existing = {rws.cell(row=x, column=1).value for x in range(2, rws.max_row + 1)}
    r = schema_io._next_row(rws)
    added = 0
    unknown = Counter()
    for year, week, tidA, tidB, sideA, sideB in games:
        for tid, side in ((tidA, sideA), (tidB, sideB)):
            tsid = schema_io.team_season_id(year, t2o[year][tid])
            for pid, slot, pts, bench in side:
                pk = keymap[pid]
                eid = schema_io.roster_entry_id(year, week, tsid, pk)
                if eid in existing:
                    continue
                mapped = SLOT_MAP.get(slot)
                if mapped is None:
                    unknown[slot] += 1
                for c, v in enumerate([eid, year, week, tsid, pk, mapped or slot,
                                       "Yes" if bench == 0 else "No", pts,
                                       None if mapped else f"Yahoo slot {slot}"], start=1):
                    if v is not None:
                        rws.cell(row=r, column=c, value=v)
                existing.add(eid)
                r += 1
                added += 1
    print(f"Weekly_Rosters rows added: {added}")
    if unknown:
        print("Unrecognized slots:", dict(unknown))

    # ---- check: starters add up to the matchup score ----
    score = {}
    for row in wb["Matchups"].iter_rows(min_row=2, values_only=True):
        if row[1] and row[4] and row[6]:
            score[(row[1], row[2], row[4])] = row[5]
            score[(row[1], row[2], row[6])] = row[7]
    bad = 0
    for year, week, tidA, tidB, sideA, sideB in games:
        for tid, side in ((tidA, sideA), (tidB, sideB)):
            tsid = schema_io.team_season_id(year, t2o[year][tid])
            tot = round(sum(e[2] or 0 for e in side if e[3] == 0), 2)
            exp = score.get((year, week, tsid))
            if exp is None or abs(tot - exp) > 0.05:
                bad += 1
                print(f"  WARNING: {tsid} wk{week}: starters {tot} vs matchup {exp}")
    print(f"Starter-total check: {len(games) * 2 - bad} of {len(games) * 2} team-weeks match")

    load_drafts(wb, t2o)
    link_players.link_players(wb)
    wb.save(WORKBOOK_PATH)
    print(f"Saved {WORKBOOK_PATH}")


if __name__ == "__main__":
    main()

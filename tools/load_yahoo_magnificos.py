#!/usr/bin/env python3
"""
FFL Magnificos -- load the Yahoo seasons (2016-2025) into FFL_Magnificos_Database.xlsx.

Inputs (read-only):
  leagues.csv / teams.csv   C:\\Users\\jchap\\Python\\Yahoo\\fantasy_football_parquet
                             (league settings + teams + manager GUIDs, every league on the login;
                             filtered here to the Magnificos keys in league_config.py)
  yahoo_scraped_magnificos.json   (this folder) weekly scores + final standings
                             scraped from the Yahoo site: {year: {rows: [[wk,tidA,sA,tidB,sB],...],
                             info: {tid: [record, final_rank]}}}

Writes Seasons, Owners, Teams (incl. FinalStanding / MadePlayoffs / PlayoffResult) and
Matchups (GameType from replaying the bracket). Safe to re-run; backs up the workbook first.

Usage:
    python load_yahoo_magnificos.py          # all Yahoo seasons
    python load_yahoo_magnificos.py 2023     # one season
"""
import json
import os
import shutil
import sys

import pandas as pd

import schema_io
from ffl_extract import load_owner_map, save_owner_map, owner_id_for
from league_config import HERE, WORKBOOK_PATH, YAHOO_LEAGUE_KEYS, assert_isolated

YAHOO_CSV_DIR = os.environ.get("MAG_CSV_DIR") or r"C:\Users\jchap\Python\Yahoo\fantasy_football_parquet"
LEAGUES_CSV = os.path.join(YAHOO_CSV_DIR, "leagues.csv")
TEAMS_CSV = os.path.join(YAHOO_CSV_DIR, "teams.csv")
SCRAPED = os.path.join(HERE, "yahoo_scraped_magnificos.json")
assert_isolated(SCRAPED, WORKBOOK_PATH)

warnings = []


def warn(msg):
    warnings.append(msg)
    print("  WARNING:", msg)


def owner_key(t):
    """Yahoo manager GUID; hidden managers (GUID '--') get one key per team."""
    g = str(t["manager_guid"])
    return f"hidden:{t['team_key']}" if g in ("--", "nan", "") else g


def owner_label(t):
    m = str(t["manager"])
    return t["name"] if m in ("--hidden--", "nan", "") else m


def classify_season(year, rows, info, playoff_start, n_playoff):
    """Replay the playoff bracket. Returns (game_types, results, alive_set)
    game_types: {row_index: GameType}; results: {tid: PlayoffResult string}"""
    playoff_teams = {t for t, (_, rank) in info.items() if rank <= n_playoff}
    alive = set(playoff_teams)
    semi_losers = set()
    gtype = {}
    result = {}
    champion = runner_up = third_winner = third_loser = None

    weeks = sorted({r[0] for r in rows})
    for w in weeks:
        idx = [i for i, r in enumerate(rows) if r[0] == w]
        if w < playoff_start:
            for i in idx:
                gtype[i] = "Regular"
            continue
        na = len(alive)
        stage = "QF" if na >= 5 else ("SF" if na >= 3 else "CH")
        out_now, win_now = set(), set()
        for i in idx:
            _, a, sa, b, sb = rows[i]
            if a in alive and b in alive:
                if sa == sb:
                    warn(f"{year} wk{w}: playoff game {a} v {b} tied {sa}; assuming lower final rank number advances")
                    win, lose = (a, b) if info[a][1] < info[b][1] else (b, a)
                else:
                    win, lose = (a, b) if sa > sb else (b, a)
                win_now.add(win)
                out_now.add(lose)
                if stage == "QF":
                    gtype[i] = "Playoff - Quarterfinal"
                    result[lose] = "Lost Wildcard/1st Round"
                elif stage == "SF":
                    gtype[i] = "Playoff - Semifinal"
                    semi_losers.add(lose)
                else:
                    gtype[i] = "Playoff - Championship"
                    champion, runner_up = win, lose
            elif a in semi_losers and b in semi_losers and na == 2:
                gtype[i] = "Playoff - 3rd Place"
                win, lose = (a, b) if sa > sb else (b, a)
                third_winner, third_loser = win, lose
            else:
                gtype[i] = "Consolation"
        alive -= out_now
    for t in semi_losers:
        result[t] = "Lost Semifinal"
    if champion is not None:
        result[champion] = "Champion"
        result[runner_up] = "Runner-Up"
    if third_winner is not None:
        result[third_winner] = "Won 3rd Place Game"
        result[third_loser] = "Lost 3rd Place Game"
    for t in playoff_teams:
        result.setdefault(t, "Other")
    for t in info:
        if t not in playoff_teams:
            result[t] = "Missed Playoffs"

    # --- sanity checks against Yahoo's own final ranks ---
    rank_of = {t: v[1] for t, v in info.items()}
    if champion is None:
        warn(f"{year}: no championship game found")
    else:
        if rank_of[champion] != 1:
            warn(f"{year}: simulated champion tid {champion} but Yahoo rank is {rank_of[champion]}")
        if rank_of[runner_up] != 2:
            warn(f"{year}: simulated runner-up tid {runner_up} but Yahoo rank is {rank_of[runner_up]}")
    if third_winner is not None and rank_of[third_winner] != 3:
        warn(f"{year}: 3rd-place winner tid {third_winner} has Yahoo rank {rank_of[third_winner]}")
    return gtype, result, playoff_teams



def main():
    only = {int(a) for a in sys.argv[1:]} if len(sys.argv) > 1 else None
    with open(SCRAPED, encoding="utf-8") as f:
        scraped = json.load(f)
    leagues = pd.read_csv(LEAGUES_CSV)
    teams = pd.read_csv(TEAMS_CSV)
    owner_map = load_owner_map()

    if os.path.exists(WORKBOOK_PATH):
        bak = WORKBOOK_PATH.replace(".xlsx", "_backup_before_yahoo.xlsx")
        shutil.copy2(WORKBOOK_PATH, bak)
        print(f"Backup saved: {bak}")
    wb = schema_io.load_workbook(WORKBOOK_PATH)

    for year in sorted(YAHOO_LEAGUE_KEYS):
        if only and year not in only:
            continue
        lid = YAHOO_LEAGUE_KEYS[year]
        lrow = leagues[(leagues["league_id"] == lid) & (leagues["season"] == year)]
        if lrow.empty:
            warn(f"{year}: {lid} not in leagues.csv, skipped")
            continue
        lrow = lrow.iloc[0]
        start_week = int(lrow["start_week"]) if pd.notna(lrow["start_week"]) else 1
        end_week = int(lrow["end_week"]) if pd.notna(lrow["end_week"]) else 16
        playoff_start = int(lrow["playoff_start_week"])
        n_playoff = int(lrow["num_playoff_teams"])
        trow = teams[(teams["league_id"] == lid) & (teams["season"] == year)]
        print(f"== {year}: {lrow['name']!r} ({lid}), {len(trow)} teams ==")
        schema_io.upsert_season(wb, year, "Yahoo", lrow["name"], lid, len(trow),
                                 f"{start_week}-{playoff_start - 1}", playoff_start,
                                 f"{playoff_start}-{end_week}", n_playoff, "", "No",
                                 notes=f"Draft: {lrow['draft_type']}" if pd.notna(lrow.get("draft_type")) else "")

        t2o = {}
        for _, t in trow.iterrows():
            label = owner_label(t)
            oid = owner_id_for(owner_map, "yahoo", owner_key(t), label)
            schema_io.upsert_owner(wb, oid, label, label)
            schema_io.upsert_team(wb, year, oid, t["name"])
            t2o[int(t["team_id"])] = oid

        d = scraped.get(str(year))
        if not d:
            warn(f"{year}: no scraped scores yet")
            save_owner_map(owner_map); wb.save(WORKBOOK_PATH)
            continue
        raw_rows = d["rows"]
        if isinstance(raw_rows, str):
            raw_rows = [r.split(",") for r in raw_rows.split(";") if r]
        rows = [(int(r[0]), int(r[1]), float(r[2]), int(r[3]), float(r[4])) for r in raw_rows]
        info = {int(k): (v[0], int(v[1])) for k, v in d["info"].items()}
        bad = [r for r in rows if r[1] not in t2o or r[3] not in t2o]
        for r in bad:
            warn(f"{year} wk{r[0]}: team id not in teams.csv: {r}")
        rows = [r for r in rows if r not in bad]

        gtype, results, playoff_teams = classify_season(year, rows, info, playoff_start, n_playoff)

        rec = {t: [0, 0, 0] for t in info}
        for i, (w, a, sa, b, sb) in enumerate(rows):
            if gtype[i] != "Regular":
                continue
            if sa > sb:
                rec[a][0] += 1; rec[b][1] += 1
            elif sb > sa:
                rec[b][0] += 1; rec[a][1] += 1
            else:
                rec[a][2] += 1; rec[b][2] += 1
        for t, (record, _) in info.items():
            if "-".join(map(str, rec[t])) != record:
                warn(f"{year} tid {t} ({t2o[t]}): games give {'-'.join(map(str, rec[t]))}, Yahoo says {record}")

        for i, (w, a, sa, b, sb) in enumerate(rows):
            schema_io.upsert_matchup(wb, year, w, gtype[i], schema_io.team_season_id(year, t2o[a]), sa,
                                     schema_io.team_season_id(year, t2o[b]), sb)
        for t, (record, rank) in info.items():
            schema_io.upsert_team(wb, year, t2o[t], "", final_standing=rank,
                                  made_playoffs="Yes" if t in playoff_teams else "No",
                                  playoff_result=results[t])
        champ = [t2o[t] for t, r in results.items() if r == "Champion"]
        print(f"  {len(rows)} games ({sum(1 for g in gtype.values() if g.startswith('Playoff'))} playoff); "
              f"champion: {champ[0] if champ else '?'}")
        save_owner_map(owner_map)
        wb.save(WORKBOOK_PATH)

    schema_io.rebuild_team_weeks(wb)
    wb.save(WORKBOOK_PATH)
    print(f"\nSaved {WORKBOOK_PATH} -- {len(warnings)} warning(s)")


if __name__ == "__main__":
    main()

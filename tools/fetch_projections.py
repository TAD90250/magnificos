#!/usr/bin/env python3
"""
Pull Sleeper's projected points for every week still to be played, plus each team's CURRENT roster,
and save them to power_projections.json. export_site_data.py reads that file to build the
Power Rankings page (best projected lineup per team per remaining week, ranked, points added up).

Run it every Thursday, before the new week starts:

    python fetch_projections.py
    python export_site_data.py

Needs internet (Sleeper's public API, no login). Played weeks are never included: the file starts at
the current NFL week. Safe to re-run; it just rewrites the file.

Which projection number is used:
  Sleeper publishes ready-made projected fantasy points in three scoring flavours (pts_ppr,
  pts_half_ppr, pts_std). The one matching your league's points-per-reception is used
  automatically. Custom scoring tweaks beyond that (e.g. unusual kicker or defense rules) are not
  reflected. To force one, set PROJECTION_FIELD below.
"""
import json
import os
import sys
import time
from datetime import datetime, timezone

import requests

import sleeper_client
from ffl_extract import SLEEPER_SEED_LEAGUE_ID, load_owner_map

from league_config import HERE, assert_isolated   # FFL Magnificos folder only
OUT = os.path.join(HERE, "power_projections.json")
assert_isolated(OUT)
PROJECTION_FIELD = None          # e.g. "pts_ppr" | "pts_half_ppr" | "pts_std"; None = pick from the league's PPR setting
LAST_WEEK = None                 # None = work it out from the league settings (normally 17)
POSITIONS = ["QB", "RB", "WR", "TE", "K", "DEF"]


def get(url, params=None, tries=3):
    for i in range(tries):
        try:
            r = requests.get(url, params=params, timeout=40)
            r.raise_for_status()
            return r.json()
        except Exception as e:                                   # noqa: BLE001
            if i == tries - 1:
                raise
            print(f"   retrying ({e})")
            time.sleep(2)


def main():
    league = sleeper_client.get_league(SLEEPER_SEED_LEAGUE_ID)
    season = int(league["season"])
    settings = league.get("settings", {})
    scoring = league.get("scoring_settings", {})
    rec = scoring.get("rec", 0)
    field = PROJECTION_FIELD or ("pts_ppr" if rec >= 1 else "pts_half_ppr" if rec >= 0.5 else "pts_std")

    state = get("https://api.sleeper.app/v1/state/nfl")
    first_week = max(1, int(state.get("week") or settings.get("leg") or 1))
    if str(state.get("season")) == str(season) and state.get("season_type") == "pre":
        first_week = 1
    ps = settings.get("playoff_week_start")
    last_week = LAST_WEEK or ((ps - 1 + 3) if ps else 17)
    print(f"{season}: projecting weeks {first_week}-{last_week} using {field}")

    # ---- current rosters, tied to the same owner ids the workbook uses ----
    owner_map = load_owner_map()
    teams = {}
    roster_owner = {}                                            # Sleeper roster_id -> owner id (for this week's matchups)
    for ro in sleeper_client.get_rosters(SLEEPER_SEED_LEAGUE_ID):
        uid = ro.get("owner_id")
        info = owner_map.get("sleeper", {}).get(uid)
        if not info:
            print(f"   WARNING: roster {ro['roster_id']} (user {uid}) is not in owner_map.json -- skipped. "
                  f"Run sleeper_backfill.py weekly first.")
            continue
        out_of_play = set(ro.get("reserve") or []) | set(ro.get("taxi") or [])     # IR / taxi cannot start
        roster_owner[ro["roster_id"]] = info["owner_id"]
        teams[info["owner_id"]] = [p for p in (ro.get("players") or []) if p not in out_of_play]
    needed = {p for ps_ in teams.values() for p in ps_}
    print(f"   {len(teams)} rosters, {len(needed)} players")

    # ---- projections, one request per week ----
    players, weeks = {}, {}
    for w in range(first_week, last_week + 1):
        url = f"https://api.sleeper.com/projections/nfl/{season}/{w}"
        data = get(url, params=[("season_type", "regular"), ("order_by", field)] + [("position[]", p) for p in POSITIONS])
        got = {}
        for e in data or []:
            pid = str(e.get("player_id"))
            if pid not in needed:
                continue
            pts = (e.get("stats") or {}).get(field)
            if pts is None:
                continue
            got[pid] = round(float(pts), 2)
            pl = e.get("player") or {}
            pos = pl.get("position") or ""
            players[pid] = {"n": (f"{pl.get('first_name', '')} {pl.get('last_name', '')}").strip() or pid,
                            "p": "DST" if pos in ("DEF", "DST") else pos}
        weeks[str(w)] = got
        print(f"   week {w}: projections for {len(got)} of your rostered players")
        time.sleep(0.4)

    if not any(weeks.values()):
        sys.exit("No projections came back -- nothing written. Check your internet connection and try again.")
    # ---- who plays whom this week (for the home page "This week's matchups") ----
    matchups = None
    try:
        mrows = sleeper_client.get_matchups(SLEEPER_SEED_LEAGUE_ID, first_week)
        by_id = {}
        for m in mrows or []:
            if m.get("matchup_id") is not None and m.get("roster_id") in roster_owner:
                by_id.setdefault(m["matchup_id"], []).append(roster_owner[m["roster_id"]])
        pairs = [v for _, v in sorted(by_id.items()) if len(v) == 2]
        if pairs:
            matchups = {"week": first_week, "pairs": pairs}
            print(f"   week {first_week}: {len(pairs)} matchups")
        else:
            print(f"   week {first_week}: Sleeper has no matchups posted yet")
    except Exception as e:                                       # noqa: BLE001
        print(f"   could not read this week's matchups ({e}) -- the home page section will be skipped")
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({"season": season, "fetchedAt": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
                   "field": field, "firstWeek": first_week, "lastWeek": last_week,
                   "rosterPositions": league.get("roster_positions") or [],
                   "teams": teams, "players": players, "weeks": weeks, "matchups": matchups}, f)
    print(f"Saved {OUT}")


if __name__ == "__main__":
    main()

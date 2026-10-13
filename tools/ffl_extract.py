#!/usr/bin/env python3
"""
Pull all FFL Magnificos history (Yahoo seasons + Sleeper seasons) into
FFL_Magnificos_Database.xlsx. League IDs live in league_config.py.

Usage:
    python ffl_extract.py discover      # read-only: list every season/owner
                                         # found, seed owner_map.json. Run
                                         # this FIRST, edit owner_map.json,
                                         # then run `full`.
    python ffl_extract.py full          # pull everything, write the workbook
    python ffl_extract.py full 2024     # pull just one Yahoo season (test run,
                                         # skips Sleeper) -- handy to confirm
                                         # Yahoo access works before doing all 10
    python ffl_extract.py weekly        # pull only the current season's
                                         # newest week(s) (run this weekly
                                         # during the season)

Safe to re-run: every row is keyed by an ID, existing IDs are skipped.
"""
import json
import os
import re
import sys
import time

import schema_io
import sleeper_client
from yahoo_client import YahooClient, find_all, merge_meta_list, numbered_items

from league_config import (HERE, OAUTH_PATH, WORKBOOK_PATH, OWNER_MAP_PATH, DEBUG_DIR,
                           SLEEPER_SEED_LEAGUE_ID, YAHOO_LEAGUE_KEYS, require_league_ids)


# ---------------------------------------------------------------- owner map
def load_owner_map():
    if os.path.exists(OWNER_MAP_PATH):
        with open(OWNER_MAP_PATH, encoding="utf-8") as f:
            return json.load(f)
    return {"yahoo": {}, "sleeper": {}}


def save_owner_map(m):
    with open(OWNER_MAP_PATH, "w", encoding="utf-8") as f:
        json.dump(m, f, indent=2)


def slugify(name, used):
    base = re.sub(r"[^A-Za-z]", "", name or "OWNER").upper()[:8] or "OWNER"
    candidate = base
    i = 1
    while candidate in used:
        i += 1
        candidate = f"{base}{i}"
    used.add(candidate)
    return candidate


def owner_id_for(owner_map, platform, external_id, display_name):
    table = owner_map[platform]
    if external_id in table:
        return table[external_id]["owner_id"]
    used = {v["owner_id"] for p in owner_map.values() for v in p.values()}
    oid = slugify(display_name, used)
    table[external_id] = {"owner_id": oid, "display_name": display_name}
    return oid


def dump_debug(name, data):
    os.makedirs(DEBUG_DIR, exist_ok=True)
    with open(os.path.join(DEBUG_DIR, f"{name}.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


# --------------------------------------------------------------- Yahoo pull
def yahoo_all_leagues(yc):
    """Every NFL league this Yahoo login has played in (all leagues, all years)."""
    data = yc.get("users;use_login=1/games;game_codes=nfl/leagues")
    dump_debug("yahoo_discover_raw", data)
    return [lg for lg in find_all(data, "league_key") if lg.get("season") and lg.get("league_key")]


def yahoo_discover_leagues(yc):
    """{year: league_meta} for ONLY the Magnificos keys in league_config.YAHOO_LEAGUE_KEYS.
    Any other league on the login (Camaradas included) is ignored."""
    wanted = {str(k): int(y) for y, k in YAHOO_LEAGUE_KEYS.items()}
    out = {}
    for league in yahoo_all_leagues(yc):
        lk = league.get("league_key")
        if lk in wanted:
            if int(league["season"]) != wanted[lk]:
                print(f"  WARNING: {lk} is season {league['season']}, config says {wanted[lk]}")
            out[wanted[lk]] = league
    for lk, y in wanted.items():
        if y not in out:
            print(f"  WARNING: configured {y}: {lk} was not found on this Yahoo login")
    return out


def classify_yahoo_playoff_week(week, playoff_start_week, last_week):
    if playoff_start_week is None or week < playoff_start_week:
        return "Regular"
    rounds_from_end = last_week - week
    if rounds_from_end == 0:
        return "Playoff - Championship"
    if rounds_from_end == 1:
        return "Playoff - Semifinal"
    return "Playoff - Quarterfinal"


def pull_yahoo_season(wb, yc, owner_map, year, league_meta, debug=False):
    league_key = league_meta["league_key"]
    league_name = league_meta.get("name", f"Yahoo {year}")
    num_teams = int(league_meta.get("num_teams", 0) or 0)
    scoring_raw = (league_meta.get("scoring_type") or "head").lower()
    scoring_type = "PPR" if "ppr" in scoring_raw and "half" not in scoring_raw else (
        "Half-PPR" if "half" in scoring_raw else "Standard")
    start_week = int(league_meta.get("start_week", 1) or 1)
    end_week = int(league_meta.get("end_week", 16) or 16)

    settings_data = yc.get(f"league/{league_key}/settings")
    if debug:
        dump_debug(f"yahoo_{year}_settings", settings_data)
    settings = next(find_all(settings_data, "playoff_start_week"), {})
    playoff_start_week = settings.get("playoff_start_week")
    playoff_start_week = int(playoff_start_week) if playoff_start_week else None
    num_playoff_teams = settings.get("num_playoff_teams", "")
    reg_weeks_str = f"{start_week}-{(playoff_start_week - 1) if playoff_start_week else end_week}"
    playoff_weeks_str = f"{playoff_start_week}-{end_week}" if playoff_start_week else ""

    teams_data = yc.get(f"league/{league_key}/teams")
    if debug:
        dump_debug(f"yahoo_{year}_teams", teams_data)
    teams_container = next(find_all(teams_data, "teams"))["teams"]
    team_key_to_tsid = {}
    for item in numbered_items(teams_container):
        meta = merge_meta_list(item["team"][0])
        team_key = meta.get("team_key")
        team_name = meta.get("name", "")
        managers = meta.get("managers", [])
        manager = managers[0]["manager"] if managers else {}
        guid = manager.get("guid") or f"unknown-{team_key}"
        nickname = manager.get("nickname", "Unknown")
        owner_id = owner_id_for(owner_map, "yahoo", guid, nickname)
        schema_io.upsert_owner(wb, owner_id, nickname, nickname)
        tsid = schema_io.upsert_team(wb, year, owner_id, team_name)
        team_key_to_tsid[team_key] = tsid
        print(f"    team {team_name!r} -> owner {owner_id} ({nickname})")

    schema_io.upsert_season(wb, year, "Yahoo", league_name, league_key, num_teams,
                             reg_weeks_str, playoff_start_week or "", playoff_weeks_str,
                             num_playoff_teams, scoring_type, "No")

    seen_pairs = set()
    for week in range(start_week, end_week + 1):
        try:
            sb = yc.get(f"league/{league_key}/scoreboard", params={"week": week})
        except Exception as e:
            print(f"    week {week}: scoreboard fetch failed ({e}), skipping")
            continue
        if debug and week == start_week:
            dump_debug(f"yahoo_{year}_scoreboard_w{week}", sb)
        matchups_container = next(find_all(sb, "matchups"), {}).get("matchups", {})
        for item in numbered_items(matchups_container):
            raw = item.get("matchup")
            if raw is None:
                continue
            m = merge_meta_list(raw) if isinstance(raw, list) else raw
            is_consolation = str(m.get("is_consolation", "0")) == "1"
            is_playoffs = str(m.get("is_playoffs", "0")) == "1"
            teams_in_match = m.get("teams", {})
            scores = []
            for t_item in numbered_items(teams_in_match):
                tmeta = merge_meta_list(t_item["team"][0])
                tk = tmeta.get("team_key")
                pts = tmeta.get("team_points", {}).get("total", 0)
                scores.append((team_key_to_tsid.get(tk, tk), float(pts or 0)))
            if len(scores) != 2:
                continue
            (a, sa), (b, sb_) = scores
            pair_key = (week, tuple(sorted([a, b])))
            if pair_key in seen_pairs:
                continue
            seen_pairs.add(pair_key)
            if is_consolation:
                game_type = "Consolation"
            elif is_playoffs:
                game_type = classify_yahoo_playoff_week(week, playoff_start_week, end_week)
            else:
                game_type = "Regular"
            schema_io.upsert_matchup(wb, year, week, game_type, a, sa, b, sb_)
        time.sleep(0.3)
    print(f"  Yahoo {year}: {league_name!r} — {len(team_key_to_tsid)} teams loaded")


# ------------------------------------------------------------- Sleeper pull
def classify_sleeper_week(week, playoff_week_start, bracket_place_by_week):
    if playoff_week_start is None or week < playoff_week_start:
        return "Regular"
    place = bracket_place_by_week.get(week)
    if place == 1:
        return "Playoff - Championship"
    if place == 3:
        return "Playoff - 3rd Place"
    return "Playoff - Semifinal"


def pull_sleeper_season(wb, owner_map, year, league_id, debug=False):
    league = sleeper_client.get_league(league_id)
    if debug:
        dump_debug(f"sleeper_{year}_league", league)
    league_name = league.get("name", f"Sleeper {year}")
    num_teams = league.get("total_rosters", 0)
    settings = league.get("settings", {})
    playoff_week_start = settings.get("playoff_week_start")
    # "leg"/"last_scored_leg" means "how far the season has gotten" -- for a
    # season still in progress that's the CURRENT week, not the final one.
    # Use it only to decide how far to actually fetch (no point requesting
    # unplayed future weeks); for the season-length metadata text, always
    # use the typical 3-round-playoff estimate so an in-progress season
    # doesn't show a nonsensical "15-3" range.
    estimated_season_length = (playoff_week_start - 1 + 3) if playoff_week_start else 17
    fetch_through_week = settings.get("last_scored_leg") or settings.get("leg") or estimated_season_length
    last_week = estimated_season_length
    scoring = league.get("scoring_settings", {})
    ppr = scoring.get("rec", 0)
    scoring_type = "PPR" if ppr >= 1 else ("Half-PPR" if ppr >= 0.5 else "Standard")

    users = sleeper_client.get_users(league_id)
    rosters = sleeper_client.get_rosters(league_id)
    if debug:
        dump_debug(f"sleeper_{year}_users", users)
        dump_debug(f"sleeper_{year}_rosters", rosters)
    user_by_id = {u["user_id"]: u for u in users}
    roster_to_tsid = {}
    for roster in rosters:
        uid = roster.get("owner_id")
        user = user_by_id.get(uid, {})
        display_name = user.get("display_name", f"user-{uid}")
        team_name = (user.get("metadata") or {}).get("team_name") or display_name
        owner_id = owner_id_for(owner_map, "sleeper", uid or f"roster-{roster['roster_id']}", display_name)
        schema_io.upsert_owner(wb, owner_id, display_name, display_name)
        tsid = schema_io.upsert_team(wb, year, owner_id, team_name)
        roster_to_tsid[roster["roster_id"]] = tsid
        print(f"    roster {team_name!r} -> owner {owner_id} ({display_name})")

    schema_io.upsert_season(wb, year, "Sleeper", league_name, league_id, num_teams,
                             f"1-{(playoff_week_start - 1) if playoff_week_start else ''}",
                             playoff_week_start or "",
                             f"{playoff_week_start}-{last_week}" if playoff_week_start else "",
                             settings.get("playoff_teams", ""), scoring_type, "No")

    bracket_place_by_week = {}
    for entry in sleeper_client.get_winners_bracket(league_id):
        if entry.get("p") and entry.get("w"):
            bracket_place_by_week[entry.get("r", 0) + (playoff_week_start or 1) - 1] = entry["p"]

    seen_pairs = set()
    for week in range(1, int(fetch_through_week) + 1):
        try:
            matchups = sleeper_client.get_matchups(league_id, week)
        except Exception as e:
            print(f"    week {week}: matchups fetch failed ({e}), skipping")
            continue
        if not matchups:
            continue
        if debug and week == 1:
            dump_debug(f"sleeper_{year}_matchups_w1", matchups)
        by_matchup_id = {}
        for entry in matchups:
            mid = entry.get("matchup_id")
            if mid is None:
                continue
            by_matchup_id.setdefault(mid, []).append(entry)
        game_type = classify_sleeper_week(week, playoff_week_start, bracket_place_by_week)
        for mid, pair in by_matchup_id.items():
            if len(pair) != 2:
                continue
            a_entry, b_entry = pair
            a = roster_to_tsid.get(a_entry["roster_id"])
            b = roster_to_tsid.get(b_entry["roster_id"])
            if not a or not b:
                continue
            pair_key = (week, tuple(sorted([a, b])))
            if pair_key in seen_pairs:
                continue
            seen_pairs.add(pair_key)
            schema_io.upsert_matchup(wb, year, week, game_type, a, a_entry.get("points", 0),
                                      b, b_entry.get("points", 0))
    print(f"  Sleeper {year}: {league_name!r} — {len(roster_to_tsid)} teams loaded")


# --------------------------------------------------------------------- main
def run_discover():
    owner_map = load_owner_map()
    print("=== Yahoo ===")
    try:
        yc = YahooClient(OAUTH_PATH)
        print("  token refreshed OK, trying baseline call (users;use_login=1)...")
        base = yc.get("users;use_login=1")
        print("  baseline call OK:", json.dumps(base)[:200])
        configured = {str(k) for k in YAHOO_LEAGUE_KEYS.values()}
        print("  All Yahoo leagues on this login ( * = in league_config.YAHOO_LEAGUE_KEYS ):")
        for lm in sorted(yahoo_all_leagues(yc), key=lambda x: (int(x["season"]), x.get("name") or "")):
            mark = "*" if lm["league_key"] in configured else " "
            print(f"  {mark} {lm['season']}: {lm.get('name')!r}  (league_key={lm['league_key']}, "
                  f"{lm.get('num_teams')} teams)")
        if not configured:
            print("  -> Copy the Magnificos rows into YAHOO_LEAGUE_KEYS in league_config.py.")
    except Exception as e:
        print(f"  Yahoo discovery FAILED: {e}")

    print("=== Sleeper ===")
    if not SLEEPER_SEED_LEAGUE_ID:
        print("  SLEEPER_SEED_LEAGUE_ID is empty in league_config.py -- skipping Sleeper.")
    else:
      try:
        chain = sleeper_client.discover_chain(SLEEPER_SEED_LEAGUE_ID)
        for season, lid in sorted(chain.items()):
            league = sleeper_client.get_league(lid)
            print(f"  {season}: {league.get('name')!r}  (league_id={lid}, "
                  f"{league.get('total_rosters')} teams)")
      except Exception as e:
        print(f"  Sleeper discovery FAILED: {e}")

    save_owner_map(owner_map)
    print(f"\nowner_map.json written to {OWNER_MAP_PATH}")
    print("Review/edit owner IDs there if you want nicer codes, then run: python ffl_extract.py full")


def run_full(debug_first=True, year_filter=None):
    require_league_ids(need_yahoo=True, need_sleeper=year_filter is None)
    owner_map = load_owner_map()
    wb = schema_io.load_workbook(WORKBOOK_PATH)

    print("=== Yahoo seasons ===")
    try:
        yc = YahooClient(OAUTH_PATH)
        leagues = yahoo_discover_leagues(yc)
    except Exception as e:
        print(f"Yahoo discovery FAILED, skipping all Yahoo seasons: {e}")
        leagues = {}
    if year_filter is not None:
        leagues = {y: lm for y, lm in leagues.items() if y == year_filter}
        print(f"(--year {year_filter}: restricting Yahoo pull to this one season, skipping Sleeper)")
    for i, year in enumerate(sorted(leagues)):
        print(f" -- {year} --")
        try:
            pull_yahoo_season(wb, yc, owner_map, year, leagues[year], debug=(debug_first and i == 0))
        except Exception as e:
            print(f"    FAILED on {year}: {e}")
        save_owner_map(owner_map)
        wb.save(WORKBOOK_PATH)

    if year_filter is None:
        print("=== Sleeper seasons ===")
        try:
            chain = sleeper_client.discover_chain(SLEEPER_SEED_LEAGUE_ID)
        except Exception as e:
            print(f"Sleeper discovery FAILED, skipping all Sleeper seasons: {e}")
            chain = {}
        for i, (season, lid) in enumerate(sorted(chain.items())):
            print(f" -- {season} --")
            try:
                pull_sleeper_season(wb, owner_map, int(season), lid, debug=(debug_first and i == 0))
            except Exception as e:
                print(f"    FAILED on {season}: {e}")
            save_owner_map(owner_map)
            wb.save(WORKBOOK_PATH)

    print("Rebuilding Team_Weeks view...")
    schema_io.rebuild_team_weeks(wb)
    wb.save(WORKBOOK_PATH)
    print(f"Done. Open {WORKBOOK_PATH} in Excel or upload it to Google Sheets.")


def run_weekly():
    """Refresh just the current (highest-year) Sleeper season's matchups."""
    require_league_ids(need_yahoo=False, need_sleeper=True)
    owner_map = load_owner_map()
    wb = schema_io.load_workbook(WORKBOOK_PATH)
    chain = sleeper_client.discover_chain(SLEEPER_SEED_LEAGUE_ID)
    if not chain:
        print("No Sleeper season found.")
        return
    season = max(chain, key=int)
    print(f"Refreshing current season: Sleeper {season}")
    pull_sleeper_season(wb, owner_map, int(season), chain[season])
    save_owner_map(owner_map)
    schema_io.rebuild_team_weeks(wb)
    wb.save(WORKBOOK_PATH)
    print(f"Done. Updated {WORKBOOK_PATH}")


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "discover"
    if mode == "discover":
        run_discover()
    elif mode == "full":
        year_filter = int(sys.argv[2]) if len(sys.argv) > 2 else None
        run_full(year_filter=year_filter)
    elif mode == "weekly":
        run_weekly()
    else:
        print(__doc__)

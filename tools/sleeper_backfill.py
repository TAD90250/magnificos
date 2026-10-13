#!/usr/bin/env python3
"""
Fill in everything the basic Sleeper pull (ffl_extract.py) leaves out, for every
Sleeper season (2025, 2026, ...):

  * Matchup GameType   -- real playoff rounds / 3rd place / consolation, from the
                          playoff bracket (the basic pull labels every playoff-week
                          game "Semifinal")
  * Teams              -- FinalStanding, MadePlayoffs, PlayoffResult (finished seasons)
  * Draft              -- every pick of the Sleeper draft(s)
  * Players            -- Sleeper players, linked to the Yahoo-era players already in
                          the sheet (Sleeper publishes each player's Yahoo id), and the
                          Sleeper id is filled in for Yahoo players who also appear there
  * Weekly_Rosters     -- starters + bench with fantasy points, every scored week

Usage (venv active, in the FFL_Magnificos folder):
    python sleeper_backfill.py            # all Sleeper seasons
    python sleeper_backfill.py weekly     # just the current season (use during 2026)

Safe to re-run. Makes a backup copy of the workbook first. Needs internet.
Sleeper asks apps to download its big player list at most once a day; this script
saves a trimmed copy (sleeper_players_cache.json) and only re-downloads when it
meets a player it doesn't have.
"""
import json
import os
import shutil
import sys

import requests

import schema_io
import sleeper_client
from ffl_extract import (WORKBOOK_PATH, SLEEPER_SEED_LEAGUE_ID, load_owner_map, save_owner_map,
                         owner_id_for, dump_debug, pull_sleeper_season)

HERE = os.path.dirname(os.path.abspath(__file__))
PLAYERS_CACHE = os.path.join(HERE, "sleeper_players_cache.json")
FANTASY_POS = {"QB", "RB", "WR", "TE", "K", "DEF"}

SLOT_MAP = {"QB": "QB", "RB": "RB", "WR": "WR", "TE": "TE", "K": "K", "DEF": "DST",
            "FLEX": "W/R/T FLEX", "WRRB_FLEX": "RB/WR FLEX", "REC_FLEX": "RB/WR FLEX",
            "SUPER_FLEX": "SUPERFLEX", "BN": "BN", "IR": "IR"}


# --------------------------------------------------------------- helpers
def api(path):
    return sleeper_client._get(path)


def resolve_bracket(entries):
    """Sleeper bracket entries reference earlier games like {"w": 1}; turn them into roster ids."""
    byid = {e["m"]: e for e in entries}

    def res(t):
        if isinstance(t, dict):
            if "w" in t:
                return byid.get(t["w"], {}).get("w")
            if "l" in t:
                return byid.get(t["l"], {}).get("l")
            return None
        return t

    out = []
    for e in sorted(entries, key=lambda e: (e.get("r", 0), e.get("m", 0))):
        out.append({"r": e.get("r"), "m": e.get("m"), "t1": res(e.get("t1")), "t2": res(e.get("t2")),
                    "w": e.get("w"), "l": e.get("l"), "p": e.get("p")})
    return out


def set_game_type(wb, year, week, ta, tb, game_type):
    ws = wb["Matchups"]
    r = schema_io._find_row(ws, lambda row: ws.cell(row=row, column=2).value == year
                            and ws.cell(row=row, column=3).value == week
                            and {ws.cell(row=row, column=5).value, ws.cell(row=row, column=7).value} == {ta, tb})
    if r:
        ws.cell(row=r, column=4).value = game_type
        return True
    return False


def classify_pair(bracket, rnd, a, b):
    if not bracket:
        return None
    maxr = max(e["r"] for e in bracket if e["r"])
    for e in bracket:
        if e["r"] == rnd and {e["t1"], e["t2"]} == {a, b}:
            if e["p"] == 1:
                return "Playoff - Championship"
            if e["p"] == 3:
                return "Playoff - 3rd Place"
            if e["p"] and e["p"] > 3:
                return "Consolation"
            if rnd == maxr - 1:
                return "Playoff - Semifinal"
            if rnd <= maxr - 2:
                return "Playoff - Quarterfinal"
            return "Playoff - Championship"
    return "Consolation"


def load_players_cache():
    if os.path.exists(PLAYERS_CACHE):
        with open(PLAYERS_CACHE, encoding="utf-8") as f:
            return json.load(f)
    return {}


def refresh_players_cache():
    print("  downloading Sleeper's NFL player list (large, once)...")
    r = requests.get(f"{sleeper_client.BASE}/players/nfl", timeout=180)
    r.raise_for_status()
    raw = r.json()
    cache = {}
    for pid, p in raw.items():
        pos = p.get("position")
        if pos not in FANTASY_POS:
            continue
        name = p.get("full_name") or f"{p.get('first_name', '')} {p.get('last_name', '')}".strip()
        cache[pid] = [name, pos, p.get("team"), p.get("yahoo_id"), p.get("last_name")]
    with open(PLAYERS_CACHE, "w", encoding="utf-8") as f:
        json.dump(cache, f, separators=(",", ":"))
    print(f"  cached {len(cache)} fantasy-position players")
    return cache


# --------------------------------------------------------------- per-season work
def season_data(wb, owner_map, year, lid):
    league = sleeper_client.get_league(lid)
    users = {u["user_id"]: u for u in sleeper_client.get_users(lid)}
    rosters = sleeper_client.get_rosters(lid)
    rid_to_tsid, rid_to_owner = {}, {}
    for ro in rosters:
        uid = ro.get("owner_id")
        disp = users.get(uid, {}).get("display_name", f"user-{uid}")
        oid = owner_id_for(owner_map, "sleeper", uid or f"roster-{ro['roster_id']}", disp)
        rid_to_tsid[ro["roster_id"]] = schema_io.team_season_id(year, oid)
        rid_to_owner[ro["roster_id"]] = oid
    st = league.get("settings", {})
    pstart = st.get("playoff_week_start") or 15
    through = int(st.get("last_scored_leg") or st.get("leg") or 0)
    weekly = {}
    for w in range(1, through + 1):
        try:
            m = sleeper_client.get_matchups(lid, w)
        except Exception as e:
            print(f"    week {w}: fetch failed ({e})")
            continue
        if m:
            weekly[w] = m
    bracket = resolve_bracket(sleeper_client.get_winners_bracket(lid))
    return dict(year=year, lid=lid, league=league, rosters=rosters, rid_to_tsid=rid_to_tsid,
                rid_to_owner=rid_to_owner, pstart=pstart, weekly=weekly, bracket=bracket)


def fix_game_types(wb, sd):
    year, bracket, pstart = sd["year"], sd["bracket"], sd["pstart"]
    n = 0
    for week, entries in sd["weekly"].items():
        if week < pstart:
            continue
        rnd = week - pstart + 1
        pairs = {}
        for e in entries:
            if e.get("matchup_id") is not None:
                pairs.setdefault(e["matchup_id"], []).append(e["roster_id"])
        for mid, rids in pairs.items():
            if len(rids) != 2:
                continue
            gt = classify_pair(bracket, rnd, rids[0], rids[1]) if bracket else "Playoff - Semifinal"
            if gt and set_game_type(wb, year, week, sd["rid_to_tsid"][rids[0]], sd["rid_to_tsid"][rids[1]], gt):
                n += 1
    print(f"  {year}: game types set for {n} playoff-week games")


def fix_standings(wb, sd):
    year, league, bracket = sd["year"], sd["league"], sd["bracket"]
    if league.get("status") != "complete" or not bracket:
        print(f"  {year}: season not finished (status {league.get('status')}); standings left blank")
        return
    ros = {r["roster_id"]: r for r in sd["rosters"]}

    def key(rid):
        s = ros[rid].get("settings", {})
        return (-(s.get("wins") or 0), -((s.get("fpts") or 0) + (s.get("fpts_decimal") or 0) / 100))

    playoff = {t for e in bracket for t in (e["t1"], e["t2"]) if isinstance(t, int)}
    place = {}
    for e in bracket:
        if e["p"] and e["w"] and e["l"]:
            place[e["w"]] = e["p"]
            place[e["l"]] = e["p"] + 1
    nxt = 1
    for rid in sorted((r for r in playoff if r not in place), key=key):  # unplaced playoff teams
        while nxt in place.values():
            nxt += 1
        place[rid] = nxt
    taken = set(place.values())
    n = 1
    for rid in sorted((r for r in ros if r not in place), key=key):
        while n in taken:
            n += 1
        place[rid] = n
        taken.add(n)
    results = {1: "Champion", 2: "Runner-Up", 3: "Won 3rd Place Game", 4: "Lost 3rd Place Game"}
    for rid, p in sorted(place.items(), key=lambda kv: kv[1]):
        if rid in playoff:
            res = results.get(p, "Lost Wildcard/1st Round")
        else:
            res = "Missed Playoffs"
        schema_io.upsert_team(wb, year, sd["rid_to_owner"][rid], "", final_standing=p,
                              made_playoffs="Yes" if rid in playoff else "No", playoff_result=res)
    champ = [rid for rid, p in place.items() if p == 1]
    print(f"  {year}: standings written; champion = {sd['rid_to_owner'][champ[0]] if champ else '?'}")


def load_draft(wb, sd):
    year, lid = sd["year"], sd["lid"]
    try:
        drafts = api(f"league/{lid}/drafts")
    except Exception as e:
        print(f"  {year}: drafts fetch failed ({e})")
        return
    nteams = len(sd["rosters"])
    total = 0
    for i, d in enumerate(drafts):
        try:
            picks = api(f"draft/{d['draft_id']}/picks")
        except Exception as e:
            print(f"  {year}: picks fetch failed ({e})")
            continue
        if i == 0 and picks:
            dump_debug(f"sleeper_{year}_draft_picks_sample", picks[:3])
        if i > 0:
            print(f"  {year}: extra draft {d.get('draft_id')} ({d.get('type')}) -- skipped, Draft sheet has one slot per round/pick")
            continue
        for p in picks:
            md = p.get("metadata") or {}
            name = f"{md.get('first_name', '')} {md.get('last_name', '')}".strip() or str(p.get("player_id"))
            rnd = p["round"]
            pick_in_round = p["pick_no"] - (rnd - 1) * nteams
            tsid = sd["rid_to_tsid"].get(p.get("roster_id"))
            note = " ".join(x for x in [md.get("position"), md.get("team"),
                                        "(keeper)" if p.get("is_keeper") else ""] if x)
            if tsid:
                schema_io.upsert_draft_pick(wb, year, rnd, pick_in_round, tsid, name, note)
                total += 1
    print(f"  {year}: {total} draft picks")


# --------------------------------------------------------------- players + lineups
def ensure_players(wb, needed_ids):
    cache = load_players_cache()
    unknown_before = set(cache.get("_missing", []))
    missing = [p for p in needed_ids if p not in cache and p not in unknown_before]
    if not cache or missing:
        cache = refresh_players_cache()
        still = sorted(p for p in needed_ids if p not in cache)
        if still:
            print(f"  {len(still)} players are not in Sleeper's list (kept by id only): {still[:10]}")
            cache["_missing"] = still
            with open(PLAYERS_CACHE, "w", encoding="utf-8") as f:
                json.dump(cache, f, separators=(",", ":"))
    cache = {k: v for k, v in cache.items() if k != "_missing"}

    ws = wb["Players"]
    by_key, by_yahoo, def_by_name = {}, {}, {}
    for r in range(2, ws.max_row + 1):
        key = ws.cell(row=r, column=1).value
        if not key:
            continue
        by_key[key] = r
        y = ws.cell(row=r, column=4).value
        if y:
            by_yahoo[str(y)] = r
        if ws.cell(row=r, column=3).value == "DST":
            def_by_name[str(ws.cell(row=r, column=2).value)] = r

    linked = 0
    for sid, (name, pos, team, yid, last) in cache.items():
        if yid and str(yid) in by_yahoo:
            r = by_yahoo[str(yid)]
            if not ws.cell(row=r, column=5).value:
                ws.cell(row=r, column=5).value = sid
                linked += 1
    print(f"  linked {linked} existing Yahoo-era players to their Sleeper ids")

    key_for = {}
    added = 0
    for sid in sorted(needed_ids):
        info = cache.get(sid)
        name, pos, team, yid, last = info if info else (sid, "", "", None, None)
        row = None
        if pos == "DEF":
            row = def_by_name.get(str(last))
            key = ws.cell(row=row, column=1).value if row else f"S-DEF:{sid}"
        elif yid and str(yid) in by_yahoo:
            row = by_yahoo[str(yid)]
            key = ws.cell(row=row, column=1).value
        elif yid:
            key = "Y" + str(yid)
        else:
            key = "S" + sid
        key_for[sid] = key
        if row is None and key in by_key:
            row = by_key[key]
        if row is None:
            row = schema_io._next_row(ws)
            ws.cell(row=row, column=1, value=key)
            ws.cell(row=row, column=2, value=name)
            ws.cell(row=row, column=3, value="DST" if pos == "DEF" else (pos or None))
            if yid and not pos == "DEF":
                ws.cell(row=row, column=4, value=str(yid))
            by_key[key] = row
            if yid:
                by_yahoo[str(yid)] = row
            added += 1
        if not ws.cell(row=row, column=5).value:
            ws.cell(row=row, column=5).value = sid
    print(f"  Players: {added} new rows added")
    return key_for


def load_lineups(wb, sd, key_for):
    year = sd["year"]
    slots = [p for p in sd["league"].get("roster_positions", []) if p != "BN"]
    ws = wb["Weekly_Rosters"]
    index = {ws.cell(row=r, column=1).value: r for r in range(2, ws.max_row + 1) if ws.cell(row=r, column=1).value}
    state = {"next": schema_io._next_row(ws)}
    for week, entries in sd["weekly"].items():
        for e in entries:
            tsid = sd["rid_to_tsid"].get(e.get("roster_id"))
            if not tsid:
                continue
            starters = e.get("starters") or []
            pts = e.get("players_points") or {}
            for i, pid in enumerate(starters):
                if pid in (None, "", "0"):
                    continue
                _upsert_roster(ws, index, state, year, week, tsid, key_for[pid],
                               SLOT_MAP.get(slots[i], slots[i]) if i < len(slots) else "BN", "Yes", pts.get(pid))
            for pid in e.get("players") or []:
                if pid in starters:
                    continue
                _upsert_roster(ws, index, state, year, week, tsid, key_for[pid], "BN", "No", pts.get(pid))
    new_rows = len(index)
    print(f"  {year}: Weekly_Rosters now holds {new_rows} rows (all years)")


def _upsert_roster(ws, index, state, year, week, tsid, pkey, slot, starter, points):
    eid = schema_io.roster_entry_id(year, week, tsid, pkey)
    r = index.get(eid)
    if r is None:
        r = state["next"]
        state["next"] += 1
        ws.cell(row=r, column=1, value=eid)
        index[eid] = r
    ws.cell(row=r, column=2, value=year)
    ws.cell(row=r, column=3, value=week)
    ws.cell(row=r, column=4, value=tsid)
    ws.cell(row=r, column=5, value=pkey)
    ws.cell(row=r, column=6, value=slot)
    ws.cell(row=r, column=7, value=starter)
    ws.cell(row=r, column=8, value=points)


# --------------------------------------------------------------- main
def main():
    weekly_only = len(sys.argv) > 1 and sys.argv[1] == "weekly"
    owner_map = load_owner_map()
    chain = sleeper_client.discover_chain(SLEEPER_SEED_LEAGUE_ID)
    seasons = sorted(chain.items(), key=lambda kv: int(kv[0]))
    if weekly_only:
        seasons = seasons[-1:]
    print("Sleeper seasons:", [s for s, _ in seasons])

    if os.path.exists(WORKBOOK_PATH):
        bak = WORKBOOK_PATH.replace(".xlsx", "_backup_before_sleeper.xlsx")
        shutil.copy2(WORKBOOK_PATH, bak)
        print(f"Backup saved: {bak}")
    wb = schema_io.load_workbook(WORKBOOK_PATH)

    sds = []
    for season, lid in seasons:
        year = int(season)
        print(f"== Sleeper {year} ==")
        pull_sleeper_season(wb, owner_map, year, lid)          # refresh teams + matchup scores
        sd = season_data(wb, owner_map, year, lid)
        fix_game_types(wb, sd)
        fix_standings(wb, sd)
        if not weekly_only:
            load_draft(wb, sd)
        sds.append(sd)

    needed = set()
    for sd in sds:
        for entries in sd["weekly"].values():
            for e in entries:
                for pid in (e.get("players") or []) + (e.get("starters") or []):
                    if pid not in (None, "", "0"):
                        needed.add(pid)
    print(f"== Players / lineups ({len(needed)} distinct Sleeper players) ==")
    key_for = ensure_players(wb, needed)
    for sd in sds:
        load_lineups(wb, sd, key_for)

    import link_players
    link_players.link_players(wb)
    save_owner_map(owner_map)
    schema_io.rebuild_team_weeks(wb)
    wb.save(WORKBOOK_PATH)
    print(f"Saved {WORKBOOK_PATH}")


if __name__ == "__main__":
    main()

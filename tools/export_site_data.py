#!/usr/bin/env python3
"""
Turns FFL_Magnificos_Database.xlsx into the small JSON files the league website reads.

    python export_site_data.py                     # workbook from ffl_extract.WORKBOOK_PATH, output to site/data
    python export_site_data.py --workbook X.xlsx --out some/folder

Everything is computed here from the raw input columns (Matchups, Teams, Draft,
Weekly_Rosters, Players), so the workbook does NOT need to be opened in Excel first.
Writes (all UTF-8 JSON):

    meta.json        league name, last update, current season / week, owner list for the menus
    owners.json      all-time leaderboard
    seasons.json     one entry per year: standings, champion, runner-up
    records.json     league-wide records (best/worst scores, blowouts, player games ...)
    h2h.json         owner-vs-owner matrix
    draft.json       every recorded draft pick, by year
    power.json       best-possible-lineup power rankings for the current season
    owner/<slug>.json   one file per owner (career, seasons, head-to-head, records, top players)

Safe to re-run any time. Run it after build_owner_sheets.py / sleeper_backfill.py weekly.
"""
import argparse
import json
import os
import re
from collections import Counter, defaultdict
from datetime import datetime, timezone

import openpyxl

import build_owner_sheets as bos
from build_owner_sheets import OWNER_NAMES, load_tables, owner_of, player_stats

HERE = os.path.dirname(os.path.abspath(__file__))

from league_config import LEAGUE_NAME, assert_isolated
SHOW_FULL_NAMES = False          # True = also publish owners' full real names on the site
DEFAULT_RECORDS_SCOPE = "regpo"     # record book opens on: reg | regpo | all
ACTIVITY_WEEKS = 1                # home-page "records broken" box covers this many of the latest game weeks
MIN_STARTS_FOR_AVG = 20            # record-book average-points-per-start list needs this many starts
TOP_N = 5                       # how many players to list in each "top players" block


# ------------------------------------------------------------------ helpers
def slugify(text):
    s = re.sub(r"[^a-z0-9]+", "-", str(text).lower()).strip("-")
    return s or "owner"


def rnd(x, n=2):
    return None if x is None else round(float(x), n)


def played(g):
    _, _, _, _, sa, _, sb = g
    return sa is not None and sb is not None and (sa > 0 or sb > 0)


def pct(w, l, t):
    n = w + l + t
    return round((w + 0.5 * t) / n, 4) if n else None


def write_json(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, separators=(",", ":"))


# ------------------------------------------------------------------ extra sheets
def load_extra(wb):
    team_rows = {}                    # tsid -> dict
    for r in wb["Teams"].iter_rows(min_row=2, values_only=True):
        if r[1] and r[2]:
            team_rows[f"{r[1]}-{r[2]}"] = {
                "year": r[1], "owner": r[2], "name": (r[3] or "").strip(),
                "finish": r[11] if isinstance(r[11], (int, float)) else None,
                "made": r[12] == "Yes", "result": r[13] or ""}
    seasons = {}
    for r in wb["Seasons"].iter_rows(min_row=2, values_only=True):
        if r[1]:
            reg = str(r[6] or "")
            m = re.match(r"\s*(\d+)\s*-\s*(\d+)", reg)
            seasons[r[1]] = {"year": r[1], "platform": r[2], "leagueName": r[3], "numTeams": r[5],
                             "regWeeks": int(m.group(2)) if m else None, "playoffStart": r[7],
                             "numPlayoffTeams": r[9], "scoring": r[10]}
    draft = []
    for r in wb["Draft"].iter_rows(min_row=2, values_only=True):
        if r[0] and r[1] and r[4]:
            draft.append((r[1], r[2], r[3], r[4], r[5]))
    return team_rows, seasons, draft


# ------------------------------------------------------------------ main build
PROJECTIONS_FILE = "power_projections.json"   # written by fetch_projections.py
SETTINGS_FILE = "league_settings.json"
RULES_FILE = "league_rules.md"
DEFAULT_SETTINGS = {
    "_help": ("League money settings. confirmed: true hides the 'placeholder' notice. seasonPayouts are the end-of-season prizes. "
              "weeklyPayouts is the schedule of weekly prizes (kind says how the winner is found; manual ones are entered in manualWinners). "
              "lowestPenalty/lowestCount: the lowest N scorers of each regular-season week pay that amount. fees: each owner's base fee and "
              "how much they have paid in so far (Total = base + penalties; Owe = Total - paid; a negative Owe means they overpaid). "
              "powerWeights: how much each part counts in the Power Rankings score (projected lineup, bench depth, record and standings, points scored), separately for weeks 1-14 (early) and the playoffs (late); each set is scaled to add up to 100. lowestOverrides: use this list instead of the computed lowest scorers for a week. penalties: extra fees, e.g. a hurt or bye-week player left in the lineup (amount is what the owner pays). "
              "payments: prize money the league has already paid out to an owner (amount > 0). "
              "toiletBowlWinner: e.g. {\"2026\": \"chaps\"} once it is decided."),
    "confirmed": False,
    "moneyStartYear": None,
    "seasonPayouts": {"first": 0, "second": 0, "third": 0, "regSeasonFirst": 0, "mostPoints": 0},
    "lowestPenalty": 0,
    "lowestCount": 0,
    "missPlayoffsPenalty": 0,
    "powerWeights": {"early": {"lineup": 75, "bench": 5, "record": 15, "points": 5},
                     "late": {"lineup": 100, "bench": 0, "record": 0, "points": 0}},
    "feeNote": "",
    "weeklyPayouts": [],
    "manualWinners": [],
    "lowestOverrides": [],
    "penalties": [],
    "fees": [],
    "toiletBowlWinner": {},
    "payments": []}

DEFAULT_RULES = """# Rules
Paste your league rules into league_rules.md (the file this page is built from).
Start a line with "# " for each section title, start a line with "- " for a bullet point, and write ordinary lines for paragraphs.
Then run export_site_data.py and upload the data folder.

# Example section
- Replace this text with your own rules.
- Each # heading becomes its own section with a link at the top of the page.
"""


def load_settings(latest_year):
    path = os.path.join(HERE, SETTINGS_FILE)
    if not os.path.exists(path):
        d = dict(DEFAULT_SETTINGS)
        d["moneyStartYear"] = latest_year
        with open(path, "w", encoding="utf-8") as f:
            json.dump(d, f, indent=2, ensure_ascii=False)
        print(f"Created {SETTINGS_FILE} with placeholder amounts -- edit it to set your real prizes.")
    with open(path, encoding="utf-8") as f:
        d = json.load(f)
    for k, v in DEFAULT_SETTINGS.items():
        d.setdefault(k, v)
    if not d.get("moneyStartYear"):
        d["moneyStartYear"] = latest_year
    return d


def load_rules(path):
    """league_rules.md: '# Title' starts a section, '- ' lines are bullets, other lines are paragraphs."""
    if not os.path.exists(path):
        with open(path, "w", encoding="utf-8") as f:
            f.write(DEFAULT_RULES)
        print(f"Created {RULES_FILE} -- paste your rules into it.")
    sections, cur = [], None
    for raw in open(path, encoding="utf-8").read().splitlines():
        line = raw.rstrip()
        if line.startswith("# "):
            cur = {"title": line[2:].strip(), "blocks": []}
            sections.append(cur)
        elif not line.strip():
            continue
        else:
            if cur is None:
                cur = {"title": "", "blocks": []}
                sections.append(cur)
            if line.lstrip().startswith(("- ", "* ")):
                item = line.lstrip()[2:].strip()
                if cur["blocks"] and cur["blocks"][-1]["t"] == "ul":
                    cur["blocks"][-1]["items"].append(item)
                else:
                    cur["blocks"].append({"t": "ul", "items": [item]})
            elif line.startswith("## "):
                cur["blocks"].append({"t": "h", "text": line[3:].strip()})
            else:
                cur["blocks"].append({"t": "p", "text": line.strip()})
    placeholder = os.path.getsize(path) == len(DEFAULT_RULES.encode("utf-8")) and open(path, encoding="utf-8").read() == DEFAULT_RULES
    return {"placeholder": placeholder, "sections": sections}


def build(workbook, out_dir):
    wb = openpyxl.load_workbook(workbook, read_only=True)
    owners, teams, games, players, roster = load_tables(wb)
    team_rows, season_info, draft = load_extra(wb)
    wb.close()

    # names: the real-name table wins over whatever the Owners sheet currently holds
    for oid, (full, nick) in OWNER_NAMES.items():
        if oid in owners:
            owners[oid] = {"full": full, "nick": nick}
    used_slugs, slug = set(), {}
    for oid in sorted(owners):
        s = base = slugify(owners[oid]["nick"])
        n = 2
        while s in used_slugs:
            s, n = f"{base}-{n}", n + 1
        used_slugs.add(s)
        slug[oid] = s
    nick = lambda oid: owners.get(oid, {}).get("nick", oid)

    games = [g for g in games if played(g)]
    played_owner = {v[1] for v in teams.values()}

    # ---- team-game rows (one per team per game) ----
    rows = []                                    # dicts: owner, opp, year, week, type, pf, pa, res
    for y, w, gt, a, sa, b, sb in games:
        for me, op, pf, pa in ((a, b, sa, sb), (b, a, sb, sa)):
            res = "W" if pf > pa else "L" if pf < pa else "T"
            rows.append({"o": owner_of(me), "p": owner_of(op), "y": y, "w": w, "t": gt,
                         "pf": pf, "pa": pa, "r": res, "ts": me})
    by_owner = defaultdict(list)
    for r in rows:
        by_owner[r["o"]].append(r)

    def game_obj(r):
        return {"score": rnd(r["pf"]), "oppScore": rnd(r["pa"]), "year": r["y"], "week": r["w"],
                "type": r["t"], "opp": nick(r["p"]), "oppSlug": slug.get(r["p"])}

    # ---- per-team-season regular-season record ----
    reg = defaultdict(lambda: [0, 0, 0, 0.0, 0.0])      # tsid -> w,l,t,pf,pa
    for r in rows:
        if r["t"] == "Regular":
            x = reg[r["ts"]]
            x["WLT".index(r["r"])] += 1
            x[3] += r["pf"]
            x[4] += r["pa"]

    # ---- draft ----
    draft_by_year = defaultdict(list)
    r1 = {}                                              # tsid -> (pick, player)
    for y, rd, pk, tsid, player in draft:
        draft_by_year[y].append((rd, pk, tsid, player))
        if rd == 1:
            r1[tsid] = (pk, player)

    # the 3rd-place game is only counted under "All games"
    scope_defs = [
        ("reg", "Regular season", lambda t: t == "Regular"),
        ("regpo", "Regular season + playoffs", lambda t: t == "Regular" or (t.startswith("Playoff") and "3rd" not in t)),
        ("po", "Playoffs only", lambda t: t.startswith("Playoff") and "3rd" not in t),
        ("all", "All games (incl. consolation)", lambda t: True)]
    # minimums per scope: (games to be ranked on rate tables, games for a season record, starts for a player average)
    # playoffs-only has far fewer games, so its bars are lower
    SCOPE_MINS = {"reg": (14, 8, MIN_STARTS_FOR_AVG), "regpo": (14, 8, MIN_STARTS_FOR_AVG),
                  "po": (4, 2, 5), "all": (14, 8, MIN_STARTS_FOR_AVG)}

    # ---- player stats (one pass per game scope) ----
    per_by_scope = {}
    for skey, _lbl, sfn in scope_defs:
        sg = [g for g in games if sfn(g[2])]
        if skey == "all":
            sroster = roster
        else:
            ok = {(y, w, t) for y, w, gt, a_, sa, b_, sb in sg for t in (a_, b_)}
            sroster = [r for r in roster if (r[0], r[1], r[2]) in ok]
        per_by_scope[skey] = player_stats(owners, teams, sg, players, sroster)

    def pinfo(pid):
        nm, pos = players.get(pid, (pid, None))
        return nm or pid, pos or ""

    # ---- owner ordering (titles, then win %) ----
    tot = {o: [0, 0, 0] for o in owners}
    for r in rows:
        tot[r["o"]]["WLT".index(r["r"])] += 1
    titles = Counter(v["owner"] for v in team_rows.values() if v["result"] == "Champion")
    order = sorted((o for o in owners if o in played_owner),
                   key=lambda o: (-titles[o], -(pct(*tot[o]) or 0)))

    # ============ owner files (every figure is computed once per game scope) ============
    leaderboard_by_scope = {k: [] for k, _, _ in scope_defs}
    for oid in order:
        all_mine = by_owner[oid]
        my_ts = sorted((ts for ts, v in team_rows.items() if v["owner"] == oid), key=lambda t: team_rows[t]["year"])
        scopes_out = []
        for skey, slabel, sfn in scope_defs:
            mine = [r for r in all_mine if sfn(r["t"])]
            per = per_by_scope[skey]
            sreg = defaultdict(lambda: [0, 0, 0, 0.0, 0.0])           # tsid -> w,l,t,pf,pa inside this scope
            for r in mine:
                x = sreg[r["ts"]]
                x["WLT".index(r["r"])] += 1
                x[3] += r["pf"]
                x[4] += r["pa"]
            seasons_out = []
            for ts in my_ts:
                v, x = team_rows[ts], sreg.get(ts, [0, 0, 0, 0.0, 0.0])
                pick = r1.get(ts)
                seasons_out.append({
                    "year": v["year"], "team": v["name"], "w": x[0], "l": x[1], "t": x[2],
                    "games": x[0] + x[1] + x[2],
                    "pf": rnd(x[3], 1), "pa": rnd(x[4], 1), "finish": v["finish"], "result": v["result"],
                    "madePlayoffs": v["made"], "champion": v["result"] == "Champion",
                    "draftR1": {"pick": pick[0], "player": pick[1]} if pick else None})
            wlt_all = [sum(1 for r in mine if r["r"] == k) for k in "WLT"]
            wlt_reg = [sum(1 for r in mine if r["r"] == k and r["t"] == "Regular") for k in "WLT"]
            wlt_po = [sum(1 for r in mine if r["r"] == k and r["t"].startswith("Playoff")) for k in "WLT"]
            pf = sum(r["pf"] for r in mine)
            pa = sum(r["pa"] for r in mine)
            finishes = [s_["finish"] for s_ in seasons_out if s_["finish"]]
            career = {
                "seasons": len(seasons_out), "firstYear": seasons_out[0]["year"], "lastYear": seasons_out[-1]["year"],
                "titles": titles[oid], "titleYears": [s_["year"] for s_ in seasons_out if s_["champion"]],
                "runnerUps": sum(1 for s_ in seasons_out if s_["result"] == "Runner-Up"),
                "playoffApps": sum(1 for s_ in seasons_out if s_["madePlayoffs"]),
                "games": len(mine),
                "all": wlt_all, "allPct": pct(*wlt_all), "reg": wlt_reg, "regPct": pct(*wlt_reg), "playoff": wlt_po,
                "pf": rnd(pf, 1), "pa": rnd(pa, 1), "avgPf": rnd(pf / len(mine), 1) if mine else None,
                "bestFinish": min(finishes) if finishes else None, "worstFinish": max(finishes) if finishes else None}

            # head to head
            h = defaultdict(lambda: {"w": 0, "l": 0, "t": 0, "pf": 0.0, "pa": 0.0, "po": [0, 0, 0]})
            for r in mine:
                x = h[r["p"]]
                x[r["r"].lower()] += 1
                x["pf"] += r["pf"]
                x["pa"] += r["pa"]
                if r["t"].startswith("Playoff"):
                    x["po"]["WLT".index(r["r"])] += 1
            h2h = sorted(({"slug": slug[p], "nick": nick(p), "w": x["w"], "l": x["l"], "t": x["t"],
                           "games": x["w"] + x["l"] + x["t"], "pct": pct(x["w"], x["l"], x["t"]),
                           "pf": rnd(x["pf"], 1), "pa": rnd(x["pa"], 1), "playoff": x["po"]}
                          for p, x in h.items() if p != oid), key=lambda d: (-d["games"], d["nick"]))

            # personal records
            valid = [r for r in mine if r["pf"] > 0 and r["pa"] > 0]
            rec = {}
            if mine:
                rec["highest"] = game_obj(max(mine, key=lambda r: r["pf"]))
                rec["lowest"] = game_obj(min(valid, key=lambda r: r["pf"])) if valid else None
                rec["biggestWin"] = game_obj(max(valid, key=lambda r: r["pf"] - r["pa"])) if valid else None
                rec["biggestLoss"] = game_obj(min(valid, key=lambda r: r["pf"] - r["pa"])) if valid else None

            # top players
            d = per.get(oid)
            top_all, top_year = {}, []
            if d and d["roster"]:
                def line(pid, **kw):
                    nm, pos = pinfo(pid)
                    span = sorted(d["years"][pid])
                    return {"name": nm, "pos": pos, "years": [span[0], span[-1]],
                            "starts": d["starts"][pid], "games": d["roster"][pid],
                            "pts": rnd(d["pts"][pid], 1), **kw}
                top_all["mostStarted"] = [line(p) for p, _ in sorted(
                    d["starts"].items(), key=lambda kv: (-kv[1], -d["pts"][kv[0]]))[:TOP_N]]
                top_all["mostRoster"] = [line(p) for p, _ in sorted(
                    d["roster"].items(), key=lambda kv: (-kv[1], -d["starts"][kv[0]]))[:TOP_N]]
                top_all["mostPoints"] = [line(p) for p, _ in sorted(
                    d["pts"].items(), key=lambda kv: -kv[1])[:TOP_N]]
                best = []
                for pts, pid, y, w, o2 in sorted(d["best"], key=lambda g: -g[0])[:TOP_N]:
                    nm, pos = pinfo(pid)
                    best.append({"name": nm, "pos": pos, "pts": rnd(pts, 1), "year": y, "week": w,
                                 "opp": nick(o2) if o2 else ""})
                top_all["bestGames"] = best
                for y in sorted(d["y_starts"]):
                    if not d["y_starts"][y]:
                        continue
                    ms = sorted(d["y_starts"][y].items(), key=lambda kv: (-kv[1], -d["y_pts"][y][kv[0]]))[0]
                    mp = max(d["y_pts"][y].items(), key=lambda kv: kv[1]) if d["y_pts"][y] else None
                    bg = max(d["y_best"][y], key=lambda g: g[0]) if d["y_best"][y] else None
                    top_year.append({
                        "year": y,
                        "mostStarted": {"name": pinfo(ms[0])[0], "pos": pinfo(ms[0])[1], "starts": ms[1]},
                        "mostPoints": {"name": pinfo(mp[0])[0], "pos": pinfo(mp[0])[1], "pts": rnd(mp[1], 1)} if mp else None,
                        "bestGame": {"name": pinfo(bg[1])[0], "pos": pinfo(bg[1])[1], "pts": rnd(bg[0], 1),
                                     "week": bg[3], "opp": nick(bg[4]) if bg[4] else ""} if bg else None})
            scopes_out.append({"key": skey, "label": slabel, "career": career, "seasons": seasons_out,
                               "h2h": h2h, "records": rec, "topAll": top_all, "topByYear": top_year})
            leaderboard_by_scope[skey].append({
                "slug": slug[oid], "nick": nick(oid), "full": owners[oid]["full"] if SHOW_FULL_NAMES else "",
                "seasons": career["seasons"], "firstYear": career["firstYear"], "lastYear": career["lastYear"],
                "titles": career["titles"], "titleYears": career["titleYears"], "runnerUps": career["runnerUps"],
                "playoffApps": career["playoffApps"], "games": len(mine),
                "w": wlt_all[0], "l": wlt_all[1], "t": wlt_all[2],
                "pct": career["allPct"], "regPct": career["regPct"], "pf": career["pf"], "avgPf": career["avgPf"]})
        out = {"id": oid, "slug": slug[oid], "nick": nick(oid),
               "full": owners[oid]["full"] if SHOW_FULL_NAMES else "",
               "defaultScope": DEFAULT_RECORDS_SCOPE, "scopes": scopes_out,
               "excludesKickersDefenses": bool(bos.EXCLUDE_K_DST)}
        write_json(os.path.join(out_dir, "owner", f"{slug[oid]}.json"), out)
    # owners page: per scope, titles first, then win % inside that scope
    lb_scopes = []
    for skey, slabel, _ in scope_defs:
        lst = sorted(leaderboard_by_scope[skey], key=lambda e: (-e["titles"], -(e["pct"] or 0)))
        lb_scopes.append({"key": skey, "label": slabel, "owners": lst})
    write_json(os.path.join(out_dir, "owners.json"), {"defaultScope": DEFAULT_RECORDS_SCOPE, "scopes": lb_scopes})

    # ============ seasons ============
    latest_year = max(season_info) if season_info else None
    seasons_out = []
    last_week = 0
    for y in sorted(season_info, reverse=True):
        si = season_info[y]
        tss = [ts for ts, v in team_rows.items() if v["year"] == y]
        table = []
        for ts in tss:
            v, x = team_rows[ts], reg[ts]
            table.append({"slug": slug.get(v["owner"]), "nick": nick(v["owner"]), "team": v["name"],
                          "w": x[0], "l": x[1], "t": x[2], "pf": rnd(x[3], 1), "pa": rnd(x[4], 1),
                          "finish": v["finish"], "result": v["result"], "madePlayoffs": v["made"]})
        if table and all(t["finish"] for t in table):
            table.sort(key=lambda t: t["finish"])
            for i, t in enumerate(table):
                t["rank"] = t["finish"]
        else:
            table.sort(key=lambda t: (-(t["w"] + 0.5 * t["t"]), -t["pf"]))
            for i, t in enumerate(table):
                t["rank"] = i + 1
        champ = next((t for t in table if t["result"] == "Champion"), None)
        runner = next((t for t in table if t["result"] == "Runner-Up"), None)
        wk = max((g[1] for g in games if g[0] == y), default=0)
        if y == latest_year:
            last_week = wk
        seasons_out.append({**si, "standings": table, "champion": champ, "runnerUp": runner,
                            "weeksPlayed": wk, "inProgress": champ is None})
    write_json(os.path.join(out_dir, "seasons.json"), seasons_out)

    # ============ all-time totals (game scopes) ============
    alltime = {"scopes": []}
    for key, label, fn in scope_defs:
        lst = []
        for oid in order:
            rs = [r for r in by_owner[oid] if fn(r["t"])]
            if not rs:
                continue
            w, l, t = (sum(1 for r in rs if r["r"] == k) for k in "WLT")
            pf, pa = sum(r["pf"] for r in rs), sum(r["pa"] for r in rs)
            lst.append({"slug": slug[oid], "nick": nick(oid), "seasons": len({r["y"] for r in rs}),
                        "games": len(rs), "w": w, "l": l, "t": t, "pct": pct(w, l, t),
                        "pf": rnd(pf, 1), "pa": rnd(pa, 1), "avgPf": rnd(pf / len(rs), 2), "avgPa": rnd(pa / len(rs), 2)})
        alltime["scopes"].append({"key": key, "label": label, "minGames": SCOPE_MINS[key][0], "owners": lst})
    write_json(os.path.join(out_dir, "alltime.json"), alltime)

    # ============ league records ============
    def ranked(key, rows_, n=10, reverse=True):
        out = []
        for r in sorted(rows_, key=key, reverse=reverse)[:n]:
            out.append({"owner": nick(r["o"]), "slug": slug.get(r["o"]), "score": rnd(r["pf"]),
                        "oppScore": rnd(r["pa"]), "margin": rnd(r["pf"] - r["pa"]), "year": r["y"],
                        "week": r["w"], "type": r["t"], "opp": nick(r["p"]), "oppSlug": slug.get(r["p"])})
        return out

    game_type = {(r["y"], r["w"], r["ts"]): r["t"] for r in rows}

    year_final = {}
    for r in rows:
        year_final[r["y"]] = max(year_final.get(r["y"], 0), r["w"])

    def build_records(key, fn, cutoff=None):
        """Every record table for one game scope (fn decides which game types count)."""
        rec_rows = [r for r in rows if fn(r["t"]) and (cutoff is None or (r["y"], r["w"]) <= cutoff)]
        # a season counts once it is finished (and, for a snapshot, once its last week has been played)
        done = lambda v: v["result"] != "" and (cutoff is None or (v["year"], year_final.get(v["year"], 0)) <= cutoff)
        valid_all = [r for r in rec_rows if r["pf"] > 0 and r["pa"] > 0]
        wins_only = [r for r in valid_all if r["pf"] > r["pa"]]
        # team-season totals inside this scope
        ts_tot = defaultdict(lambda: [0, 0, 0, 0.0])
        for r in rec_rows:
            x = ts_tot[r["ts"]]
            x["WLT".index(r["r"])] += 1
            x[3] += r["pf"]
        season_pf = sorted(({"owner": nick(v["owner"]), "slug": slug.get(v["owner"]), "year": v["year"],
                             "team": v["name"], "pf": rnd(ts_tot[ts][3], 1), "w": ts_tot[ts][0], "l": ts_tot[ts][1], "t": ts_tot[ts][2]}
                            for ts, v in team_rows.items() if ts_tot[ts][3] > 0 and done(v)),
                           key=lambda d: -d["pf"])
        best_rec = sorted(({"owner": nick(v["owner"]), "slug": slug.get(v["owner"]), "year": v["year"],
                            "team": v["name"], "w": ts_tot[ts][0], "l": ts_tot[ts][1], "t": ts_tot[ts][2],
                            "pct": pct(*ts_tot[ts][:3]), "pf": rnd(ts_tot[ts][3], 1)}
                           for ts, v in team_rows.items() if sum(ts_tot[ts][:3]) >= SCOPE_MINS[key][1] and done(v)),
                          key=lambda d: (-d["pct"], -d["pf"]))
        # player stats: points a player scored while STARTED; needs a recorded game of this scope that week
        pgames = []
        tot_pts, tot_starts = defaultdict(float), Counter()
        tot_years, tot_owner = defaultdict(set), defaultdict(lambda: defaultdict(float))
        for y, w, tsid, pid, started, pts in roster:
            if not (started and pts is not None and tsid in teams):
                continue
            nm, pos = pinfo(pid)
            if bos.EXCLUDE_K_DST and pos in ("K", "DST"):
                continue
            gtp = game_type.get((y, w, tsid))
            if cutoff is not None and (y, w) > cutoff:
                continue
            if gtp is None or not fn(gtp):
                continue                      # no recorded game that week (eliminated team), or outside this scope
            pgames.append((pts, nm, pos, y, w, owner_of(tsid)))
            tot_pts[pid] += pts
            tot_starts[pid] += 1
            tot_years[pid].add(y)
            tot_owner[pid][owner_of(tsid)] += pts
        pgames.sort(key=lambda g: -g[0])

        def pline(pid):
            nm, pos = pinfo(pid)
            top_o = max(tot_owner[pid].items(), key=lambda kv: kv[1])[0]
            span = sorted(tot_years[pid])
            return {"name": nm, "pos": pos, "pts": rnd(tot_pts[pid], 1), "starts": tot_starts[pid],
                    "avg": rnd(tot_pts[pid] / tot_starts[pid], 1), "years": [span[0], span[-1]],
                    "owners": len(tot_owner[pid]), "topOwner": nick(top_o), "topSlug": slug.get(top_o)}

        return {
            "minStarts": SCOPE_MINS[key][2], "minSeasonGames": SCOPE_MINS[key][1],
            "playerTotals": [pline(pid) for pid, _ in sorted(tot_pts.items(), key=lambda kv: -kv[1])[:10]],
            # average points per start: only players with enough starts, so one big game can't top the list
            "playerAvg": [pline(pid) for pid in sorted((p for p in tot_pts if tot_starts[p] >= SCOPE_MINS[key][2]),
                                                       key=lambda p: -(tot_pts[p] / tot_starts[p]))[:10]],
            # same two lists without quarterbacks, who otherwise dominate both
            "playerTotalsNoQB": [pline(pid) for pid, _ in sorted(((p, v) for p, v in tot_pts.items() if pinfo(p)[1] != "QB"),
                                                                 key=lambda kv: -kv[1])[:10]],
            "playerAvgNoQB": [pline(pid) for pid in sorted((p for p in tot_pts if tot_starts[p] >= SCOPE_MINS[key][2] and pinfo(p)[1] != "QB"),
                                                           key=lambda p: -(tot_pts[p] / tot_starts[p]))[:10]],
            "playerStarts": [pline(pid) for pid in sorted(tot_starts, key=lambda p: (-tot_starts[p], -tot_pts[p]))[:10]],
            "highScores": ranked(lambda r: r["pf"], rec_rows),
            "lowScores": ranked(lambda r: r["pf"], valid_all, reverse=False),
            "blowouts": ranked(lambda r: r["pf"] - r["pa"], wins_only),
            "closest": ranked(lambda r: r["pf"] - r["pa"], wins_only, reverse=False),
            "seasonPoints": season_pf[:10],
            "seasonPointsLow": season_pf[::-1][:10],
            "bestRecords": best_rec[:10],
            "worstRecords": best_rec[::-1][:10],
            "playerGames": [{"name": nm, "pos": pos, "pts": rnd(pts, 1), "year": y, "week": w,
                             "owner": nick(o), "slug": slug.get(o)} for pts, nm, pos, y, w, o in pgames[:10]]}

    records = {
        "excludesKickersDefenses": bool(bos.EXCLUDE_K_DST),
        "defaultScope": DEFAULT_RECORDS_SCOPE,
        "scopes": [{"key": key, "label": label, **build_records(key, fn)} for key, label, fn in scope_defs]}
    write_json(os.path.join(out_dir, "records.json"), records)

    # ============ "records broken" box for the home page ============
    # Replays the record book (and the all-time tables) as of each of the last game weeks and reports what changed.
    weeks_played = sorted({(r["y"], r["w"]) for r in rows})
    rs_fn = next(fn for k, _, fn in scope_defs if k == "regpo")
    win = weeks_played[-ACTIVITY_WEEKS:]
    prior = ([(0, 0)] + weeks_played)[len(weeks_played) - len(win):len(weeks_played)]   # week before each window week

    REC_CATS = [("highScores", "Highest single-game score", lambda e: (e["slug"], e["year"], e["week"]), "game"),
                ("lowScores", "Lowest single-game score", lambda e: (e["slug"], e["year"], e["week"]), "game"),
                ("blowouts", "Biggest blowout", lambda e: (e["slug"], e["year"], e["week"]), "margin"),
                ("closest", "Closest finish", lambda e: (e["slug"], e["year"], e["week"]), "margin"),
                ("playerGames", "Biggest player game", lambda e: (e["name"], e["slug"], e["year"], e["week"]), "pgame"),
                ("seasonPoints", "Most points in a season", lambda e: (e["slug"], e["year"]), "season"),
                ("seasonPointsLow", "Fewest points in a season", lambda e: (e["slug"], e["year"]), "season"),
                ("bestRecords", "Best season record", lambda e: (e["slug"], e["year"]), "srec"),
                ("worstRecords", "Worst season record", lambda e: (e["slug"], e["year"]), "srec"),
                ("playerTotals", "Most career points by a player", lambda e: (e["name"], e["pos"], e["years"][0]), "ptot"),
                ("playerAvg", "Highest average points per start", lambda e: (e["name"], e["pos"], e["years"][0]), "pavg"),
                ("playerStarts", "Most career starts by a player", lambda e: (e["name"], e["pos"], e["years"][0]), "pst"),
                ("playerTotalsNoQB", "Most career points, no quarterbacks", lambda e: (e["name"], e["pos"], e["years"][0]), "ptot"),
                ("playerAvgNoQB", "Highest average per start, no quarterbacks", lambda e: (e["name"], e["pos"], e["years"][0]), "pavg")]

    def detail(e, kind):
        if kind == "game":
            return f"{e['score']:.2f} vs {e['opp']} ({e['oppScore']:.2f})"
        if kind == "margin":
            return f"+{e['margin']:.2f} over {e['opp']}"
        if kind == "pgame":
            return f"{e['pts']} pts for {e['owner']}"
        if kind == "season":
            return f"{e['pf']:,} pts, {e['team']}"
        if kind == "srec":
            return f"{e['w']}-{e['l']}" + (f"-{e['t']}" if e["t"] else "") + f", {e['team']}"
        if kind == "ptot":
            return f"{e['pts']:,} pts"
        if kind == "pavg":
            return f"{e['avg']} per start ({e['starts']} starts)"
        return f"{e['starts']} starts"

    def holder(e, kind):
        if kind == "pgame":
            return e["name"], None
        if kind in ("ptot", "pavg", "pst"):
            return e["name"], None
        return e["owner"], e["slug"]

    snap_cache = {}
    snap = lambda c: snap_cache.setdefault(c, build_records("regpo", rs_fn, cutoff=c))
    rec_events = []
    for wk, pw in zip(win, prior):
        before, after = snap(pw), snap(wk)
        for cat, title, keyf, kind in REC_CATS:
            b_rank = {keyf(e): i + 1 for i, e in enumerate(before[cat])}
            for i, e in enumerate(after[cat]):
                rank, prev = i + 1, b_rank.get(keyf(e))
                if prev is not None and prev <= rank:
                    continue
                who, sl = holder(e, kind)
                rec_events.append({"year": wk[0], "week": wk[1], "cat": title, "rank": rank, "prev": prev,
                                   "who": who, "slug": sl, "detail": detail(e, kind),
                                   "season": kind in ("season", "srec"),
                                   "career": kind in ("ptot", "pavg", "pst")})

    # all-time tables as of a week (same scope as the record book)
    AT = [("w", "Total wins", False, lambda x: x["w"]), ("l", "Total losses", False, lambda x: x["l"]),
          ("pct", "Overall win percentage", True, lambda x: x["pct"] or 0), ("pf", "Total points for", False, lambda x: x["pf"]),
          ("apf", "Points for per game", True, lambda x: x["pf"] / x["n"]), ("pa", "Total points against", False, lambda x: x["pa"]),
          ("apa", "Points against per game", True, lambda x: x["pa"] / x["n"])]

    def at_ranks(cutoff):
        stat = {}
        for oid in order:
            rs = [r for r in by_owner[oid] if rs_fn(r["t"]) and (r["y"], r["w"]) <= cutoff]
            if not rs:
                continue
            w, l, t = (sum(1 for r in rs if r["r"] == k) for k in "WLT")
            stat[oid] = {"w": w, "l": l, "t": t, "pct": pct(w, l, t), "n": len(rs),
                         "pf": sum(r["pf"] for r in rs), "pa": sum(r["pa"] for r in rs)}
        out = {}
        for tid, title, rate, val in AT:
            pool = {o: round(val(x), 4) for o, x in stat.items() if not rate or x["n"] >= SCOPE_MINS["regpo"][0]}
            out[tid] = {o: 1 + sum(1 for v2 in pool.values() if v2 > v) for o, v in pool.items()}
        return out, stat

    at_events = []
    cache_at = {}
    atr = lambda c: cache_at.setdefault(c, at_ranks(c))
    for wk, pw in zip(win, prior):
        (rb, _), (ra, st) = atr(pw), atr(wk)
        for tid, title, rate, val in AT:
            for oid, r in ra[tid].items():
                p = rb[tid].get(oid)
                if p is None or r >= p:
                    continue
                if rate and r != 1:
                    continue                  # per-game and percentage tables only matter when the lead changes
                x = st[oid]
                v = (f"{val(x) * 100:.1f}%" if tid == "pct" else f"{val(x):,.1f}" if tid in ("pf", "pa", "apf", "apa") else f"{val(x)}")
                at_events.append({"year": wk[0], "week": wk[1], "cat": title, "rank": r, "prev": p,
                                  "who": nick(oid), "slug": slug.get(oid), "value": v})

    activity = {"scope": "Regular season + playoffs", "through": {"year": weeks_played[-1][0], "week": weeks_played[-1][1]} if weeks_played else None,
                "weeks": [{"year": y, "week": w} for y, w in win],
                "records": sorted(rec_events, key=lambda e: (-e["year"], -e["week"], e["rank"])),
                "alltime": sorted(at_events, key=lambda e: (-e["year"], -e["week"], e["rank"]))}
    write_json(os.path.join(out_dir, "activity.json"), activity)

    # ============ rules, payouts (weekly + season end), penalties and money ============
    settings = load_settings(latest_year)
    write_json(os.path.join(out_dir, "rules.json"), load_rules(os.path.join(HERE, RULES_FILE)))
    start_year = settings["moneyStartYear"]
    slug_to_oid = {v: k for k, v in slug.items()}
    oid_of = lambda name: slug_to_oid.get(str(name or "").strip().lower())
    sp = settings["seasonPayouts"]

    # started players per (year, week): (points, name, position, owner)
    starters = defaultdict(list)
    season_pts = defaultdict(lambda: defaultdict(float))     # year -> pid -> points over every rostered week
    roster_at = {}                                           # (year, week, pid) -> owner
    for y, w, tsid, pid, started, pts in roster:
        if tsid not in teams or pts is None:
            continue
        season_pts[y][pid] += pts
        roster_at[(y, w, pid)] = owner_of(tsid)
        if started and game_type.get((y, w, tsid)) == "Regular":
            nm, pos = pinfo(pid)
            starters[(y, w)].append((pts, nm, pos or "", owner_of(tsid), pid))

    reg_rows = defaultdict(list)                             # (year, week) -> team-game rows (regular season)
    for r in rows:
        if r["t"] == "Regular" and (r["pf"] > 0 or r["pa"] > 0):
            reg_rows[(r["y"], r["w"])].append(r)

    def standings_after(y, k):
        st = defaultdict(lambda: [0, 0, 0, 0.0])
        for (yy, w), lst in reg_rows.items():
            if yy == y and w <= k:
                for r in lst:
                    st[r["o"]]["WLT".index(r["r"])] += 1
                    st[r["o"]][3] += r["pf"]
        rank = sorted(st, key=lambda o: (-((st[o][0] + 0.5 * st[o][2]) / max(1, sum(st[o][:3]))), -st[o][3], nick(o)))
        return st, rank

    wlt = lambda w_, l_, t_: f"{w_}-{l_}" + (f"-{t_}" if t_ else "")
    years = sorted((y for y in season_info if y >= start_year))
    payouts_out = {"seasons": []}
    money_seasons = []

    for y in sorted(years, reverse=True):
        tss = [ts for ts, v in team_rows.items() if v["year"] == y]
        n_teams = len(tss)
        done_weeks = sorted(w for (yy, w), lst in reg_rows.items() if yy == y and len(lst) >= n_teams)
        last_done = max(done_weeks, default=0)
        reg_end = max((w for (yy, w) in reg_rows if yy == y), default=0)
        reg_len = max(reg_end, 14)
        complete = any(team_rows[ts]["result"] == "Champion" for ts in tss)
        won = defaultdict(float)                              # owner -> payout money won
        manual = {(m["week"], m["kind"]): oid_of(m["winner"]) for m in settings["manualWinners"] if m["year"] == y}

        def W(o, txt):
            return {"o": o, "txt": txt}

        def high_score(w, **_):
            scored = [r for r in reg_rows[(y, w)] if r["pf"] > 0]
            top = max(r["pf"] for r in scored)
            tied = [r for r in scored if abs(r["pf"] - top) < 1e-9]
            if len(tied) > 1:                                  # tie-breaker: the team with the highest-scoring player
                best = {r["o"]: max([p[0] for p in starters[(y, w)] if p[3] == r["o"]] or [0]) for r in tied}
                m = max(best.values())
                tied = [r for r in tied if abs(best[r["o"]] - m) < 1e-9]
            return [W(r["o"], f"{r['pf']:.2f} points") for r in tied]

        def top_pos(w, pos=None, **_):
            lst = [p for p in starters[(y, w)] if (pos is None and p[2] not in ("QB", "")) or p[2] == pos]
            if not lst:
                return []
            top = max(p[0] for p in lst)
            return [W(p[3], f"{p[1]} ({p[2]}) {rnd(p[0], 1)} pts") for p in lst if abs(p[0] - top) < 1e-9]

        def beats_high_loser(w, **_):
            losers = [r for r in reg_rows[(y, w)] if r["r"] == "L"]
            if not losers:
                return []
            top = max(r["pf"] for r in losers)
            return [W(r["p"], f"beat {nick(r['o'])}, who lost with {r['pf']:.2f}") for r in losers if abs(r["pf"] - top) < 1e-9]

        def blowout(w, **_):
            wins = [r for r in reg_rows[(y, w)] if r["r"] == "W"]
            top = max(r["pf"] - r["pa"] for r in wins)
            return [W(r["o"], f"+{r['pf'] - r['pa']:.2f} over {nick(r['p'])}") for r in wins if abs(r["pf"] - r["pa"] - top) < 1e-9]

        def first_after(w, **_):
            st, rank = standings_after(y, w)
            o = rank[0]
            return [W(o, f"{wlt(st[o][0], st[o][1], st[o][2])}, {st[o][3]:,.1f} PF")]

        def beats_best_loser(w, **_):
            st, rank = standings_after(y, w)
            losers = {r["o"]: r for r in reg_rows[(y, w)] if r["r"] == "L"}
            best = next((o for o in rank if o in losers), None)
            if not best:
                return []
            r = losers[best]
            return [W(r["p"], f"beat {nick(best)}, the best-ranked loser (#{rank.index(best) + 1})")]

        def season_top_non_qb(w, **_):
            best, bp = None, 0
            for pid, pts in season_pts[y].items():
                if pinfo(pid)[1] in ("QB", ""):
                    continue
                if pts > bp and (y, w, pid) in roster_at:
                    best, bp = pid, pts
            if best is None:
                return []
            nm, pos = pinfo(best)
            return [W(roster_at[(y, w, best)], f"{nm} ({pos}) {rnd(bp, 1)} pts on the season")]

        def manual_pick(w, kind=None, **_):
            o = manual.get((w, kind))
            return [W(o, "entered in league_settings.json")] if o else []

        KINDS = {"highScore": high_score, "topPos": top_pos, "beatsHighLoser": beats_high_loser, "blowout": blowout,
                 "firstAfter": first_after, "bestRankLoser": beats_best_loser, "seasonTopNonQB": season_top_non_qb,
                 "manual": manual_pick}

        schedule = []
        for pc in settings["weeklyPayouts"]:
            item = {"id": pc["id"], "label": pc["label"], "week": pc.get("week"), "rule": pc.get("rule", ""),
                    "amount": pc["amount"], "winners": [], "status": "pending", "note": ""}
            if pc["kind"] == "undefeated":
                # last team to lose its undefeated record; tied teams keep playing until one loses and another does not
                U, tied = set(team_rows[ts]["owner"] for ts in tss), None
                res = None
                for w in done_weeks:
                    losers = {r["o"] for r in reg_rows[(y, w)] if r["r"] == "L"}
                    if tied is None:
                        new = U - losers
                        if not new:
                            if len(U) == 1:
                                res = set(U); break
                            tied = set(U)
                        else:
                            U = new
                    else:
                        lost = tied & losers
                        if 0 < len(lost) < len(tied):
                            tied = tied - lost
                            if len(tied) == 1:
                                res = set(tied); break
                if res is None and tied is None and last_done >= reg_len and U:
                    res = set(U)                                 # never lost all season
                if res:
                    item["winners"] = [W(o, "last team to lose its undefeated record") for o in sorted(res, key=nick)]
                    item["status"] = "final"
                else:
                    cur = tied if tied is not None else U
                    item["status"] = "open"
                    item["note"] = f"still undefeated after week {last_done}: " + ", ".join(sorted(nick(o) for o in cur)) if tied is None else \
                        "tied teams: " + ", ".join(sorted(nick(o) for o in tied))
            else:
                w = pc["week"]
                if pc["kind"] == "manual" or w in done_weeks:
                    fn = KINDS[pc["kind"]]
                    item["winners"] = [x for x in (fn(w, pos=pc.get("pos"), kind=pc.get("id")) or [])]
                    item["status"] = "final" if item["winners"] else "pending"
                    if pc["kind"] == "manual" and not item["winners"]:
                        item["note"] = "winner not entered yet" if w in done_weeks else ""
            share = item["amount"] / len(item["winners"]) if item["winners"] else 0
            for x in item["winners"]:
                won[x["o"]] += share
            item["winners"] = [{"slug": slug.get(x["o"]), "nick": nick(x["o"]), "txt": x["txt"]} for x in item["winners"]]
            schedule.append(item)

        # ---- weekly penalty: the lowest scorers of the week pay ----
        pen_amt, pen_n = settings["lowestPenalty"], settings["lowestCount"]
        lowest, owes_pen = [], defaultdict(float)
        for w in done_weeks:
            sc = sorted((r for r in reg_rows[(y, w)] if r["pf"] > 0), key=lambda r: r["pf"])
            if len(sc) < pen_n:
                continue
            cut = sc[pen_n - 1]["pf"]
            lows = [r for r in sc if r["pf"] <= cut + 1e-9]
            ov = next((x for x in settings["lowestOverrides"] if x["year"] == y and x["week"] == w), None)
            if ov:                                           # the commissioner's own record wins over the computed list
                names = {oid_of(n) for n in ov["owners"]}
                lows = [r for r in sc if r["o"] in names]
            for r in lows:
                owes_pen[r["o"]] += pen_amt
            lowest.append({"week": w, "amount": pen_amt, "owners": [{"slug": slug.get(r["o"]), "nick": nick(r["o"]), "pf": rnd(r["pf"])} for r in lows]})
        manual_pens = []
        for pn in settings["penalties"]:
            o = oid_of(pn["owner"])
            if pn["year"] == y and o:
                owes_pen[o] += pn["amount"]
                manual_pens.append({"week": pn.get("week"), "slug": slug.get(o), "nick": nick(o), "amount": pn["amount"], "note": pn.get("note", "")})

        # ---- season-end payouts ----
        champ = next((team_rows[ts]["owner"] for ts in tss if team_rows[ts]["result"] == "Champion"), None)
        runner = next((team_rows[ts]["owner"] for ts in tss if team_rows[ts]["result"] == "Runner-Up"), None)
        third = next((team_rows[ts]["owner"] for ts in tss if team_rows[ts]["finish"] == 3), None)
        st_end, rank_end = standings_after(y, reg_len)
        reg_complete = last_done >= reg_len
        top_pf_o = max(st_end, key=lambda o: st_end[o][3]) if st_end else None
        missed = [team_rows[ts]["owner"] for ts in tss if not team_rows[ts]["made"]] if complete else []
        tb_pot = settings["missPlayoffsPenalty"] * len(missed)
        tb_winner = oid_of((settings.get("toiletBowlWinner") or {}).get(str(y)))
        season_end = []

        def end_item(i, label, amount, who, txt, ready, rule=""):
            ws_ = [who] if (ready and who) else []
            for o in ws_:
                won[o] += amount
            season_end.append({"id": i, "label": label, "amount": amount, "rule": rule, "status": "final" if ws_ else "pending",
                               "winners": [{"slug": slug.get(o), "nick": nick(o), "txt": txt} for o in ws_]})

        end_item("first", "1st place", sp["first"], champ, "champion", complete)
        end_item("second", "2nd place", sp["second"], runner, "runner-up", complete)
        end_item("third", "3rd place", sp["third"], third, "3rd place game", complete)
        end_item("regFirst", "1st in regular-season standings", sp["regSeasonFirst"], rank_end[0] if rank_end else None,
                 f"{wlt(*st_end[rank_end[0]][:3])}" if rank_end else "", reg_complete)
        end_item("mostPoints", "Most points in the regular season", sp["mostPoints"], top_pf_o,
                 f"{st_end[top_pf_o][3]:,.1f} PF" if top_pf_o else "", reg_complete)
        end_item("toilet", f"Toilet bowl winner (pot of ${tb_pot})" if complete else "Toilet bowl winner",
                 tb_pot if complete else settings["missPlayoffsPenalty"] * 6, tb_winner, "toilet bowl", complete and tb_winner is not None,
                 rule="Funded by the $%d penalty from each team that misses the playoffs." % settings["missPlayoffsPenalty"])
        for o in missed:
            owes_pen[o] += settings["missPlayoffsPenalty"]

        weekly_total = sum(x["amount"] for x in schedule)
        board = sorted(({"slug": slug.get(o), "nick": nick(o), "won": rnd(v)} for o, v in won.items() if v),
                       key=lambda d: (-d["won"], d["nick"]))
        payouts_out["seasons"].append({"year": y, "weeksPlayed": last_done, "complete": complete, "schedule": schedule,
                                       "seasonEnd": season_end, "lowest": sorted(lowest, key=lambda d: -d["week"]),
                                       "penalties": manual_pens, "board": board, "weeklyTotal": weekly_total,
                                       "feeNote": settings.get("feeNote", "")})

        # ---- money per owner ----
        fee_of = {oid_of(f["owner"]): f for f in settings["fees"] if oid_of(f["owner"])}
        paid_to = defaultdict(float)
        pay_rows = []
        for p in settings["payments"]:
            o = oid_of(p.get("owner"))
            if p.get("year") == y and o:
                paid_to[o] += p["amount"]
                pay_rows.append({"slug": slug.get(o), "nick": nick(o), "amount": p["amount"], "note": p.get("note", "")})
        owners_out = []
        for ts in tss:
            o = team_rows[ts]["owner"]
            f = fee_of.get(o)
            pen = owes_pen.get(o, 0.0)
            row = {"slug": slug.get(o), "nick": nick(o), "won": rnd(won.get(o, 0.0)), "penalties": rnd(pen),
                   "feeBase": None, "feeTotal": None, "paidIn": None, "owe": None, "paidOut": rnd(paid_to.get(o, 0.0))}
            if f:
                total = f["base"] + pen
                row.update({"feeBase": rnd(f["base"]), "feeTotal": rnd(total), "paidIn": rnd(f.get("paid", 0)),
                            "owe": rnd(total - f.get("paid", 0))})
                row["net"] = rnd(row["won"] - row["owe"] - row["paidOut"])
            else:
                row["net"] = rnd(row["won"] - pen - row["paidOut"])
            owners_out.append(row)
        owners_out.sort(key=lambda d: (-d["net"], d["nick"]))
        money_seasons.append({"year": y, "complete": complete, "owners": owners_out, "payments": pay_rows})

    write_json(os.path.join(out_dir, "payouts.json"), payouts_out)
    write_json(os.path.join(out_dir, "money.json"), {"confirmed": settings["confirmed"], "feeNote": settings.get("feeNote", ""), "seasons": money_seasons})

    # ============ head-to-head matrix (one per game scope) ============
    h2h_scopes = []
    for skey, slabel, sfn in scope_defs:
        m = {slug[o]: {slug[p]: [0, 0, 0] for p in order if p != o} for o in order}
        seq = defaultdict(list)                  # (owner, opponent) -> results in date order
        for r in sorted(rows, key=lambda r: (r["y"], r["w"])):
            if sfn(r["t"]) and r["o"] in slug and r["p"] in slug and r["o"] != r["p"] and r["o"] in order and r["p"] in order:
                m[slug[r["o"]]][slug[r["p"]]]["WLT".index(r["r"])] += 1
                seq[(slug[r["o"]], slug[r["p"]])].append((r["r"], r["y"], r["w"]))
        # streaks: signed current run (+n = owner has won the last n meetings, -n = lost the last n; a tie ends a run),
        # and the longest winning run the owner has had against that opponent
        cur = {a: {} for a in m}
        best = {a: {} for a in m}
        for (a, b), res in seq.items():
            n, last = 0, None
            for x in reversed(res):
                if last is None:
                    if x[0] == "T":
                        break
                    last = x[0]
                if x[0] == last:
                    n += 1
                else:
                    break
            cur[a][b] = n if last == "W" else -n if last == "L" else 0
            run = top = 0
            for x in res:
                run = run + 1 if x[0] == "W" else 0
                top = max(top, run)
            best[a][b] = top
        h2h_scopes.append({"key": skey, "label": slabel, "matrix": m, "streak": cur, "best": best})
    write_json(os.path.join(out_dir, "h2h.json"),
               {"order": [slug[o] for o in order], "nicks": {slug[o]: nick(o) for o in order},
                "defaultScope": DEFAULT_RECORDS_SCOPE, "scopes": h2h_scopes})

    # ============ draft ============
    years_out = []
    for y in sorted(draft_by_year, reverse=True):
        picks = sorted(draft_by_year[y])
        # Yahoo numbers 2016's picks on a 14-slot board (3 placeholder "Bye Week" slots), which leaves gaps;
        # show picks as 1..n inside each round
        by_rd = defaultdict(list)
        for rd_, pk_, ts_, pl_ in picks:
            by_rd[rd_].append((rd_, pk_, ts_, pl_))
        picks = [(rd_, i + 1, ts_, pl_) for rd_ in sorted(by_rd) for i, (_, _, ts_, pl_) in enumerate(by_rd[rd_])]
        n_teams = (season_info.get(y) or {}).get("numTeams") or len({p[2] for p in picks})
        rounds = max(p[0] for p in picks)
        expected = n_teams * rounds
        years_out.append({
            "year": y, "teams": n_teams, "rounds": rounds, "count": len(picks), "expected": expected,
            "complete": len(picks) >= expected,
            "picks": [{"round": rd, "pick": pk, "owner": nick(owner_of(ts)), "slug": slug.get(owner_of(ts)),
                       "team": (team_rows.get(ts) or {}).get("name", ""), "player": pl or ""}
                      for rd, pk, ts, pl in picks]})
    write_json(os.path.join(out_dir, "draft.json"), years_out)

    # ============ power rankings (projected best lineup, weeks still to play) ============
    # fetch_projections.py saves Sleeper's projected points for every remaining week plus each team's
    # current roster. Each remaining week every team's roster is fed into the starting slots (QB, 2 RB,
    # 2 WR, TE, one RB/WR/TE flex, K, DST) to build its best projected lineup; teams are ranked on it
    # (1st earns N points, last earns 1) and the points add up. Weeks already played never count.
    PR_SLOTS = [("QB", ("QB",), 1), ("RB", ("RB",), 2), ("WR", ("WR",), 2), ("TE", ("TE",), 1),
                ("K", ("K",), 1), ("DST", ("DST", "DEF"), 1)]
    PR_FLEX = ("RB", "WR", "TE")
    PR_BLOCKS = [("early", "Weeks 1-14", 1, 14), ("late", "Playoffs", 15, 17)]

    # Magnificos: the starting slots come from the league's own Sleeper roster settings (saved by
    # fetch_projections.py), e.g. QB, 2 RB, 2 WR, TE, 3 FLEX, K, DEF. Fixed slots are filled first,
    # then each flex slot takes the best leftover eligible player.
    FLEX_ELIG = {"FLEX": ("RB", "WR", "TE"), "SUPER_FLEX": ("QB", "RB", "WR", "TE"),
                 "REC_FLEX": ("WR", "TE"), "WRRB_FLEX": ("RB", "WR")}
    PR_FLEX_SLOTS = [("FLEX", PR_FLEX)]

    def set_slots(roster_positions):
        nonlocal PR_SLOTS, PR_FLEX_SLOTS
        rp = [x for x in (roster_positions or []) if x not in ("BN", "IR", "TAXI")]
        if not rp:
            return
        fixed, flex = [], []
        for pos_ in ("QB", "RB", "WR", "TE", "K", "DEF"):
            n_ = rp.count(pos_)
            if n_:
                fixed.append(("DST" if pos_ == "DEF" else pos_, ("DST", "DEF") if pos_ == "DEF" else (pos_,), n_))
        for x in rp:
            if x in FLEX_ELIG:
                flex.append((x, FLEX_ELIG[x]))
        PR_SLOTS, PR_FLEX_SLOTS = fixed, flex

    def best_lineup(entries):                        # entries: (pts, pid, pos)
        pool = sorted(entries, key=lambda e: -e[0])
        chosen, used = [], set()
        for slot, positions, n in PR_SLOTS:
            k = 0
            for e in pool:
                if k < n and e[2] in positions and e[1] not in used:
                    chosen.append((slot, e)); used.add(e[1]); k += 1
        for fslot, elig in PR_FLEX_SLOTS:                # each flex: best leftover eligible player
            for e in pool:
                if e[2] in elig and e[1] not in used:
                    chosen.append((fslot, e)); used.add(e[1])
                    break
        return chosen

    power = {"year": latest_year, "asOf": None, "firstWeek": None, "blocks": [], "weeks": {}}
    # weights per table: "early" = weeks 1-14, "late" = playoffs (weeks 15-17). A flat powerWeights dict
    # (no early/late keys) is used for the early table, and the playoffs table stays 100% projected lineup.
    pw_cfg = settings.get("powerWeights") or {}
    PW_DEFAULT = {"early": {"lineup": 75, "bench": 5, "record": 15, "points": 5},
                  "late": {"lineup": 100, "bench": 0, "record": 0, "points": 0}}

    def block_weights(bkey):
        src = pw_cfg.get(bkey) if ("early" in pw_cfg or "late" in pw_cfg) else (pw_cfg if bkey == "early" else None)
        src = src or PW_DEFAULT[bkey]
        w_ = {k: float(src.get(k, 0)) for k in ("lineup", "bench", "record", "points")}
        tot_ = sum(w_.values()) or 1.0
        return w_, tot_, {k: round(v / tot_ * 100, 1) for k, v in w_.items()}
    BENCH_N = 4                                        # best leftover players that make up "bench depth"
    proj_path = os.path.join(HERE, PROJECTIONS_FILE)
    proj = None
    if os.path.exists(proj_path):
        with open(proj_path, encoding="utf-8") as f:
            proj = json.load(f)
        if proj.get("season") != latest_year:
            print(f"   {PROJECTIONS_FILE} is for {proj.get('season')}, not {latest_year} -- ignored")
            proj = None
    else:
        print(f"   No {PROJECTIONS_FILE} yet -- run fetch_projections.py to fill in the Power Rankings page.")
    # ============ this week's matchups (home page) ============
    tw = (proj or {}).get("matchups") if proj else None
    tw_path = os.path.join(out_dir, "thisweek.json")
    if tw and tw.get("pairs"):
        wk_now, yr_now = int(tw["week"]), latest_year
        def _own(oid_):
            t_ = team_rows.get(f"{yr_now}-{oid_}")
            return t_["owner"] if t_ and t_["owner"] in slug else None
        pairs_ = [(_own(a_), _own(b_)) for a_, b_ in tw["pairs"]]
        pairs_ = [(a_, b_) for a_, b_ in pairs_ if a_ and b_]
        tw_scopes = []
        for skey, slabel, sfn in scope_defs:
            ms = []
            for a_, b_ in pairs_:
                res = sorted(((r["y"], r["w"], r) for r in rows
                              if r["o"] == a_ and r["p"] == b_ and sfn(r["t"]) and (r["y"], r["w"]) < (yr_now, wk_now)),
                             key=lambda x: (x[0], x[1]))
                w_, l_, t_ = (sum(1 for x in res if x[2]["r"] == c) for c in "WLT")
                n, last_res = 0, None
                for x in reversed(res):
                    if last_res is None:
                        if x[2]["r"] == "T":
                            break
                        last_res = x[2]["r"]
                    if x[2]["r"] == last_res:
                        n += 1
                    else:
                        break
                streak = n if last_res == "W" else -n if last_res == "L" else 0
                def _longest(who):
                    run = top = 0
                    for x in res:
                        mine = x[2]["r"] == "W" if who == "a" else x[2]["r"] == "L"
                        run = run + 1 if mine else 0
                        top = max(top, run)
                    return top
                lastg = None
                if res:
                    y_, wk_, r_ = res[-1]
                    lastg = {"year": y_, "week": wk_, "a": rnd(r_["pf"]), "b": rnd(r_["pa"]), "type": r_["t"]}
                ms.append({"a": {"slug": slug[a_], "nick": nick(a_)}, "b": {"slug": slug[b_], "nick": nick(b_)},
                           "rec": [w_, l_, t_], "streak": streak, "bestA": _longest("a"), "bestB": _longest("b"), "last": lastg})
            tw_scopes.append({"key": skey, "label": slabel, "matchups": ms})
        write_json(tw_path, {"season": yr_now, "week": wk_now, "asOf": proj.get("fetchedAt"),
                             "defaultScope": DEFAULT_RECORDS_SCOPE, "scopes": tw_scopes})
    elif os.path.exists(tw_path):
        os.remove(tw_path)                              # stale file from an earlier week: the home page hides the section
    pr_lineup, pr_bench = {}, {}                       # week -> {tsid: projected points}
    ts_of = {}
    if proj:
        power["asOf"], power["firstWeek"] = proj.get("fetchedAt"), proj.get("firstWeek")
        set_slots(proj.get("rosterPositions"))
        pinfo_p = proj.get("players", {})
        for oid_, plist in proj.get("teams", {}).items():
            tsid = f"{latest_year}-{oid_}"
            if tsid in team_rows and team_rows[tsid]["owner"] in slug:
                ts_of[tsid] = plist
        for w_s, pts_map in proj.get("weeks", {}).items():
            w = int(w_s)
            for tsid, plist in ts_of.items():
                ents = [(float(pts_map[p]), p, (pinfo_p.get(p) or {}).get("p", "")) for p in plist if p in pts_map]
                lu = best_lineup(ents)
                if not lu:
                    continue
                used_ids = {e[1] for _, e in lu}
                bench = sorted((e[0] for e in ents if e[1] not in used_ids and e[2] in ("QB", "RB", "WR", "TE")), reverse=True)
                pr_lineup.setdefault(w, {})[tsid] = round(sum(e[0] for _, e in lu), 2)
                pr_bench.setdefault(w, {})[tsid] = round(sum(bench[:BENCH_N]), 2)

    def rank_scores(vals):                              # {key: value} -> {key: (rank, 0-100 score)}; 1st = 100, last = 0
        order_ = sorted(vals.items(), key=lambda kv: -kv[1])
        n_ = len(order_)
        return {k: (i + 1, (n_ - 1 - i) / (n_ - 1) * 100 if n_ > 1 else 50.0) for i, (k, _) in enumerate(order_)}

    # season-to-date record, standing and points scored (regular season, played games)
    rec_score = {}
    if ts_of:
        n_t = len(ts_of)
        recs = {ts: reg[ts] for ts in ts_of}
        stand = sorted(recs, key=lambda t: (-(recs[t][0] + 0.5 * recs[t][2]), -recs[t][3]))
        pf_rank = rank_scores({t: recs[t][3] for t in recs})
        for i, ts in enumerate(stand):
            w_, l_, t_ = recs[ts][0], recs[ts][1], recs[ts][2]
            g_ = w_ + l_ + t_
            winpct = (w_ + 0.5 * t_) / g_ * 100 if g_ else 50.0
            standing = (n_t - 1 - i) / (n_t - 1) * 100 if n_t > 1 and g_ else 50.0
            rec_score[ts] = {"record": round((winpct + standing) / 2, 1), "wins": w_, "losses": l_, "ties": t_,
                             "standing": i + 1 if g_ else None,
                             "pf": rnd(recs[ts][3], 1), "pfScore": round(pf_rank[ts][1] if g_ else 50.0, 1)}

    if pr_lineup:
        for w in sorted(pr_lineup):
            ls, bs = rank_scores(pr_lineup[w]), rank_scores(pr_bench[w])
            n = len(ls)
            power["weeks"][str(w)] = [{
                "slug": slug[team_rows[tsid]["owner"]], "nick": nick(team_rows[tsid]["owner"]),
                "team": team_rows[tsid]["name"], "rank": ls[tsid][0], "points": n - ls[tsid][0] + 1,
                "proj": pr_lineup[w][tsid], "bench": pr_bench[w][tsid], "benchRank": bs[tsid][0],
                "lineupScore": ls[tsid][1], "benchScore": bs[tsid][1]}
                for tsid in sorted(ls, key=lambda t: ls[t][0])]
    for bkey, blabel, w0, w1 in PR_BLOCKS:
        wks = [w for w in sorted(pr_lineup) if w0 <= w <= w1]
        tab = {}
        for w in wks:
            for e in power["weeks"][str(w)]:
                t = tab.setdefault(e["slug"], {"slug": e["slug"], "nick": e["nick"], "team": e["team"], "points": 0,
                                               "proj": 0.0, "bench": 0.0, "ls": 0.0, "bs": 0.0, "ranks": {}})
                t["points"] += e["points"]; t["proj"] += e["proj"]; t["bench"] += e["bench"]
                t["ls"] += e["lineupScore"]; t["bs"] += e["benchScore"]
                t["ranks"][str(w)] = e["rank"]
        ts_by_slug = {slug[team_rows[ts]["owner"]]: ts for ts in ts_of}
        pw, pw_sum, pw_pct = block_weights(bkey)
        for sl, t in tab.items():
            k = len(t["ranks"])
            rs = rec_score.get(ts_by_slug.get(sl), {})
            t["lineupScore"] = round(t["ls"] / k, 1)
            t["benchScore"] = round(t["bs"] / k, 1)
            t["recordScore"] = rs.get("record", 50.0)
            t["pointsScore"] = rs.get("pfScore", 50.0)
            t["wins"], t["losses"], t["ties"] = rs.get("wins", 0), rs.get("losses", 0), rs.get("ties", 0)
            t["standing"], t["pf"] = rs.get("standing"), rs.get("pf", 0)
            t["score"] = round((pw["lineup"] * t["lineupScore"] + pw["bench"] * t["benchScore"] +
                                pw["record"] * t["recordScore"] + pw["points"] * t["pointsScore"]) / pw_sum, 1)
            t["avgRank"] = rnd(sum(t["ranks"].values()) / k, 2)
            t["proj"], t["bench"] = rnd(t["proj"], 1), rnd(t["bench"], 1)
            for d_ in ("ls", "bs"):
                del t[d_]
        lst = sorted(tab.values(), key=lambda t: (-t["score"], -t["proj"]))
        for i, t in enumerate(lst):
            t["rank"] = i + 1
        power["blocks"].append({"key": bkey, "label": blabel, "firstWeek": w0, "lastWeek": w1,
                                "weeksPlayed": wks, "weights": pw_pct, "teams": lst})
    write_json(os.path.join(out_dir, "power.json"), power)

    # ============ meta ============
    meta = {"leagueName": LEAGUE_NAME,
            "updated": datetime.now(timezone.utc).strftime("%Y-%m-%d"),
            "currentSeason": latest_year, "currentWeek": last_week,
            "firstYear": min(season_info) if season_info else None,
            "owners": [{"slug": slug[o], "nick": nick(o), "titles": titles[o]} for o in order],
            "champions": [{"year": s["year"], "slug": s["champion"]["slug"], "nick": s["champion"]["nick"],
                           "team": s["champion"]["team"]} for s in sorted(seasons_out, key=lambda s: s["year"])
                          if s["champion"]]}
    write_json(os.path.join(out_dir, "meta.json"), meta)
    return meta, len(order)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook", default=None)
    ap.add_argument("--out", default=os.path.join(HERE, "site", "data"))
    a = ap.parse_args()
    from ffl_extract import WORKBOOK_PATH
    wbp = a.workbook or WORKBOOK_PATH
    assert_isolated(wbp, a.out)        # never read/write Camaradas files
    meta, n = build(wbp, a.out)
    print(f"Exported {n} owners, {len(meta['champions'])} champions; season {meta['currentSeason']} "
          f"through week {meta['currentWeek']}.")
    print(f"JSON files written to {a.out}")


if __name__ == "__main__":
    main()

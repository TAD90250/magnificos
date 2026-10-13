#!/usr/bin/env python3
"""
Builds one tab per owner plus three shared sheets inside FFL_Magnificos_Database.xlsx:

  Leaderboard     all-time table of every owner (click a name to jump to their tab)
  H2H_Matrix      every owner vs every other owner (W-L)
  League_Records  top-10 highest single-game scores and biggest margins of victory
  <one tab per owner>
      career summary, season-by-season (with the round-1 draft pick), head-to-head
      vs everyone, personal records (highest / lowest score, biggest win / loss),
      and top players -- all-time and for each year.

Everything except "top players" is a live formula, so the tabs update themselves when
the workbook is recalculated. The top-players blocks are computed here (they cover
~30,000 lineup rows), so re-run this script after new weeks are loaded:

    python build_owner_sheets.py

Safe to re-run: it deletes and rebuilds only the tabs it created. Backup made first.
Definitions used for the player stats:
  started        = in a starting lineup slot (not bench / IR)
  points         = fantasy points scored while STARTED (bench points don't count for the team)
  games on roster = every week the player was on the team, started or not
"""
import os
import re
import shutil
from collections import Counter, defaultdict

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

import schema_io
from ffl_extract import WORKBOOK_PATH

EXCLUDE_K_DST = False        # True = ignore kickers and team defenses in the player stats

from league_config import OWNER_NAMES, LEAGUE_NAME   # Magnificos names live in league_config.py

FONT = schema_io.FONT_NAME
B = 10000                                        # Team_Weeks formula bound
F_ = "Team_Weeks!$%s$2:$%s$" + str(B)


def tw(col):
    return F_ % (col, col)


TW_YEAR, TW_WEEK, TW_TYPE, TW_PF, TW_PA, TW_RES, TW_MARGIN, TW_OWNER, TW_OPP = (
    tw("B"), tw("C"), tw("D"), tw("G"), tw("H"), tw("I"), tw("J"), tw("K"), tw("L"))
TEAMS = lambda c: f"Teams!${c}$2:${c}$1000"
DRAFT = lambda c: f"Draft!${c}$2:${c}$3000"

NAVY = "1F4E78"
FILL_TITLE = PatternFill("solid", fgColor=NAVY)
FILL_SECTION = PatternFill("solid", fgColor="2E75B6")
FILL_HEAD = PatternFill("solid", fgColor="D9E1F2")
FILL_ALT = PatternFill("solid", fgColor="F2F6FB")
FILL_GOLD = PatternFill("solid", fgColor="FFE699")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(bottom=THIN)
GENERATED_SHARED = ["Leaderboard", "H2H_Matrix", "League_Records"]


# ---------------------------------------------------------------- data loading
def load_tables(wb):
    owners = {}
    for r in wb["Owners"].iter_rows(min_row=2, values_only=True):
        if r[0]:
            nick = re.split(r"[,/]", str(r[2] or r[1] or r[0]))[0].strip() or r[0]
            owners[r[0]] = {"full": r[1] or r[0], "nick": nick}
    teams = {}                                     # tsid -> (year, owner, name)
    for r in wb["Teams"].iter_rows(min_row=2, values_only=True):
        if r[1] and r[2]:
            teams[f"{r[1]}-{r[2]}"] = (r[1], r[2], r[3])
    games = []                                     # year, week, type, tsA, scA, tsB, scB
    for r in wb["Matchups"].iter_rows(min_row=2, values_only=True):
        if r[1] and r[4] and r[6]:
            games.append((r[1], r[2], r[3], r[4], r[5], r[6], r[7]))
    players = {}
    for r in wb["Players"].iter_rows(min_row=2, values_only=True):
        if r[0]:
            players[r[0]] = (r[1], r[2])
    roster = []                                    # year, week, tsid, pid, starter, pts
    for r in wb["Weekly_Rosters"].iter_rows(min_row=2, values_only=True):
        if r[0]:
            roster.append((r[1], r[2], r[3], r[4], r[6] == "Yes", r[7]))
    return owners, teams, games, players, roster


def owner_of(tsid):
    return tsid.split("-", 1)[1]


def safe_title(name, used):
    t = re.sub(r"[\[\]:*?/\\]", "", name)[:31] or "Owner"
    base, n = t, 2
    while t.lower() in used:
        t = f"{base[:28]}_{n}"
        n += 1
    used.add(t.lower())
    return t


# ---------------------------------------------------------------- python-side stats
def player_stats(owners, teams, games, players, roster):
    opp = {}
    for y, w, gt, a, sa, b, sb in games:
        opp[(y, w, a)] = b
        opp[(y, w, b)] = a
    per = defaultdict(lambda: {
        "starts": Counter(), "roster": Counter(), "pts": defaultdict(float), "best": [],
        "y_starts": defaultdict(Counter), "y_pts": defaultdict(lambda: defaultdict(float)),
        "y_best": defaultdict(list), "years": defaultdict(set)})
    for y, w, tsid, pid, started, pts in roster:
        if tsid not in teams:
            continue
        pos = players.get(pid, (None, None))[1]
        if EXCLUDE_K_DST and pos in ("K", "DST"):
            continue
        o = owner_of(tsid)
        d = per[o]
        d["roster"][pid] += 1
        d["years"][pid].add(y)
        if started:
            d["starts"][pid] += 1
            d["y_starts"][y][pid] += 1
            if pts is not None:
                d["pts"][pid] += pts
                d["y_pts"][y][pid] += pts
                game = (pts, pid, y, w, owner_of(opp[(y, w, tsid)]) if (y, w, tsid) in opp else "")
                d["best"].append(game)
                d["y_best"][y].append(game)
    return per


def pname(players, pid):
    return players.get(pid, (pid, None))[0] or pid


def yrs(span):
    span = sorted(span)
    return f"{span[0]}" if span[0] == span[-1] else f"{span[0]}-{span[-1]}"


# ---------------------------------------------------------------- styling helpers
def put(ws, ref, value, bold=False, color=None, fill=None, size=10, fmt=None, align=None, italic=False, wrap=False):
    c = ws[ref]
    c.value = value
    c.font = Font(name=FONT, size=size, bold=bold, italic=italic, color=color)
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if align or wrap:
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    return c


def section(ws, row, text, last_col=11):
    for c in range(1, last_col + 1):
        ws.cell(row=row, column=c).fill = FILL_SECTION
    put(ws, f"A{row}", text, bold=True, color="FFFFFF", fill=FILL_SECTION, size=11)
    ws.row_dimensions[row].height = 20


def head_row(ws, row, labels, start_col=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i)
        c.value = lab
        c.font = Font(name=FONT, size=10, bold=True)
        c.fill = FILL_HEAD
        c.alignment = Alignment(horizontal="center" if i else "left", vertical="center", wrap_text=True)
        c.border = BOX


def style_body(ws, row, first_col, last_col, alt=False, fmt=None):
    for c in range(first_col, last_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name=FONT, size=10, bold=cell.font.bold if cell.font else False)
        if alt:
            cell.fill = FILL_ALT
        cell.border = BOX
        if c > first_col:
            cell.alignment = Alignment(horizontal="center", vertical="center")


def merge(ws, rng, value=None, **kw):
    ws.merge_cells(rng)
    if value is not None:
        put(ws, rng.split(":")[0], value, **kw)


def nick_formula(cell):
    return f'IFERROR(INDEX(Owners!$C$2:$C$100,MATCH({cell},Owners!$A$2:$A$100,0)),{cell})'


def rec_text(w, l, t):
    return f'{w}&"-"&{l}&"-"&{t}'


# ---------------------------------------------------------------- owner sheet
def build_owner_sheet(wb, title, oid, info, owners, teams, games, per, players, opp_order):
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = "2E75B6"
    ws.sheet_view.showGridLines = False
    widths = {"A": 30, "B": 9, "C": 8, "D": 8, "E": 8, "F": 11, "G": 11, "H": 9, "I": 24, "J": 30, "K": 12,
              "L": 6, "M": 6}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.column_dimensions["L"].hidden = True
    ws.column_dimensions["M"].hidden = True

    for c in range(1, 12):
        ws.cell(row=1, column=c).fill = FILL_TITLE
    put(ws, "A1", f"{info['nick']}  —  {info['full']}", bold=True, color="FFFFFF", fill=FILL_TITLE, size=16)
    ws.row_dimensions[1].height = 30
    put(ws, "A2", "OwnerID", italic=True, color="7F7F7F")
    put(ws, "B2", oid, italic=True, color="7F7F7F")
    ID = "$B$2"

    # ---- career summary -------------------------------------------------
    section(ws, 4, "CAREER SUMMARY")
    cw = lambda extra: f'COUNTIFS({TW_OWNER},{ID},{extra})'
    W_all, L_all, T_all = (cw(f'{TW_RES},"{r}"') for r in "WLT")
    W_reg, L_reg, T_reg = (cw(f'{TW_TYPE},"Regular",{TW_RES},"{r}"') for r in "WLT")
    W_po, L_po, T_po = (cw(f'{TW_TYPE},"Playoff*",{TW_RES},"{r}"') for r in "WLT")
    games_all = f"({W_all}+{L_all}+{T_all})"
    games_reg = f"({W_reg}+{L_reg}+{T_reg})"
    left = [
        ("Seasons played", f'=COUNTIF({TEAMS("C")},{ID})', None),
        ("Overall record (W-L-T)", f'={rec_text(W_all, L_all, T_all)}', None),
        ("Overall win %", f'=IFERROR(({W_all}+0.5*{T_all})/{games_all},"")', "0.0%"),
        ("Regular season (W-L-T)", f'={rec_text(W_reg, L_reg, T_reg)}', None),
        ("Regular season win %", f'=IFERROR(({W_reg}+0.5*{T_reg})/{games_reg},"")', "0.0%"),
        ("Playoff record (W-L)", f'={W_po}&"-"&{L_po}', None),
    ]
    right = [
        ("Championships", f'=COUNTIFS({TEAMS("C")},{ID},{TEAMS("N")},"Champion")', None),
        ("Runner-up finishes", f'=COUNTIFS({TEAMS("C")},{ID},{TEAMS("N")},"Runner-Up")', None),
        ("Playoff appearances", f'=COUNTIFS({TEAMS("C")},{ID},{TEAMS("M")},"Yes")', None),
        ("Best / worst finish",
         f'=IF(COUNTIFS({TEAMS("C")},{ID},{TEAMS("L")},">0")=0,"",MINIFS({TEAMS("L")},{TEAMS("C")},{ID})&" / "&MAXIFS({TEAMS("L")},{TEAMS("C")},{ID}))', None),
        ("Points for / against",
         f'=TEXT(SUMIFS({TW_PF},{TW_OWNER},{ID}),"#,##0")&" / "&TEXT(SUMIFS({TW_PA},{TW_OWNER},{ID}),"#,##0")', None),
        ("Average points per game", f'=IFERROR(SUMIFS({TW_PF},{TW_OWNER},{ID})/{games_all},"")', "0.0"),
    ]
    for i, ((l1, f1, n1), (l2, f2, n2)) in enumerate(zip(left, right)):
        r = 5 + i
        put(ws, f"A{r}", l1, bold=True)
        merge(ws, f"B{r}:C{r}", f1, fmt=n1, align="center")
        merge(ws, f"F{r}:H{r}", l2, bold=True)
        put(ws, f"I{r}", f2, fmt=n2, align="center")
        for c in range(1, 12):
            ws.cell(row=r, column=c).border = BOX
    row = 12

    # ---- season by season ----------------------------------------------------
    section(ws, row, "SEASON BY SEASON  (record = regular season)")
    row += 1
    head_row(ws, row, ["Team name", "Year", "W", "L", "T", "Points for", "Points against", "Finish",
                       "Playoff result", "Round-1 draft pick", "Title"])
    ws.row_dimensions[row].height = 30
    years = sorted({v[0] for k, v in teams.items() if v[1] == oid})
    first_season_row = row + 1
    for i, y in enumerate(years):
        row += 1
        key = f'$B{row}&"-"&{ID}'
        m = f'MATCH({key},{TEAMS("A")},0)'
        pull = lambda col: f'=IFERROR(IF(INDEX({TEAMS(col)},{m})="","",INDEX({TEAMS(col)},{m})),"")'
        put(ws, f"A{row}", pull("D"))
        put(ws, f"B{row}", y)
        for col, tcol, fmt in (("C", "F", None), ("D", "G", None), ("E", "H", None),
                               ("F", "I", "#,##0.0"), ("G", "J", "#,##0.0"), ("H", "L", None), ("I", "N", None)):
            put(ws, f"{col}{row}", pull(tcol), fmt=fmt)
        # helper: row of this team's round-1 pick in Draft (hidden col L)
        put(ws, f"L{row}", f'=SUMPRODUCT(MAX(({DRAFT("E")}={key})*({DRAFT("C")}=1)*ROW({DRAFT("E")})))')
        put(ws, f"J{row}", f'=IF(L{row}=0,"",INDEX(Draft!$F:$F,L{row})&" (#"&INDEX(Draft!$D:$D,L{row})&")")')
        put(ws, f"K{row}", f'=IF(I{row}="Champion","🏆","")', align="center")
        style_body(ws, row, 1, 11, alt=bool(i % 2))
        ws[f"A{row}"].alignment = Alignment(horizontal="left")
        ws[f"J{row}"].alignment = Alignment(horizontal="left")
    last_season_row = row
    row += 1
    put(ws, f"A{row}", "Total", bold=True)
    for col in "CDEFG":
        put(ws, f"{col}{row}", f"=SUM({col}{first_season_row}:{col}{last_season_row})", bold=True,
            fmt="#,##0.0" if col in "FG" else None, align="center")
    for c in range(1, 12):
        ws.cell(row=row, column=c).border = Border(top=Side(style="thin", color="000000"))
    ws.conditional_formatting.add(
        f"A{first_season_row}:K{last_season_row}",
        FormulaRule(formula=[f'$I{first_season_row}="Champion"'], fill=FILL_GOLD))
    row += 2

    # ---- head to head --------------------------------------------------------
    section(ws, row, "HEAD-TO-HEAD  (all games, regular season + playoffs)")
    row += 1
    head_row(ws, row, ["Opponent", "Games", "W", "L", "T", "Win %", "Points for", "Points against",
                       "Playoff meetings (W-L)"])
    ws.row_dimensions[row].height = 30
    first = row + 1
    for i, other in enumerate(opp_order):
        row += 1
        put(ws, f"L{row}", other)
        put(ws, f"A{row}", "=" + nick_formula(f"$L{row}"))
        cnt = lambda r: f'COUNTIFS({TW_OWNER},{ID},{TW_OPP},$L{row},{TW_RES},"{r}")'
        put(ws, f"C{row}", "=" + cnt("W"))
        put(ws, f"D{row}", "=" + cnt("L"))
        put(ws, f"E{row}", "=" + cnt("T"))
        put(ws, f"B{row}", f"=C{row}+D{row}+E{row}")
        put(ws, f"F{row}", f'=IFERROR((C{row}+0.5*E{row})/B{row},"")', fmt="0.0%")
        put(ws, f"G{row}", f'=SUMIFS({TW_PF},{TW_OWNER},{ID},{TW_OPP},$L{row})', fmt="#,##0.0")
        put(ws, f"H{row}", f'=SUMIFS({TW_PA},{TW_OWNER},{ID},{TW_OPP},$L{row})', fmt="#,##0.0")
        pw = f'COUNTIFS({TW_OWNER},{ID},{TW_OPP},$L{row},{TW_TYPE},"Playoff*",{TW_RES},"W")'
        pl = f'COUNTIFS({TW_OWNER},{ID},{TW_OPP},$L{row},{TW_TYPE},"Playoff*",{TW_RES},"L")'
        put(ws, f"I{row}", f'=IF({pw}+{pl}=0,"",{pw}&"-"&{pl})')
        style_body(ws, row, 1, 9, alt=bool(i % 2))
        ws[f"A{row}"].alignment = Alignment(horizontal="left")
    last = row
    if last >= first:
        ws.conditional_formatting.add(
            f"F{first}:F{last}", FormulaRule(formula=[f'AND(ISNUMBER(F{first}),F{first}>0.5)'],
                                              font=Font(name=FONT, color="1E7B34", bold=True)))
        ws.conditional_formatting.add(
            f"F{first}:F{last}", FormulaRule(formula=[f'AND(ISNUMBER(F{first}),F{first}<0.5)'],
                                              font=Font(name=FONT, color="C00000", bold=True)))
    row += 2

    # ---- personal records ---------------------------------------------------------
    section(ws, row, "PERSONAL RECORDS")
    row += 1
    head_row(ws, row, ["Record", "Score / margin", "Year", "Week", "Opponent", "", "", "Opp. score"])
    ws.merge_cells(f"E{row}:G{row}")
    recs = [
        ("Highest score", f'=MAXIFS({TW_PF},{TW_OWNER},{ID})', TW_PF),
        ("Lowest score", f'=MINIFS({TW_PF},{TW_OWNER},{ID},{TW_PF},">0",{TW_PA},">0")', TW_PF),
        ("Biggest win (margin)", f'=MAXIFS({TW_MARGIN},{TW_OWNER},{ID},{TW_PF},">0",{TW_PA},">0")', TW_MARGIN),
        ("Biggest loss (margin)", f'=MINIFS({TW_MARGIN},{TW_OWNER},{ID},{TW_PF},">0",{TW_PA},">0")', TW_MARGIN),
    ]
    for i, (lab, f, rng) in enumerate(recs):
        row += 1
        put(ws, f"A{row}", lab, bold=True)
        put(ws, f"B{row}", f, fmt="0.00")
        put(ws, f"M{row}", f'=SUMPRODUCT(MAX(({TW_OWNER}={ID})*({rng}=B{row})*ROW({TW_OWNER})))')
        idx = lambda col: f'=IF($M{row}=0,"",INDEX(Team_Weeks!${col}:${col},$M{row}))'
        put(ws, f"C{row}", idx("B"))
        put(ws, f"D{row}", idx("C"))
        merge(ws, f"E{row}:G{row}", f'=IF($M{row}=0,"",{nick_formula(f"INDEX(Team_Weeks!$L:$L,$M{row})")})', align="center")
        put(ws, f"H{row}", idx("H"), fmt="0.00")
        style_body(ws, row, 1, 8, alt=bool(i % 2))
        ws[f"A{row}"].alignment = Alignment(horizontal="left")
    row += 2

    # ---- top players -------------------------------------------------------------
    d = per.get(oid)
    section(ws, row, "TOP PLAYERS — ALL TIME  (points counted while started)" + ("  [kickers & defenses excluded]" if EXCLUDE_K_DST else ""))
    row += 1
    head_row(ws, row, ["", "Player", "", "", "", "Total", "", "Details", "", "", ""])
    ws.merge_cells(f"B{row}:E{row}")
    ws.merge_cells(f"F{row}:G{row}")
    ws.merge_cells(f"H{row}:K{row}")
    lines = []
    if d and d["roster"]:
        pid, n = max(d["starts"].items(), key=lambda kv: (kv[1], d["pts"][kv[0]])) if d["starts"] else (None, 0)
        if pid:
            lines.append(("Most-started player", pname(players, pid), f"{n} starts",
                          f"{yrs(d['years'][pid])} · {d['pts'][pid]:,.1f} pts as a starter"))
        pid, n = max(d["roster"].items(), key=lambda kv: (kv[1], d["starts"][kv[0]]))
        lines.append(("Most games on roster", pname(players, pid), f"{n} games",
                      f"{yrs(d['years'][pid])} · started {d['starts'][pid]} of them"))
        if d["pts"]:
            pid, pts = max(d["pts"].items(), key=lambda kv: kv[1])
            lines.append(("Most points scored (as starter)", pname(players, pid), f"{pts:,.1f} pts",
                          f"{yrs(d['years'][pid])} · {d['starts'][pid]} starts"))
        if d["best"]:
            pts, pid, y, w, o2 = max(d["best"], key=lambda g: g[0])
            lines.append(("Best single game", pname(players, pid), f"{pts:,.1f} pts",
                          f"{y} week {w}" + (f" vs {owners.get(o2, {}).get('nick', o2)}" if o2 else "")))
    for i, (lab, name, tot, det) in enumerate(lines):
        row += 1
        put(ws, f"A{row}", lab, bold=True)
        merge(ws, f"B{row}:E{row}", name)
        merge(ws, f"F{row}:G{row}", tot, align="center")
        merge(ws, f"H{row}:K{row}", det)
        for c in range(1, 12):
            ws.cell(row=row, column=c).border = BOX
            if i % 2:
                ws.cell(row=row, column=c).fill = FILL_ALT
    row += 2

    section(ws, row, "TOP PLAYERS — BY YEAR")
    row += 1
    head_row(ws, row, ["Year", "Most-started player", "", "", "Starts", "Most points scored", "", "", "Points",
                       "Best single game"])
    ws.merge_cells(f"B{row}:D{row}")
    ws.merge_cells(f"F{row}:H{row}")
    ws.row_dimensions[row].height = 30
    for i, y in enumerate(years):
        if not d or y not in d["y_starts"]:
            continue
        row += 1
        pid, n = max(d["y_starts"][y].items(), key=lambda kv: (kv[1], d["y_pts"][y][kv[0]]))
        ppid, ppts = max(d["y_pts"][y].items(), key=lambda kv: kv[1]) if d["y_pts"][y] else (None, 0)
        best = max(d["y_best"][y], key=lambda g: g[0]) if d["y_best"][y] else None
        put(ws, f"A{row}", y, bold=True)
        merge(ws, f"B{row}:D{row}", pname(players, pid))
        put(ws, f"E{row}", n, align="center")
        merge(ws, f"F{row}:H{row}", pname(players, ppid) if ppid else "")
        put(ws, f"I{row}", round(ppts, 1) if ppid else "", fmt="#,##0.0", align="center")
        if best:
            put(ws, f"J{row}", f"{pname(players, best[1])} · {best[0]:,.1f} (wk {best[3]})")
        for c in range(1, 11):
            ws.cell(row=row, column=c).border = BOX
            if i % 2:
                ws.cell(row=row, column=c).fill = FILL_ALT
        ws[f"A{row}"].alignment = Alignment(horizontal="left")
    ws.freeze_panes = "A3"
    return ws


# ---------------------------------------------------------------- shared sheets
def build_leaderboard(wb, order, titles, owners):
    ws = wb.create_sheet("Leaderboard")
    ws.sheet_properties.tabColor = "C00000"
    ws.sheet_view.showGridLines = False
    for c in range(1, 13):
        ws.cell(row=1, column=c).fill = FILL_TITLE
    put(ws, "A1", f"{LEAGUE_NAME} — All-Time Leaderboard", bold=True, color="FFFFFF", fill=FILL_TITLE, size=16)
    ws.row_dimensions[1].height = 30
    put(ws, "A2", "Click a name to open that owner's tab. Records include regular season and playoffs.",
        italic=True, color="7F7F7F")
    labels = ["Owner", "Seasons", "Titles", "Runner-ups", "Playoff apps", "W", "L", "T", "Win %",
              "Points for", "Avg pts/game", "OwnerID"]
    head_row(ws, 4, labels)
    ws.row_dimensions[4].height = 30
    for col, w in zip("ABCDEFGHIJKL", (22, 9, 8, 11, 12, 7, 7, 6, 9, 12, 13, 12)):
        ws.column_dimensions[col].width = w
    ws.column_dimensions["L"].hidden = True
    for i, oid in enumerate(order):
        r = 5 + i
        put(ws, f"L{r}", oid)
        put(ws, f"A{r}", f'=HYPERLINK("#\'{titles[oid]}\'!A1","{owners[oid]["nick"]}")', color="0563C1")
        put(ws, f"B{r}", f'=COUNTIF({TEAMS("C")},$L{r})')
        put(ws, f"C{r}", f'=COUNTIFS({TEAMS("C")},$L{r},{TEAMS("N")},"Champion")')
        put(ws, f"D{r}", f'=COUNTIFS({TEAMS("C")},$L{r},{TEAMS("N")},"Runner-Up")')
        put(ws, f"E{r}", f'=COUNTIFS({TEAMS("C")},$L{r},{TEAMS("M")},"Yes")')
        for col, res in (("F", "W"), ("G", "L"), ("H", "T")):
            put(ws, f"{col}{r}", f'=COUNTIFS({TW_OWNER},$L{r},{TW_RES},"{res}")')
        put(ws, f"I{r}", f'=IFERROR((F{r}+0.5*H{r})/(F{r}+G{r}+H{r}),"")', fmt="0.0%")
        put(ws, f"J{r}", f'=SUMIFS({TW_PF},{TW_OWNER},$L{r})', fmt="#,##0")
        put(ws, f"K{r}", f'=IFERROR(J{r}/(F{r}+G{r}+H{r}),"")', fmt="0.0")
        style_body(ws, r, 1, 11, alt=bool(i % 2))
        ws[f"A{r}"].alignment = Alignment(horizontal="left")
    ws.freeze_panes = "B5"
    return ws


def build_h2h(wb, order, owners):
    ws = wb.create_sheet("H2H_Matrix")
    ws.sheet_properties.tabColor = "C00000"
    ws.sheet_view.showGridLines = False
    n = len(order)
    for c in range(1, n + 3):
        ws.cell(row=1, column=c).fill = FILL_TITLE
    put(ws, "A1", "Head-to-Head Matrix  (row owner's record vs column owner, all games)", bold=True,
        color="FFFFFF", fill=FILL_TITLE, size=14)
    ws.row_dimensions[1].height = 28
    ws.column_dimensions["A"].width = 16
    put(ws, "A3", "Row vs column →", italic=True, color="7F7F7F")
    for j, o in enumerate(order):
        col = get_column_letter(2 + j)
        ws.column_dimensions[col].width = 9
        c = put(ws, f"{col}3", owners[o]["nick"], bold=True, fill=FILL_HEAD, align="center", wrap=True)
        put(ws, f"{col}2", o, color="FFFFFF", size=6)          # hidden-ish id row
    ws.row_dimensions[3].height = 30
    ws.row_dimensions[2].hidden = True
    ws.column_dimensions[get_column_letter(n + 3)].hidden = True
    for i, o in enumerate(order):
        r = 4 + i
        put(ws, f"{get_column_letter(n + 3)}{r}", o)
        put(ws, f"A{r}", owners[o]["nick"], bold=True, fill=FILL_HEAD)
        idc = f"${get_column_letter(n + 3)}{r}"
        for j, o2 in enumerate(order):
            col = get_column_letter(2 + j)
            if o == o2:
                put(ws, f"{col}{r}", "—", align="center", color="BFBFBF")
                continue
            w = f'COUNTIFS({TW_OWNER},{idc},{TW_OPP},{col}$2,{TW_RES},"W")'
            l = f'COUNTIFS({TW_OWNER},{idc},{TW_OPP},{col}$2,{TW_RES},"L")'
            t = f'COUNTIFS({TW_OWNER},{idc},{TW_OPP},{col}$2,{TW_RES},"T")'
            put(ws, f"{col}{r}", f'=IF({w}+{l}+{t}=0,"",{w}&"-"&{l}' + f'&IF({t}>0,"-"&{t},""))', align="center")
    rng = f"B4:{get_column_letter(n + 1)}{n + 3}"
    for sign, color in ((">", "C6EFCE"), ("<", "FFC7CE")):
        # green if more wins than losses, red if fewer
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'AND(B4<>"",B4<>"—",VALUE(LEFT(B4,FIND("-",B4)-1)){sign}VALUE(MID(B4,FIND("-",B4)+1,FIND("-",B4&"-",FIND("-",B4)+1)-FIND("-",B4)-1)))'],
            fill=PatternFill("solid", fgColor=color)))
    ws.freeze_panes = "B4"
    return ws


def build_records(wb):
    ws = wb.create_sheet("League_Records")
    ws.sheet_properties.tabColor = "C00000"
    ws.sheet_view.showGridLines = False
    for c in range(1, 9):
        ws.cell(row=1, column=c).fill = FILL_TITLE
    put(ws, "A1", "League Records — all seasons", bold=True, color="FFFFFF", fill=FILL_TITLE, size=16)
    ws.row_dimensions[1].height = 30
    for col, w in zip("ABCDEFGH", (6, 22, 10, 8, 8, 22, 11, 26)):
        ws.column_dimensions[col].width = w
    ws.column_dimensions["I"].hidden = True
    blocks = [("TOP 10 HIGHEST SINGLE-GAME SCORES", TW_PF, False), ("TOP 10 BIGGEST MARGINS OF VICTORY", TW_MARGIN, True)]
    row = 3
    for title, rng, margin in blocks:
        section(ws, row, title, last_col=8)
        row += 1
        head_row(ws, row, ["#", "Owner", "Score" if not margin else "Margin", "Year", "Week", "Opponent",
                           "Opp. score", "Game type"])
        for n in range(1, 11):
            row += 1
            put(ws, f"A{row}", n, align="center")
            put(ws, f"C{row}", f'=LARGE({rng},{n})', fmt="0.00", align="center")
            put(ws, f"I{row}", f'=SUMPRODUCT(MAX(({rng}=C{row})*ROW({rng})))')
            ix = lambda col: f'INDEX(Team_Weeks!${col}:${col},$I{row})'
            put(ws, f"B{row}", "=" + nick_formula(ix("K")))
            put(ws, f"D{row}", "=" + ix("B"), align="center")
            put(ws, f"E{row}", "=" + ix("C"), align="center")
            put(ws, f"F{row}", "=" + nick_formula(ix("L")))
            put(ws, f"G{row}", "=" + ix("H"), fmt="0.00", align="center")
            put(ws, f"H{row}", "=" + ix("D"))
            for c in range(1, 9):
                ws.cell(row=row, column=c).border = BOX
                if n % 2:
                    ws.cell(row=row, column=c).fill = FILL_ALT
                if c in (2, 6, 8):
                    ws.cell(row=row, column=c).font = Font(name=FONT, size=10)
        row += 2
    return ws


# ---------------------------------------------------------------- main
def new_names(order, titles):
    return GENERATED_SHARED + [titles[o] for o in order]


def main():
    if os.path.exists(WORKBOOK_PATH):
        bak = WORKBOOK_PATH.replace(".xlsx", "_backup_before_owner_tabs.xlsx")
        shutil.copy2(WORKBOOK_PATH, bak)
        print(f"Backup saved: {bak}")
    wb = schema_io.load_workbook(WORKBOOK_PATH)

    for oid, (full, nick) in OWNER_NAMES.items():
        schema_io.update_owner_names(wb, oid, full, nick)
    owners, teams, games, players, roster = load_tables(wb)
    print(f"{len(owners)} owners, {len(games)} games, {len(roster)} lineup rows")

    # order: most titles first, then overall win %
    rec = defaultdict(lambda: [0, 0, 0])
    for y, w, gt, a, sa, b, sb in games:
        sa, sb = sa or 0, sb or 0
        oa, ob = owner_of(a), owner_of(b)
        if sa > sb:
            rec[oa][0] += 1; rec[ob][1] += 1
        elif sb > sa:
            rec[ob][0] += 1; rec[oa][1] += 1
        else:
            rec[oa][2] += 1; rec[ob][2] += 1
    titles_won = Counter()
    for r in wb["Teams"].iter_rows(min_row=2, values_only=True):
        if r[2] and r[13] == "Champion":
            titles_won[r[2]] += 1
    played = {v[1] for v in teams.values()}

    def pct(o):
        w, l, t = rec[o]
        return (w + 0.5 * t) / max(1, w + l + t)

    order = sorted((o for o in owners if o in played), key=lambda o: (-titles_won[o], -pct(o)))
    # delete tabs from a previous run (shared sheets + any tab named after an owner nickname)
    core = {"Lookups", "Owners", "Seasons", "Teams", "Matchups", "Team_Weeks", "Players", "README",
            "Weekly_Rosters", "Draft"}
    prev = set(GENERATED_SHARED) | {re.sub(r"[\[\]:*?/\\]", "", owners[o]["nick"])[:31] for o in owners}
    for name in list(wb.sheetnames):
        if name in prev and name not in core:
            del wb[name]
    used = {n.lower() for n in wb.sheetnames}
    titles = {o: safe_title(owners[o]["nick"], used) for o in order}

    schema_io.rebuild_team_weeks(wb)            # adds the OwnerID / OppOwnerID helper columns
    per = player_stats(owners, teams, games, players, roster)

    opp_count = defaultdict(Counter)
    for y, w, gt, a, sa, b, sb in games:
        opp_count[owner_of(a)][owner_of(b)] += 1
        opp_count[owner_of(b)][owner_of(a)] += 1

    build_leaderboard(wb, order, titles, owners)
    build_h2h(wb, order, owners)
    build_records(wb)
    for oid in order:
        opp_order = [o for o, _ in opp_count[oid].most_common() if o != oid]
        build_owner_sheet(wb, titles[oid], oid, owners[oid], owners, teams, games, per, players, opp_order)
        print(f"  built tab {titles[oid]!r}")

    # newer Excel functions must carry the _xlfn. prefix inside the file or Excel shows #NAME?
    for n in new_names(order, titles):
        for row in wb[n].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    c.value = re.sub(r"(?<![\w.])(MAXIFS|MINIFS)\(", r"_xlfn.\1(", c.value)

    # new sheets to the front (Lookups stays hidden first), shared sheets then owners
    new = GENERATED_SHARED + [titles[o] for o in order]
    others = [ws for ws in wb._sheets if ws.title not in new]
    wb._sheets = [wb[n] for n in new] + others
    wb.active = 0
    wb.save(WORKBOOK_PATH)
    print(f"Saved {WORKBOOK_PATH}")
    print("Open it in Excel (or upload to Google Sheets) to let the formulas calculate.")


if __name__ == "__main__":
    main()

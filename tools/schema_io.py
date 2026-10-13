"""
Builds (if missing) and incrementally loads the FFL League Database workbook.
Owners / Seasons / Teams / Matchups get real data + live formulas written
per row as they're added, so re-running is safe (IDs are de-duplicated) and
standings/records keep recalculating in Excel/Sheets on open.

Players / Weekly_Rosters sheets are created (headers only) for a later pass.
"""
import os
from copy import copy

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter, column_index_from_string

FONT_NAME = "Arial"
MATCH_BOUND = 5000   # generous fixed upper bound baked into formulas, so we
TEAMS_BOUND = 1000   # never have to rewrite other sheets' formulas later

HEADER_INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
HEADER_CALC_FILL = PatternFill("solid", fgColor="D9E1F2")
HEADER_KEY_FILL = PatternFill("solid", fgColor="D6DCE4")
TITLE_FILL = PatternFill("solid", fgColor="1F4E78")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FONT = Font(name=FONT_NAME, size=10, bold=True)
TITLE_FONT = Font(name=FONT_NAME, size=14, bold=True, color="FFFFFF")
BODY_FONT = Font(name=FONT_NAME, size=10)

SHEET_COLUMNS = {
    "Owners": [("OwnerID", "key"), ("FullName", "input"), ("Nickname(s) / Past Team Names", "input"),
               ("ActiveSince (Year)", "input"), ("ActiveUntil (Year, blank = still active)", "input"),
               ("Notes", "input")],
    "Seasons": [("SeasonID (= Year)", "key"), ("Year", "input"), ("Platform", "input"),
                ("LeagueName", "input"), ("LeagueID (platform's league id)", "input"),
                ("NumTeams", "input"), ("RegSeasonWeeks (e.g. 1-13)", "input"),
                ("PlayoffStartWeek", "input"), ("PlayoffWeeks (e.g. 14-16)", "input"),
                ("NumPlayoffTeams", "input"), ("ScoringType", "input"), ("HasMedianGame", "input"),
                ("Champion (TeamSeasonID)", "calc"), ("Notes", "input")],
    "Teams": [("TeamSeasonID", "key"), ("SeasonID (Year)", "input"), ("OwnerID", "input"),
              ("TeamName", "input"), ("DivisionName", "input"),
              ("RegSeasonWins", "calc"), ("RegSeasonLosses", "calc"), ("RegSeasonTies", "calc"),
              ("RegSeasonPointsFor", "calc"), ("RegSeasonPointsAgainst", "calc"), ("WinPct", "calc"),
              ("FinalStanding", "input"), ("MadePlayoffs", "input"), ("PlayoffResult", "input"),
              ("Notes", "input")],
    "Matchups": [("MatchupID", "key"), ("SeasonID (Year)", "input"), ("Week", "input"), ("GameType", "input"),
                 ("TeamA_TeamSeasonID", "input"), ("TeamA_Score", "input"),
                 ("TeamB_TeamSeasonID", "input"), ("TeamB_Score", "input"),
                 ("Winner_TeamSeasonID", "calc"), ("Margin", "calc"), ("Notes", "input")],
    "Team_Weeks": [("TeamWeekID", "key"), ("SeasonID (Year)", "calc"), ("Week", "calc"), ("GameType", "calc"),
                   ("TeamSeasonID", "calc"), ("OpponentTeamSeasonID", "calc"), ("PointsFor", "calc"),
                   ("PointsAgainst", "calc"), ("Result", "calc"), ("Margin", "calc"),
                   ("OwnerID", "calc"), ("OppOwnerID", "calc")],
    "Players": [("PlayerID", "input"), ("PlayerName", "input"), ("PrimaryPosition", "input"),
                ("YahooPlayerID", "input"), ("SleeperPlayerID", "input"), ("Notes", "input")],
    "Weekly_Rosters": [("RosterEntryID", "key"), ("SeasonID (Year)", "input"), ("Week", "input"),
                        ("TeamSeasonID", "input"), ("PlayerID", "input"), ("RosterSlot", "input"),
                        ("IsStarter", "input"), ("PlayerPoints", "input"), ("Notes", "input")],
    "Draft": [("DraftPickID", "key"), ("SeasonID (Year)", "input"), ("Round", "input"),
              ("PickInRound", "input"), ("TeamSeasonID", "input"), ("PlayerName", "input"),
              ("Notes", "input")],
}

LOOKUPS = {
    "A": ("Platform", ["Yahoo", "Sleeper"]),
    "B": ("ScoringType", ["Standard", "Half-PPR", "PPR", "Custom"]),
    "C": ("YesNo", ["Yes", "No"]),
    "D": ("GameType", ["Regular", "Median", "Playoff - Quarterfinal", "Playoff - Semifinal",
                        "Playoff - Championship", "Playoff - 3rd Place", "Consolation"]),
    "E": ("PlayoffResult", ["Missed Playoffs", "Lost Wildcard/1st Round", "Lost Semifinal",
                             "Runner-Up", "Champion", "Won 3rd Place Game", "Lost 3rd Place Game",
                             "Consolation Champion", "Other"]),
    "F": ("RosterSlot", ["QB", "RB", "WR", "TE", "W/R/T FLEX", "RB/WR FLEX", "SUPERFLEX",
                          "DST", "K", "BN", "IR"]),
}


def _style_header(ws, col, text, kind):
    cell = ws.cell(row=1, column=col, value=text)
    cell.font = HEADER_FONT
    cell.fill = {"input": HEADER_INPUT_FILL, "calc": HEADER_CALC_FILL, "key": HEADER_KEY_FILL}[kind]
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    cell.border = BORDER


def build_workbook(path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    lk = wb.create_sheet("Lookups")
    lk_ranges = {}
    for col, (header, values) in LOOKUPS.items():
        cidx = column_index_from_string(col)
        c = lk.cell(row=1, column=cidx, value=header)
        c.font = HEADER_FONT
        c.fill = HEADER_KEY_FILL
        for i, v in enumerate(values, start=2):
            lk.cell(row=i, column=cidx, value=v).font = BODY_FONT
        lk_ranges[header] = f"Lookups!${col}$2:${col}${len(values) + 1}"
    lk.sheet_state = "hidden"

    for name, cols in SHEET_COLUMNS.items():
        ws = wb.create_sheet(name)
        for i, (h, k) in enumerate(cols, start=1):
            _style_header(ws, i, h, k)
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 30
        widths = [max(14, min(34, len(h) // 1 + 4)) for h, _ in cols]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w

    def add_dv(sheet, col_letter, source, last_row=3000):
        dv = DataValidation(type="list", formula1=source, allow_blank=True, showDropDown=False)
        wb[sheet].add_data_validation(dv)
        dv.add(f"{col_letter}2:{col_letter}{last_row}")

    add_dv("Seasons", "C", lk_ranges["Platform"])
    add_dv("Seasons", "K", lk_ranges["ScoringType"])
    add_dv("Seasons", "L", lk_ranges["YesNo"])
    add_dv("Teams", "M", lk_ranges["YesNo"])
    add_dv("Teams", "N", lk_ranges["PlayoffResult"])
    add_dv("Matchups", "D", lk_ranges["GameType"])
    add_dv("Weekly_Rosters", "F", lk_ranges["RosterSlot"])
    add_dv("Weekly_Rosters", "G", lk_ranges["YesNo"])

    ws = wb.create_sheet("README")
    wb.move_sheet("README", offset=-len(wb.sheetnames))
    ws.column_dimensions["A"].width = 110
    ws.sheet_view.showGridLines = False
    title = ws.cell(row=1, column=1, value="FANTASY FOOTBALL LEAGUE DATABASE")
    title.font = TITLE_FONT
    title.fill = TITLE_FILL
    lines = [
        "", "FFL Magnificos -- populated by ffl_extract.py from the Yahoo Fantasy API and the Sleeper API.",
        "Re-running the script is safe — it skips rows whose ID already exists.",
        "", "Header colors: YELLOW = real data (input). BLUE/GRAY = formula, auto-calculated.",
        "Team_Weeks is fully rebuilt from Matchups every run — never edit it directly.",
    ]
    for i, t in enumerate(lines, start=2):
        ws.cell(row=i, column=1, value=t).font = BODY_FONT

    wb.save(path)
    return wb


def load_workbook(path):
    if not os.path.exists(path):
        build_workbook(path)
    return openpyxl.load_workbook(path)


def _col(sheet, name):
    for i, (h, _) in enumerate(SHEET_COLUMNS[sheet], start=1):
        if h == name:
            return i
    raise KeyError(name)


def _existing_ids(ws, id_col=1, start_row=2):
    """Only valid for sheets whose ID column (col A) is a plain value, not a
    formula (Owners, Seasons, Players). For formula-ID sheets (Teams,
    Matchups, Weekly_Rosters) use the dedicated _existing_*_keys() helpers
    below instead — a formula cell's .value is its formula TEXT, not the
    computed ID, so comparing it to a computed ID string never matches."""
    ids = set()
    for row in ws.iter_rows(min_row=start_row, min_col=id_col, max_col=id_col, values_only=True):
        if row[0]:
            ids.add(row[0])
    return ids


def _next_row(ws, id_col=1, start_row=2):
    r = start_row
    while ws.cell(row=r, column=id_col).value not in (None, ""):
        r += 1
    return r


def _find_row(ws, predicate, start_row=2):
    for r in range(start_row, ws.max_row + 1):
        if predicate(r):
            return r
    return None


def team_season_id(year, owner_id):
    return f"{year}-{owner_id}"


def matchup_id(year, week, a, b):
    return f"{year}-W{week}-{a}-{b}"


def roster_entry_id(year, week, team_season_id_, player_id):
    return f"{year}-W{week}-{team_season_id_}-{player_id}"


def draft_pick_id(year, rnd, pick_in_round):
    return f"{year}-R{int(rnd):02d}-P{int(pick_in_round):02d}"


def ensure_draft_sheet(wb):
    """Adds the Draft sheet (with the same header styling build_workbook
    uses) to a workbook that was created before this sheet existed."""
    if "Draft" in wb.sheetnames:
        return
    cols = SHEET_COLUMNS["Draft"]
    ws = wb.create_sheet("Draft")
    for i, (h, k) in enumerate(cols, start=1):
        _style_header(ws, i, h, k)
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 30
    for i, (h, _) in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(14, min(34, len(h) + 4))


def upsert_owner(wb, owner_id, full_name, nickname="", active_since="", active_until="", notes=""):
    ws = wb["Owners"]
    if owner_id in _existing_ids(ws):
        return
    r = _next_row(ws)
    vals = [owner_id, full_name, nickname, active_since, active_until, notes]
    for c, v in enumerate(vals, start=1):
        ws.cell(row=r, column=c, value=v)


def upsert_season(wb, year, platform, league_name, league_id, num_teams, reg_weeks, playoff_start,
                   playoff_weeks, num_playoff_teams, scoring_type, has_median, notes=""):
    ws = wb["Seasons"]
    r = _find_row(ws, lambda row: ws.cell(row=row, column=1).value == year)
    if r:
        # season settings can still be in flux for an in-progress season
        # (e.g. the estimated playoff week range) -- keep them current
        vals = {6: num_teams, 7: reg_weeks, 8: playoff_start, 9: playoff_weeks,
                10: num_playoff_teams, 11: scoring_type}
        for c, v in vals.items():
            if v not in ("", None):
                ws.cell(row=r, column=c, value=v)
        return
    r = _next_row(ws)
    vals = [year, year, platform, league_name, league_id, num_teams, reg_weeks, playoff_start,
            playoff_weeks, num_playoff_teams, scoring_type, has_median]
    for c, v in enumerate(vals, start=1):
        ws.cell(row=r, column=c, value=v)
    ws.cell(row=r, column=13, value=(
        f'=IFERROR(INDEX(Teams!$A$2:$A${TEAMS_BOUND},'
        f'SUMPRODUCT((Teams!$B$2:$B${TEAMS_BOUND}=$A{r})*(Teams!$N$2:$N${TEAMS_BOUND}="Champion")'
        f'*ROW(Teams!$B$2:$B${TEAMS_BOUND}))-1),"")'
    ))
    ws.cell(row=r, column=14, value=notes)


def upsert_team(wb, year, owner_id, team_name, division="", final_standing="", made_playoffs="",
                playoff_result="", notes=""):
    tsid = team_season_id(year, owner_id)
    ws = wb["Teams"]
    r = _find_row(ws, lambda row: ws.cell(row=row, column=2).value == year
                  and ws.cell(row=row, column=3).value == owner_id)
    if r:
        # update the input columns in place (standings/notes can change week to week)
        for c, v in ((5, division), (12, final_standing), (13, made_playoffs), (14, playoff_result)):
            if v not in ("", None):
                ws.cell(row=r, column=c, value=v)
        return tsid
    r = _next_row(ws)
    ws.cell(row=r, column=1, value=f'=IF(OR($B{r}="",$C{r}=""),"",$B{r}&"-"&$C{r})')
    vals = {2: year, 3: owner_id, 4: team_name, 5: division, 12: final_standing,
            13: made_playoffs, 14: playoff_result, 15: notes}
    for c, v in vals.items():
        if v not in ("", None):
            ws.cell(row=r, column=c, value=v)
    ws.cell(row=r, column=6, value=(
        f'=SUMPRODUCT((Matchups!$B$2:$B${MATCH_BOUND}=$B{r})*(Matchups!$D$2:$D${MATCH_BOUND}="Regular")*'
        f'(((Matchups!$E$2:$E${MATCH_BOUND}=$A{r})*(Matchups!$F$2:$F${MATCH_BOUND}>Matchups!$H$2:$H${MATCH_BOUND}))+'
        f'((Matchups!$G$2:$G${MATCH_BOUND}=$A{r})*(Matchups!$H$2:$H${MATCH_BOUND}>Matchups!$F$2:$F${MATCH_BOUND}))))'
    ))
    ws.cell(row=r, column=7, value=(
        f'=SUMPRODUCT((Matchups!$B$2:$B${MATCH_BOUND}=$B{r})*(Matchups!$D$2:$D${MATCH_BOUND}="Regular")*'
        f'(((Matchups!$E$2:$E${MATCH_BOUND}=$A{r})*(Matchups!$F$2:$F${MATCH_BOUND}<Matchups!$H$2:$H${MATCH_BOUND}))+'
        f'((Matchups!$G$2:$G${MATCH_BOUND}=$A{r})*(Matchups!$H$2:$H${MATCH_BOUND}<Matchups!$F$2:$F${MATCH_BOUND}))))'
    ))
    ws.cell(row=r, column=8, value=(
        f'=SUMPRODUCT((Matchups!$B$2:$B${MATCH_BOUND}=$B{r})*(Matchups!$D$2:$D${MATCH_BOUND}="Regular")*'
        f'(((Matchups!$E$2:$E${MATCH_BOUND}=$A{r})*(Matchups!$F$2:$F${MATCH_BOUND}=Matchups!$H$2:$H${MATCH_BOUND}))+'
        f'((Matchups!$G$2:$G${MATCH_BOUND}=$A{r})*(Matchups!$H$2:$H${MATCH_BOUND}=Matchups!$F$2:$F${MATCH_BOUND}))))'
    ))
    ws.cell(row=r, column=9, value=(
        f'=SUMPRODUCT((Matchups!$B$2:$B${MATCH_BOUND}=$B{r})*(Matchups!$D$2:$D${MATCH_BOUND}="Regular")*'
        f'((Matchups!$E$2:$E${MATCH_BOUND}=$A{r})*Matchups!$F$2:$F${MATCH_BOUND}+'
        f'(Matchups!$G$2:$G${MATCH_BOUND}=$A{r})*Matchups!$H$2:$H${MATCH_BOUND}))'
    ))
    ws.cell(row=r, column=10, value=(
        f'=SUMPRODUCT((Matchups!$B$2:$B${MATCH_BOUND}=$B{r})*(Matchups!$D$2:$D${MATCH_BOUND}="Regular")*'
        f'((Matchups!$E$2:$E${MATCH_BOUND}=$A{r})*Matchups!$H$2:$H${MATCH_BOUND}+'
        f'(Matchups!$G$2:$G${MATCH_BOUND}=$A{r})*Matchups!$F$2:$F${MATCH_BOUND}))'
    ))
    ws.cell(row=r, column=11, value=f'=IFERROR((F{r}+0.5*H{r})/(F{r}+G{r}+H{r}),"")')
    ws.cell(row=r, column=11).number_format = "0.0%"
    return tsid


def upsert_matchup(wb, year, week, game_type, team_a, score_a, team_b, score_b, notes=""):
    mid = matchup_id(year, week, team_a, team_b)

    def _matches(row):
        bv, cv, ev, gv = (ws.cell(row=row, column=c).value for c in (2, 3, 5, 7))
        return bv == year and cv == week and {ev, gv} == {team_a, team_b}

    ws = wb["Matchups"]
    r = _find_row(ws, _matches)
    if r:
        # scores can still be live when pulled mid-week -- refresh them
        a_in_slot_e = ws.cell(row=r, column=5).value == team_a
        ws.cell(row=r, column=6, value=score_a if a_in_slot_e else score_b)
        ws.cell(row=r, column=8, value=score_b if a_in_slot_e else score_a)
        return mid
    r = _next_row(ws)
    ws.cell(row=r, column=1, value=(
        f'=IF(OR($B{r}="",$C{r}="",$E{r}="",$G{r}=""),"",$B{r}&"-W"&$C{r}&"-"&$E{r}&"-"&$G{r})'
    ))
    vals = [None, year, week, game_type, team_a, score_a, team_b, score_b]
    for c, v in enumerate(vals, start=1):
        if v is not None:
            ws.cell(row=r, column=c, value=v)
    ws.cell(row=r, column=9, value=(
        f'=IF(OR($F{r}="",$H{r}=""),"",IF($F{r}>$H{r},$E{r},IF($H{r}>$F{r},$G{r},"Tie")))'
    ))
    ws.cell(row=r, column=10, value=f'=IF(OR($F{r}="",$H{r}=""),"",ABS($F{r}-$H{r}))')
    ws.cell(row=r, column=6).number_format = "0.00"
    ws.cell(row=r, column=8).number_format = "0.00"
    ws.cell(row=r, column=10).number_format = "0.00"
    ws.cell(row=r, column=11, value=notes)
    return mid


def rebuild_team_weeks(wb):
    ws = wb["Team_Weeks"]
    for col, label in ((11, "OwnerID"), (12, "OppOwnerID")):      # helper columns used by the owner tabs
        h = ws.cell(row=1, column=col)
        if not h.value:
            h.value = label
            src = ws.cell(row=1, column=10)
            h.font, h.fill, h.alignment, h.border = (copy(src.font), copy(src.fill),
                                                      copy(src.alignment), copy(src.border))
            ws.column_dimensions[get_column_letter(col)].width = 14
    for row in ws.iter_rows(min_row=2, max_row=max(ws.max_row, 2)):
        for cell in row:
            cell.value = None
    m = wb["Matchups"]
    last = m.max_row
    rows = [r for r in range(2, last + 1) if m.cell(row=r, column=1).value not in (None, "")]
    dest = 2
    for perspective in ("A", "B"):
        for mr in rows:
            # Yahoo "Bye Week" placeholder opponents (team id ending -BYE) get no row of their own
            if perspective == "B" and str(m.cell(row=mr, column=7).value or "").endswith("-BYE"):
                continue
            if perspective == "A":
                self_c, opp_c, pf_c, pa_c = "E", "G", "F", "H"
            else:
                self_c, opp_c, pf_c, pa_c = "G", "E", "H", "F"
            ws.cell(row=dest, column=2, value=f"=Matchups!$B{mr}")
            ws.cell(row=dest, column=3, value=f"=Matchups!$C{mr}")
            ws.cell(row=dest, column=4, value=f"=Matchups!$D{mr}")
            ws.cell(row=dest, column=5, value=f"=Matchups!${self_c}{mr}")
            ws.cell(row=dest, column=6, value=f"=Matchups!${opp_c}{mr}")
            ws.cell(row=dest, column=7, value=f"=Matchups!${pf_c}{mr}")
            ws.cell(row=dest, column=8, value=f"=Matchups!${pa_c}{mr}")
            ws.cell(row=dest, column=9, value=(
                f'=IF(OR(G{dest}="",H{dest}=""),"",IF(G{dest}>H{dest},"W",IF(G{dest}<H{dest},"L","T")))'
            ))
            ws.cell(row=dest, column=10, value=f'=IF(OR(G{dest}="",H{dest}=""),"",G{dest}-H{dest})')
            ws.cell(row=dest, column=11, value=f'=IF(E{dest}="","",MID(E{dest},6,50))')
            ws.cell(row=dest, column=12, value=f'=IF(F{dest}="","",MID(F{dest},6,50))')
            ws.cell(row=dest, column=1, value=(
                f'=IF(OR(E{dest}="",F{dest}="",B{dest}="",C{dest}=""),"",'
                f'B{dest}&"-W"&C{dest}&"-"&E{dest}&"-vs-"&F{dest})'
            ))
            dest += 1


def update_owner_names(wb, owner_id, full_name=None, nickname=None):
    """Correct an existing Owners row's display fields in place (e.g. a
    Sleeper-only owner whose real name/nickname is now known)."""
    ws = wb["Owners"]
    for r in range(2, ws.max_row + 1):
        if ws.cell(row=r, column=1).value == owner_id:
            if full_name is not None:
                ws.cell(row=r, column=2, value=full_name)
            if nickname is not None:
                ws.cell(row=r, column=3, value=nickname)
            return True
    return False


def clear_owner_row(wb, owner_id):
    """Blanks an Owners row whose identity turned out to be a duplicate of
    another (merged-away) owner."""
    ws = wb["Owners"]
    for r in range(2, ws.max_row + 1):
        if ws.cell(row=r, column=1).value == owner_id:
            for c in range(1, len(SHEET_COLUMNS["Owners"]) + 1):
                ws.cell(row=r, column=c).value = None  # cell(value=None) is a no-op
            return True
    return False


def reassign_team_owner_id(wb, old_owner_id, new_owner_id, year_max=None):
    """Identity-correction helper: repoints Teams rows (and the matching
    literal TeamSeasonID text in Draft) from old_owner_id to new_owner_id.
    year_max limits this to years <= year_max so a Sleeper-era owner who
    legitimately still uses old_owner_id for later seasons is untouched."""
    ws = wb["Teams"]
    n = 0
    for r in range(2, ws.max_row + 1):
        yv = ws.cell(row=r, column=2).value
        if ws.cell(row=r, column=3).value != old_owner_id:
            continue
        if year_max is not None and not (isinstance(yv, int) and yv <= year_max):
            continue
        ws.cell(row=r, column=3, value=new_owner_id)
        n += 1

    if "Draft" in wb.sheetnames:
        dws = wb["Draft"]
        for r in range(2, dws.max_row + 1):
            yv = dws.cell(row=r, column=2).value
            if not isinstance(yv, int):
                continue
            if year_max is not None and yv > year_max:
                continue
            if dws.cell(row=r, column=5).value == f"{yv}-{old_owner_id}":
                dws.cell(row=r, column=5, value=f"{yv}-{new_owner_id}")
    return n


def upsert_draft_pick(wb, year, rnd, pick_in_round, team_season_id_, player_name, notes=""):
    did = draft_pick_id(year, rnd, pick_in_round)
    ws = wb["Draft"]
    r = _find_row(ws, lambda row: ws.cell(row=row, column=2).value == year
                  and ws.cell(row=row, column=3).value == rnd
                  and ws.cell(row=row, column=4).value == pick_in_round)
    if r:
        ws.cell(row=r, column=5, value=team_season_id_)
        ws.cell(row=r, column=6, value=player_name)
        return did
    r = _next_row(ws)
    vals = [did, year, rnd, pick_in_round, team_season_id_, player_name, notes]
    for c, v in enumerate(vals, start=1):
        ws.cell(row=r, column=c, value=v)
    return did

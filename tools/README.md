# FFL Magnificos — League Database & Website

History, stats and website data for the **FFL Magnificos** fantasy football league
(10 teams): Yahoo 2016–2025 and Sleeper 2026 onward.

Folder: `C:\Users\jchap\Python\Yahoo\FFL_Magnificos`

---

## What's in the database

`FFL_Magnificos_Database.xlsx` holds:

| Sheet | Contents |
|---|---|
| Seasons | One row per season: platform, league name and ID, playoff format |
| Owners | 16 owners with a short OwnerID, real name and nickname |
| Teams | Each owner's team per season: record, points, final standing, playoff result |
| Matchups | Every game 2016–2026: regular season, playoffs and consolation |
| Team_Weeks | One row per team per week (rebuilt automatically, don't edit) |
| Players | Every player who appeared, linked across Yahoo and Sleeper |
| Weekly_Rosters | Starters, bench and IR with fantasy points for every team-week |
| Draft | Every draft pick 2016–2026 |

Checks run when the data was loaded:
- Regular-season records rebuilt from the scores match Yahoo's standings for all 100 Yahoo team-seasons.
- Starters' points add up to the final score in all 1,582 Yahoo team-weeks.

### Seasons

| Year | Platform | League name | League ID |
|---|---|---|---|
| 2016 | Yahoo | #*+%% | 359.l.955440 |
| 2017 | Yahoo | It doesn't matter | 371.l.642629 |
| 2018 | Yahoo | Jagger Tutus and a Prius | 380.l.37947 |
| 2019 | Yahoo | Jagger Tutus and a Prius | 390.l.288171 |
| 2020 | Yahoo | Best FFL Ever! | 399.l.62341 |
| 2021 | Yahoo | Best FFL Ever! | 406.l.697498 |
| 2022 | Yahoo | Best FFL Ever! | 414.l.87648 |
| 2023 | Yahoo | Best FFL and Ribs | 423.l.284024 |
| 2024 | Yahoo | Best FFL and Ribs | 449.l.248943 |
| 2025 | Yahoo | Best FFL and Ribs | 461.l.434599 |
| 2026 | Sleeper | Los Magnificos | 1390438454407467008 |

### Owners

| OwnerID | Name | Nickname | Notes |
|---|---|---|---|
| JESUS | Jesus | Chaps | Sleeper: jchaps17 |
| ALEJANDR | Alejandro | Lopez | Sleeper: alexaodesouza |
| AMIR | Amir | Amir | Includes 2016 "Todos Me La Pelan"; Sleeper: puentetunechi |
| CARLOS | Carlos | Callito | "A+ Mr.200"; Sleeper: callito |
| STRETCH | Carlos | Stretch | StretchLife, joined 2025; Sleeper: Stretch36 |
| EFRAIN | Efrain | E-Man | Sleeper: EMan79 |
| FELIPE | Felipe | Felipe | Sleeper: troquero |
| ISAI | Isai | Isai | Sleeper: isaij17 |
| SAL | Sal | Sal | Sleeper: Savastianiporto |
| ROBERTO | Roberto | Rooney | 3 Yahoo accounts + lagalaxy1515; Sleeper: hueroelchampion |
| ART | Art | Art | 2016–2024 (two Yahoo accounts) |
| JUAN | Juan | Juan | |
| JORGE | Jorge | Jorge | |
| ALEX | Alex | Barba | |
| MARK | Mark | Mark | |
| LRDCHNY | LRDCHNY | LRDCHNY | 2016 only |

To change a name or nickname, edit `OWNER_NAMES` in `league_config.py`.
The website's owner page address comes from the nickname.

---

## Kept separate from Camaradas

This league never touches the Camaradas League files.

- All settings are in **`league_config.py`**: league name, Yahoo keys, Sleeper ID, owner names and file paths.
- Every script stops with `STOP:` if a path points into `fantasy_env\ffl_pull` or `Camaradas_League_parquet`.
- Yahoo pulls only use the league keys listed in `league_config.py`.
- It has its own workbook, `owner_map.json`, Yahoo login copy (`oauth2.json`), player cache, money settings and rules.
- **Always `cd` into this folder before running a script.**
- Publish to its **own GitHub repo** (e.g. `ffl-magnificos`), never the Camaradas repo.

---

## Weekly routine (during the season)

```
C:\Users\jchap\fantasy_env\Scripts\activate
cd C:\Users\jchap\Python\Yahoo\FFL_Magnificos
python sleeper_backfill.py weekly
python fetch_projections.py        (Thursdays, for the Power Rankings page)
python export_site_data.py
```

Then on GitHub: **Add file → Upload files**, drag in the `site\data` folder and commit.
The site updates in about a minute.

Optional: `python build_owner_sheets.py` rebuilds the per-owner Excel tabs.

---

## Website

- `site\` holds the website (plain HTML/CSS/JS plus JSON data). Nothing to build.
- To preview locally, run `cd site`, then `python -m http.server 8000`, then open http://localhost:8000.
- One-time GitHub Pages setup:
  1. Create a new public repo.
  2. Upload everything *inside* `site`.
  3. Go to Settings → Pages → Deploy from branch `main`, folder `/ (root)`.
- **Prizes and rules:** `export_site_data.py` creates `league_settings.json` (money, all zeros for now) and
  `league_rules.md`. Edit them, then run the export again. Until then those pages show a "placeholder" notice.
- **Power Rankings** (`power.html`): `fetch_projections.py` saves Sleeper's projected points for every remaining week and each
  team's current roster (IR/taxi excluded) to `power_projections.json`. The export builds each team's best projected
  lineup per week using the league's own slots (QB, 2 RB, 2 WR, TE, 3 FLEX, K, DEF) and ranks the teams. The score
  weights are `powerWeights` in `league_settings.json`: weeks 1-14 = lineup 75, bench 5, record 15, points 5; playoffs
  (weeks 15-17) = lineup 100. Games already played are never projected.
- **Photos:** `python prepare_photos.py --public "C:\path\to\photos"` (public) or `--private` (password-protected).
  Name files like `2023_draft day.jpg`. Needs `pip install pillow cryptography` once.
- **Renamed owners:** if you change a nickname, delete `site\data\owner` before exporting, and delete the
  old owner file from the GitHub repo.

---

## Scripts

| Script | What it does |
|---|---|
| `league_config.py` | All league settings and the Camaradas isolation guard |
| `sleeper_backfill.py` | Pulls Sleeper seasons: scores, playoffs, drafts, lineups, players (`weekly` = current season only) |
| `ffl_extract.py` | Core Yahoo/Sleeper puller (`discover`, `full`, `weekly`); Yahoo needs API access (see below) |
| `load_yahoo_magnificos.py` | Loads Yahoo seasons, teams, owners, scores and standings from `yahoo_scraped_magnificos.json` + `teams.csv` |
| `load_yahoo_lineups_magnificos.py` | Loads Yahoo lineups and drafts from `yahoo_lineups_magnificos.json` / `yahoo_drafts_magnificos.json`, then links players |
| `link_players.py` | Merges Yahoo and Sleeper versions of the same player |
| `build_owner_sheets.py` | Builds owner tabs, leaderboard, head-to-head matrix and records in the workbook |
| `fetch_projections.py` | Pulls Sleeper projections and current rosters for the Power Rankings page (run Thursdays) |
| `export_site_data.py` | Turns the workbook into the website's JSON files in `site\data` |
| `prepare_photos.py` | Shrinks public photos and encrypts private ones |
| `reauth.py` | Gets a fresh Yahoo login token (`--scope` asks for Fantasy access explicitly) |

### Data files

| File | Source |
|---|---|
| `yahoo_scraped_magnificos.json` | Weekly scores and final standings from the Yahoo site, 2016–2025 |
| `yahoo_lineups_magnificos.json` | Starters, bench and IR with points for all 791 Yahoo games |
| `yahoo_drafts_magnificos.json` | 1,500 Yahoo draft picks, 2016–2025 |
| `yahoo_positions_magnificos.json` | Real position (QB/RB/WR/TE/K) for Yahoo players, from Sleeper's player list |
| `owner_map.json` | Yahoo GUID / Sleeper user ID → OwnerID |
| `sleeper_players_cache.json` | Trimmed Sleeper player list (auto-refreshed) |
| `teams.csv`, `leagues.csv` | Read-only, from `..\fantasy_football_parquet` |

---

## Rebuilding from scratch

The Yahoo history doesn't need the Yahoo API. It loads from the saved JSON files above.

```
python load_yahoo_magnificos.py              # seasons, teams, owners, scores, standings
python load_yahoo_lineups_magnificos.py      # lineups, drafts, player linking
python sleeper_backfill.py                   # Sleeper seasons (2026+)
python export_site_data.py
```

Every script backs up the workbook before changing it (`*_backup_before_*.xlsx`). All of them are safe to re-run.

---

## Known issues

- **Yahoo API:** returns 403 "application is not authorized". The Yahoo app needs **Fantasy Sports → Read**
  permission at developer.yahoo.com/apps, followed by `python reauth.py`. This isn't needed for the
  saved Yahoo history.
- **Yahoo rate limit:** reading many Yahoo pages quickly triggers a temporary "Request denied" block that lasts about an hour.

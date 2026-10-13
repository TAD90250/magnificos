"""
FFL Magnificos -- the ONE place league-specific settings live.

Every script in this folder imports from here. Nothing in this folder reads
or writes the Camaradas League files (C:\\Users\\jchap\\fantasy_env\\ffl_pull or
C:\\Users\\jchap\\Python\\Yahoo\\Camaradas_League_parquet) -- the guard at the
bottom stops any script cold if a path ever points there.
"""
import os

HERE = os.path.dirname(os.path.abspath(__file__))

# ------------------------------------------------------------ fill these in
LEAGUE_NAME = "FFL Magnificos"

# Yahoo seasons: {year: "league_key"}. Run `python ffl_extract.py discover`
# to list every Yahoo league on your login, then copy the Magnificos ones here.
# ONLY keys listed here are ever pulled.
YAHOO_LEAGUE_KEYS = {
    2016: "359.l.955440", 2017: "371.l.642629", 2018: "380.l.37947", 2019: "390.l.288171",
    2020: "399.l.62341", 2021: "406.l.697498", 2022: "414.l.87648", 2023: "423.l.284024",
    2024: "449.l.248943", 2025: "461.l.434599",
}

# Current Sleeper league ID (the URL number on sleeper.com/leagues/<ID>).
# Earlier Sleeper seasons are found automatically via previous_league_id.
SLEEPER_SEED_LEAGUE_ID = "1390438454407467008"   # 2026

# Real names / nicknames per OwnerID (filled after discover; optional).
#   "OWNERID": ("Full Name", "Nickname"),
OWNER_NAMES = {
    "JESUS": ("Jesus", "Chaps"), "ALEJANDR": ("Alejandro", "Lopez"), "AMIR": ("Amir", "Amir"),
    "CARLOS": ("Carlos", "Callito"), "STRETCH": ("Carlos", "Stretch"), "EFRAIN": ("Efrain", "E-Man"),
    "FELIPE": ("Felipe", "Felipe"), "ISAI": ("Isai", "Isai"), "SAL": ("Sal", "Sal"),
    "ROBERTO": ("Roberto", "Rooney"), "ART": ("Art", "Art"), "JUAN": ("Juan", "Juan"),
    "JORGE": ("Jorge", "Jorge"), "ALEX": ("Alex", "Barba"), "MARK": ("Mark", "Mark"),
    "LRDCHNY": ("LRDCHNY", "LRDCHNY"),
}

# ------------------------------------------------------------ files (this folder only)
OAUTH_PATH = os.path.join(HERE, "oauth2.json")            # this league's own copy
WORKBOOK_PATH = os.path.join(HERE, "FFL_Magnificos_Database.xlsx")
OWNER_MAP_PATH = os.path.join(HERE, "owner_map.json")
DEBUG_DIR = os.path.join(HERE, "debug_dumps")
SITE_DIR = os.path.join(HERE, "site")

# ------------------------------------------------------------ isolation guard
_PROTECTED = [
    r"C:\Users\jchap\fantasy_env\ffl_pull",
    r"C:\Users\jchap\Python\Yahoo\Camaradas_League_parquet",
]


def _norm(p):
    return os.path.normcase(os.path.abspath(p)).rstrip("\\/")


def assert_isolated(*paths):
    """Raise if any path is inside a Camaradas folder."""
    for p in paths:
        np_ = _norm(p)
        for prot in _PROTECTED:
            npr = _norm(prot)
            if np_ == npr or np_.startswith(npr + os.sep):
                raise SystemExit(f"STOP: {p} is inside the Camaradas League folder ({prot}). "
                                 "Magnificos scripts must run from their own folder.")


def require_league_ids(need_yahoo=True, need_sleeper=True):
    missing = []
    if need_yahoo and not YAHOO_LEAGUE_KEYS:
        missing.append("YAHOO_LEAGUE_KEYS")
    if need_sleeper and not SLEEPER_SEED_LEAGUE_ID:
        missing.append("SLEEPER_SEED_LEAGUE_ID")
    if missing:
        raise SystemExit("Fill in " + " and ".join(missing) + " in league_config.py first "
                         "(run `python ffl_extract.py discover` to find the Yahoo keys).")


assert_isolated(HERE, OAUTH_PATH, WORKBOOK_PATH, OWNER_MAP_PATH, DEBUG_DIR, SITE_DIR)

"""Sleeper public API — no auth required."""
import requests

BASE = "https://api.sleeper.app/v1"


def _get(path):
    r = requests.get(f"{BASE}/{path}", timeout=30)
    r.raise_for_status()
    return r.json()


def get_league(league_id):
    return _get(f"league/{league_id}")


def get_users(league_id):
    return _get(f"league/{league_id}/users")


def get_rosters(league_id):
    return _get(f"league/{league_id}/rosters")


def get_matchups(league_id, week):
    return _get(f"league/{league_id}/matchups/{week}")


def get_winners_bracket(league_id):
    try:
        return _get(f"league/{league_id}/winners_bracket")
    except Exception:
        return []


def get_losers_bracket(league_id):
    try:
        return _get(f"league/{league_id}/losers_bracket")
    except Exception:
        return []


def discover_chain(seed_league_id):
    """Walk previous_league_id backward from the seed to find every Sleeper
    season in this league's history. Returns {season(str year): league_id}."""
    chain = {}
    league_id = seed_league_id
    seen = set()
    while league_id and league_id not in seen and league_id != "0":
        seen.add(league_id)
        league = get_league(league_id)
        season = league.get("season")
        chain[season] = league_id
        league_id = league.get("previous_league_id")
    return chain

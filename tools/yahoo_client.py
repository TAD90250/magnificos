"""
Minimal Yahoo Fantasy Sports API client using raw OAuth2 (the yahoo_oauth /
yahoo_fantasy_api packages fail to build on some boxes — their 'rauth'
dependency doesn't install on modern setuptools). Reads/refreshes the token
file in place, same format yahoo_oauth uses, so it's a drop-in replacement
for oauth2.json.

Yahoo's JSON responses nest the same object's attributes across a list of
one-key dicts, and repeat elements under numeric string keys ("0","1",...,
"count") instead of a plain list. find_all() / merge_meta_list() below are
generic helpers for digging through that shape without hand-coding the exact
nesting for every endpoint (which Yahoo is not fully consistent about).
"""
import json
import time
import requests

TOKEN_URL = "https://api.login.yahoo.com/oauth2/get_token"
BASE = "https://fantasysports.yahooapis.com/fantasy/v2"


class YahooClient:
    def __init__(self, token_path):
        self.token_path = token_path
        with open(token_path) as f:
            self.tok = json.load(f)
        # Yahoo access tokens last ~1hr; just refresh up front every run.
        self._refresh()

    def _save(self):
        with open(self.token_path, "w") as f:
            json.dump(self.tok, f, indent=4)

    def _refresh(self):
        resp = requests.post(
            TOKEN_URL,
            data={
                "grant_type": "refresh_token",
                "refresh_token": self.tok["refresh_token"],
                "redirect_uri": "oob",
            },
            auth=(self.tok["consumer_key"], self.tok["consumer_secret"]),
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        self.tok["access_token"] = data["access_token"]
        self.tok["refresh_token"] = data.get("refresh_token", self.tok["refresh_token"])
        self.tok["token_type"] = data.get("token_type", "bearer")
        self.tok["token_time"] = time.time()
        self._save()

    def get(self, path, params=None):
        params = dict(params or {})
        params["format"] = "json"
        url = f"{BASE}/{path}"
        headers = {"Authorization": f"Bearer {self.tok['access_token']}"}
        resp = requests.get(url, headers=headers, params=params, timeout=30)
        if resp.status_code == 401:
            self._refresh()
            headers = {"Authorization": f"Bearer {self.tok['access_token']}"}
            resp = requests.get(url, headers=headers, params=params, timeout=30)
        if resp.status_code >= 400:
            # raise_for_status() alone drops the response body, which is
            # where Yahoo actually explains *why* (missing app permission,
            # bad resource path, etc.) -- surface it instead of guessing.
            raise RuntimeError(
                f"{resp.status_code} error for {url}\n"
                f"Response body: {resp.text[:2000]}"
            )
        return resp.json()


def find_all(obj, key):
    """Recursively yield every dict anywhere in obj that contains `key`."""
    if isinstance(obj, dict):
        if key in obj:
            yield obj
        for v in obj.values():
            yield from find_all(v, key)
    elif isinstance(obj, list):
        for v in obj:
            yield from find_all(v, key)


def numbered_items(container):
    """Yahoo repeats elements as {"0": {...}, "1": {...}, "count": N}."""
    if not isinstance(container, dict):
        return
    for k, v in container.items():
        if k != "count":
            yield v


def merge_meta_list(lst):
    """A Yahoo object's flat attributes often arrive as a list of one-key
    dicts (e.g. [{"team_key": ...}, {"name": ...}, ...]) — merge them into
    a single dict. Non-dict / already-complex entries are kept as-is under
    their own key when they are dicts; ignored otherwise."""
    merged = {}
    if isinstance(lst, dict):
        return lst
    for item in lst:
        if isinstance(item, dict):
            merged.update(item)
    return merged

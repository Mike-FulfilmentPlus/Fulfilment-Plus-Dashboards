"""
cc_auth.py  -  Cached CartonCloud OAuth token

Stores the token in .cc_token_cache.json and reuses it until 5 minutes
before expiry, avoiding repeated token requests in the same session.
"""

import json, datetime, os, requests

CACHE_FILE = ".cc_token_cache.json"
BUFFER_SECS = 300   # refresh 5 minutes before actual expiry


def get_cc_token(creds):
    """Return a valid CartonCloud access token, using cache if possible."""
    # Check cache
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE) as f:
                cached = json.load(f)
            expires_at = datetime.datetime.fromisoformat(cached["expires_at"])
            if datetime.datetime.utcnow() < expires_at:
                return cached["access_token"]
        except Exception:
            pass  # cache corrupt or missing — fall through to fresh request

    # Request new token
    r = requests.post(
        "https://api.cartoncloud.com/uaa/oauth/token",
        auth=(creds["CARTONCLOUD_CLIENT_ID"], creds["CARTONCLOUD_CLIENT_SECRET"]),
        headers={"Accept-Version": "1"},
        data={"grant_type": "client_credentials"},
        timeout=30,
    )
    r.raise_for_status()
    data = r.json()
    token = data["access_token"]
    expires_in = int(data.get("expires_in", 3600))

    # Save to cache
    expires_at = datetime.datetime.utcnow() + datetime.timedelta(seconds=expires_in - BUFFER_SECS)
    try:
        with open(CACHE_FILE, "w") as f:
            json.dump({"access_token": token, "expires_at": expires_at.isoformat()}, f)
    except Exception:
        pass  # caching is best-effort

    return token

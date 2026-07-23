#!/usr/bin/env python3
"""Graph の生の会議時刻と、slack_attachment_reader の変換結果を並べて表示する診断。"""
import os
from datetime import datetime, timedelta, timezone

import msal
import requests

GRAPH_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
MSAL_CACHE_FILE = os.path.expanduser("~/.slack_attachment_reader_msal_cache.bin")
DRIVERS = ["giga03@t2.auto", "giga04@t2.auto", "giga05@t2.auto", "giga06@t2.auto"]
JST = timezone(timedelta(hours=9))


def token():
    cache = msal.SerializableTokenCache()
    with open(MSAL_CACHE_FILE) as f:
        cache.deserialize(f.read())
    app = msal.PublicClientApplication(
        GRAPH_CLIENT_ID,
        authority="https://login.microsoftonline.com/organizations",
        token_cache=cache)
    for acc in app.get_accounts():
        r = app.acquire_token_silent(
            ["Calendars.Read", "Calendars.Read.Shared", "OnlineMeetings.Read"], account=acc)
        if r and "access_token" in r:
            return r["access_token"]
    raise SystemExit("silent token取得失敗")


def convert_like_tool(dt_str):
    """apply_teams_meeting と同じ変換 (Zに正規化 → UTCとみなす → JST)"""
    s = dt_str.split(".")[0] + "Z" if "." in dt_str else (dt_str if dt_str.endswith("Z") else dt_str + "Z")
    return datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone(JST)


def main():
    tok = token()
    headers = {"Authorization": f"Bearer {tok}"}
    now = datetime.utcnow()
    start = (now - timedelta(days=10)).isoformat() + "Z"
    end = (now + timedelta(days=15)).isoformat() + "Z"

    for driver in DRIVERS:
        r = requests.get(
            f"https://graph.microsoft.com/v1.0/users/{driver}/calendarView"
            f"?startDateTime={start}&endDateTime={end}"
            f"&$select=subject,start,end,onlineMeeting,isOnlineMeeting&$top=500",
            headers=headers, timeout=60)
        if not r.ok:
            print(f"[{driver}] 取得失敗 {r.status_code}: {r.text[:150]}")
            continue
        events = [e for e in r.json().get("value", []) if e.get("isOnlineMeeting")]
        print(f"\n===== {driver}: オンライン会議 {len(events)} 件 =====")
        for ev in events:
            subj = (ev.get("subject") or "")[:40]
            s_raw, s_tz = ev["start"]["dateTime"], ev["start"]["timeZone"]
            e_raw, e_tz = ev["end"]["dateTime"], ev["end"]["timeZone"]
            s_jst = convert_like_tool(s_raw)
            e_jst = convert_like_tool(e_raw)
            print(f"  ▶ {subj}")
            print(f"     raw start: {s_raw}  (timeZone={s_tz})")
            print(f"     raw end  : {e_raw}  (timeZone={e_tz})")
            print(f"     ツール変換(UTC→JST): {s_jst:%Y-%m-%d %H:%M} 〜 {e_jst:%Y-%m-%d %H:%M}")


if __name__ == "__main__":
    main()

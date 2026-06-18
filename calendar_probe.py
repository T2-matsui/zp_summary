#!/usr/bin/env python3
"""
calendar_probe.py

自分以外のドライバーのカレンダー/予定表にアクセスできるかをテストするスクリプト。

使い方:
    python calendar_probe.py <ドライバーのメールアドレス または UPN>

例:
    python calendar_probe.py driver-a@example.com

何を試すか:
    1. 自分のカレンダーが読めるか(基準)
    2. 指定ドライバーのカレンダーを Calendars.Read.Shared で読めるか
    3. ユーザー検索 (User.Read.All 系)
    4. グループ予定表アクセス (グループID指定不要・自分が所属するグループ一覧)
"""

import json
import os
import sys
from datetime import datetime, timedelta

import msal
import requests


GRAPH_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
MSAL_CACHE_FILE = os.path.expanduser("~/.slack_attachment_reader_msal_cache.bin")
# 試したいスコープ (Group.Read.All は管理者承認必須なので外した最小版)
SCOPES = [
    "Files.Read.All",
    "OnlineMeetings.Read",
    "Calendars.Read",
    "Calendars.Read.Shared",  # ←追加: 他人のカレンダー
    "User.Read",
]


def get_graph_token() -> str:
    cache = msal.SerializableTokenCache()
    if os.path.exists(MSAL_CACHE_FILE):
        with open(MSAL_CACHE_FILE) as f:
            cache.deserialize(f.read())

    app = msal.PublicClientApplication(
        GRAPH_CLIENT_ID,
        authority="https://login.microsoftonline.com/organizations",
        token_cache=cache,
    )

    result = None
    for account in app.get_accounts():
        result = app.acquire_token_silent(SCOPES, account=account)
        if result:
            break

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Device flow開始失敗: {flow}")
        print("\n=== Microsoft認証 (予定表アクセス用追加スコープ要求) ===",
              file=sys.stderr)
        print(flow["message"], file=sys.stderr)
        print("=" * 60, file=sys.stderr)
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(f"トークン取得失敗: {result.get('error_description', result)}")

    if cache.has_state_changed:
        with open(MSAL_CACHE_FILE, "w") as f:
            f.write(cache.serialize())

    return result["access_token"]


def test_self_calendar(token: str):
    """Test 1: 自分のカレンダー読み取り(基準)"""
    print("\n=== Test 1: 自分のカレンダー読み取り(基準確認) ===")
    headers = {"Authorization": f"Bearer {token}"}
    now = datetime.utcnow()
    start = (now - timedelta(days=7)).isoformat() + "Z"
    end = (now + timedelta(days=14)).isoformat() + "Z"
    r = requests.get(
        f"https://graph.microsoft.com/v1.0/me/calendarView"
        f"?startDateTime={start}&endDateTime={end}&$top=5"
        f"&$select=subject,start,end",
        headers=headers, timeout=30,
    )
    print(f"status: {r.status_code}")
    if r.ok:
        events = r.json().get("value", [])
        print(f"  ✓ 取得成功: 直近21日で {len(events)} 件の予定")
        for ev in events[:3]:
            print(f"    - {ev.get('subject', '(no subject)')[:50]}")
    else:
        print(f"  ✗ 失敗: {r.text[:200]}")


def test_other_calendar(token: str, target_upn: str):
    """Test 2: 他人のカレンダー読み取り"""
    print(f"\n=== Test 2: {target_upn} のカレンダー読み取り ===")
    headers = {"Authorization": f"Bearer {token}"}
    now = datetime.utcnow()
    start = (now - timedelta(days=7)).isoformat() + "Z"
    end = (now + timedelta(days=14)).isoformat() + "Z"
    r = requests.get(
        f"https://graph.microsoft.com/v1.0/users/{target_upn}/calendarView"
        f"?startDateTime={start}&endDateTime={end}&$top=10"
        f"&$select=subject,start,end,organizer,onlineMeeting,isOnlineMeeting",
        headers=headers, timeout=30,
    )
    print(f"status: {r.status_code}")
    if r.ok:
        events = r.json().get("value", [])
        online_count = sum(1 for ev in events if ev.get("isOnlineMeeting"))
        print(f"  ✓ 取得成功: {len(events)} 件の予定 (うちオンライン会議 {online_count} 件)")
        for ev in events[:5]:
            online = "[Teams]" if ev.get("isOnlineMeeting") else "[obj]"
            print(f"    {online} {ev.get('subject', '(no subject)')[:50]}")
    elif r.status_code == 403:
        print(f"  ✗ 権限不足 (Calendars.Read.Sharedで取れない = カレンダーが共有されていない)")
        print(f"     body: {r.text[:300]}")
    elif r.status_code == 404:
        print(f"  ✗ ユーザーが見つからない (UPN/メールアドレスを確認)")
    else:
        print(f"  ✗ 失敗: {r.text[:300]}")


def test_my_groups(token: str):
    """Test 3: 自分が所属するMicrosoft 365 グループ一覧"""
    print("\n=== Test 3: 自分が所属する Microsoft 365 グループ ===")
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(
        "https://graph.microsoft.com/v1.0/me/memberOf"
        "?$select=id,displayName,groupTypes,mail",
        headers=headers, timeout=30,
    )
    print(f"status: {r.status_code}")
    if r.ok:
        groups = r.json().get("value", [])
        m365_groups = [g for g in groups if "Unified" in (g.get("groupTypes") or [])]
        print(f"  ✓ 全グループ {len(groups)} 件 / M365グループ {len(m365_groups)} 件")
        for g in m365_groups[:5]:
            print(f"    - {g.get('displayName', '')} (id={g.get('id', '')[:8]}...)")
        if m365_groups:
            print("\n  ↓ 上記グループのいずれかが「配車予定」「ドライバー」等なら、")
            print("     そのIDで calendar/events を引けば共有予定表が取れる可能性あり")
    else:
        print(f"  ✗ 失敗: {r.text[:300]}")


def test_shared_calendars(token: str):
    """Test 4: 自分のOutlookに追加されている共有カレンダー一覧"""
    print("\n=== Test 4: 自分のOutlookに追加されている共有カレンダー ===")
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(
        "https://graph.microsoft.com/v1.0/me/calendars"
        "?$select=id,name,owner,canShare,canEdit",
        headers=headers, timeout=30,
    )
    print(f"status: {r.status_code}")
    if r.ok:
        cals = r.json().get("value", [])
        print(f"  ✓ アクセス可能なカレンダー {len(cals)} 件")
        for c in cals[:10]:
            owner = (c.get("owner") or {}).get("address", "")
            print(f"    - {c.get('name', '')[:30]} (owner: {owner})")
    else:
        print(f"  ✗ 失敗: {r.text[:300]}")


def main():
    print("Microsoft Graph 予定表アクセス プローブ")
    print("=" * 60)

    target_upn = sys.argv[1] if len(sys.argv) > 1 else None

    print("MSAL認証中...")
    token = get_graph_token()
    print("✓ トークン取得完了")

    # 各テストを実行
    test_self_calendar(token)
    test_shared_calendars(token)
    # test_my_groups(token)  # Group.Read.All を要求するので無効化
    if target_upn:
        test_other_calendar(token, target_upn)
    else:
        print("\n=== Test 2: スキップ ===")
        print("  ドライバーのメール/UPNを引数で指定すると、そのカレンダーを試します")
        print("  例: python calendar_probe.py driver-a@example.com")


if __name__ == "__main__":
    main()

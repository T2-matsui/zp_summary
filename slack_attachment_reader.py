#!/usr/bin/env python3
"""Slackの運行記録チャンネルから Teams 会議URL を含む親メッセージを読み取り、
Microsoft Graph (会議情報 / ドライバー共有予定表) と突き合わせて運行記録を
legs.json として出力する社内ツール。詳細は README.md 参照。

    python slack_attachment_reader.py --config config.json

依存: slack_sdk requests python-dotenv msal / 環境変数 SLACK_TOKEN (or SLACK_BOT_TOKEN)
"""

import argparse
import json
import os
import re
import sys
import traceback
import unicodedata
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import requests
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

JST = timezone(timedelta(hours=9))
# 対象日(content_date)に対して投稿を拾う窓。運行記録は運行日の前後どちらにも投稿されるため
# 「投稿日 ±POST_WINDOW_DAYS」で粗く絞り、実日付判定は Teams会議側で行う
POST_WINDOW_DAYS = 14
TEAMS_URL_RE = re.compile(
    r'https://teams\.microsoft\.com/l/meetup-join/[^\s<>"\'|]+', re.IGNORECASE)

# Microsoft Graph CLI Tools の公式アプリID (Azure ADへの登録不要)
GRAPH_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
MSAL_CACHE_FILE = os.path.expanduser("~/.slack_attachment_reader_msal_cache.bin")

# Slackスレッド本文 → tracking メタ抽出 (半角/全角の括弧・コロン両対応)
_PAREN_O, _PAREN_C, _COLON = r'[\(（]', r'[\)）]', r'[:：]'
TRACKING_DATE_TITLE_RE = re.compile(r'(\d{1,2})月(\d{1,2})日')
TRACKNAME_RE = re.compile(r'【([^】]+)】')
TRACKNUM_RE = re.compile(r'(\d+)')
CUSTOMER_TAIL_ARROW_RE = re.compile(r'→')
# 正規の号車表記は 【giga05】(半角小文字 giga + 2桁) と 【重要運行】のみ。
# 大文字・全角・桁不足 (GIGA05 / ｇｉｇａ０５ / giga5) は giga05 に正規化した上で通知に報告する。
# 【GIGA05→06】(号車変更) のように号車が一意に定まらない表記は補正せず警告だけに留める。
TRACK_CANONICAL_RE = re.compile(r'^(?:giga\d{2}|重要運行)$')
# 正規化できる号車表記 (NFKC 後に giga+1〜2桁とみなせるもの)
GIGA_NORMALIZABLE_RE = re.compile(r'^giga\s*(\d{1,2})$', re.IGNORECASE)
# 号車変更表記 (giga05→06 / GIGA05->giga06 等)。変更後 = 最後の号車を採用する
GIGA_CHANGE_RE = re.compile(
    r'^giga\s*\d{1,2}(?:\s*(?:→|⇒|➡|=>|->)\s*(?:giga\s*)?\d{1,2})+$', re.IGNORECASE)
DRIVER_LINE_RE = re.compile(
    rf'ドライバー\s*{_PAREN_O}幹線{_PAREN_C}\s*{_COLON}\s*<@(U[A-Z0-9]+)>')
OPERATOR_LINE_RE = re.compile(
    rf'オペレータ\s*{_PAREN_O}幹線{_PAREN_C}\s*{_COLON}\s*<@(U[A-Z0-9]+)>')
SW_VER_RE = re.compile(
    rf'SW\s*Ver\.?\s*{_COLON}\s*WIP\s*{_COLON}\s*(v\S+)', re.IGNORECASE)
ROUTE_LINE_RE = re.compile(rf'自動運転区間\s*{_COLON}\s*([^\n]+)')


def log(msg):
    print(msg, file=sys.stderr)


def _iso_z(s: str) -> str:
    """dateTime文字列の末尾マイクロ秒を落とし 'Z' 終端に揃える。"""
    if not s:
        return s
    if "." in s:
        return s.split(".")[0] + "Z"
    return s if s.endswith("Z") else s + "Z"


def extract_teams_urls(msg: dict) -> list[str]:
    urls: set[str] = set()
    urls.update(TEAMS_URL_RE.findall(msg.get("text", "") or ""))
    for att in msg.get("attachments", []) or []:
        for key in ("from_url", "original_url", "title_link", "fallback", "text"):
            urls.update(TEAMS_URL_RE.findall(att.get(key) or ""))
    return list(urls)


# ---------- Slack ----------

def get_token() -> str:
    token = os.environ.get("SLACK_TOKEN") or os.environ.get("SLACK_BOT_TOKEN")
    if not token:
        sys.exit("環境変数 SLACK_TOKEN (xoxp-...) が設定されていません。")
    return token


def resolve_channel_id(client: WebClient, channel: str) -> str:
    """チャンネル名を ID に変換。C/G/D で始まる ID はそのまま返す。"""
    channel = channel.lstrip("#")
    if channel[:1] in {"C", "G", "D"} and channel[1:].isalnum() and channel.isupper():
        return channel
    # public+private で検索。groups:read が無ければ public のみで再試行
    for types in ("public_channel,private_channel", "public_channel"):
        cursor = None
        try:
            while True:
                resp = client.conversations_list(limit=200, cursor=cursor, types=types)
                for ch in resp["channels"]:
                    if ch["name"] == channel:
                        return ch["id"]
                cursor = resp.get("response_metadata", {}).get("next_cursor")
                if not cursor:
                    break
            break
        except SlackApiError as e:
            if e.response.get("error") == "missing_scope" and "private" in types:
                log("[情報] groups:read 未付与のため publicチャンネルのみで再検索します")
                continue
            raise
    sys.exit(f"チャンネル '{channel}' が見つかりません "
             "(private の場合は groups:read / groups:history が必要)")


def fetch_messages(client: WebClient, channel_id: str, limit: int,
                   since_ts: str | None = None) -> list[dict]:
    kwargs = {"channel": channel_id, "limit": limit}
    if since_ts:
        kwargs["oldest"] = since_ts
    return list(client.conversations_history(**kwargs)["messages"])


def collect_teams_posts(messages: list[dict],
                        target_date: "date | None" = None) -> list[dict]:
    """Teams URL を含む親メッセージを抽出。target_date 指定時は投稿日 ±POST_WINDOW_DAYS で
    粗く絞る (実日付判定は Teams会議側で行う)。"""
    results: list[dict] = []
    seen_ts: set[str] = set()
    for msg in messages:
        parent = msg.get("_parent_msg") or msg
        ts = parent.get("ts", "")
        if not ts or ts in seen_ts or not extract_teams_urls(parent):
            continue
        if target_date:
            try:
                posted = datetime.fromtimestamp(float(ts), tz=JST).date()
                if abs((posted - target_date).days) > POST_WINDOW_DAYS:
                    continue
            except (ValueError, TypeError):
                continue
        seen_ts.add(ts)
        snippet = (parent.get("text") or "").splitlines()[0][:40] \
            if parent.get("text") else ts
        results.append({"_msg": parent, "_posted_ts": ts,
                        "name": f"[teams投稿] {snippet}"})
    return results


def get_message_permalink(client: WebClient, channel_id: str, ts: str) -> str:
    try:
        return client.chat_getPermalink(channel=channel_id, message_ts=ts).get("permalink", "")
    except SlackApiError:
        return ""


def resolve_user_name(user_id: str, client: WebClient, cache: dict) -> str:
    """ユーザーID→表示名。display_name → real_name → name の優先順。"""
    if user_id in cache:
        return cache[user_id]
    try:
        user = client.users_info(user=user_id)["user"]
        prof = user.get("profile", {}) or {}
        name = (prof.get("display_name") or "").strip() \
            or (prof.get("real_name") or "").strip() \
            or (user.get("real_name") or "").strip() \
            or user.get("name") or f"<@{user_id}>"
        cache[user_id] = name
        return name
    except SlackApiError as e:
        err = e.response.get("error", "unknown")
        log(f"[警告] ユーザー解決失敗 ({user_id}): {err}"
            + (" -- users:read スコープが必要" if err == "missing_scope" else ""))
        cache[user_id] = f"<@{user_id}>"
        return cache[user_id]


def normalize_slack_text(text: str) -> str:
    """Slack装飾 (リンク表記・絵文字・太字など) を除去。"""
    if not text:
        return ""
    text = re.sub(r'<([^|>\s]+)\|([^>]+)>', r'\2', text)   # <url|display> → display
    text = re.sub(r'<(https?://[^>]+)>', r'\1', text)       # <url> → url
    text = re.sub(r'(?<![a-zA-Z0-9_.]):[a-zA-Z0-9_+\-]+:(?!//)', '', text)  # :emoji:
    text = re.sub(r'[\*_~]+', '', text)                     # *bold* _italic_ ~strike~
    return text


def normalize_trackname(track: str) -> str:
    """号車表記を正規形 (半角小文字 giga + 2桁) に揃える。

    'GIGA05' 'ｇｉｇａ０５' 'giga5' → 'giga05' / '重要運行' → '重要運行'
    号車変更表記 'giga05→06' は **変更後 (最後) の号車** 'giga06' を採用する。
    'giga100' のようにどの号車か決められない表記は NFKC 正規化だけして返す。
    出力 (build_legs_record) と重複判定 (legs_dedup_key) の両方がここを通るので、
    表記ゆれが別レコードとして二重登録されない。merge_legs.py 側にも同じ関数がある。
    """
    s = unicodedata.normalize("NFKC", track or "").strip()
    m = GIGA_NORMALIZABLE_RE.match(s)
    if m:
        return f"giga{m.group(1).zfill(2)}"
    if GIGA_CHANGE_RE.match(s):
        # 「05→06」= 06 に変更された、の意。変更後の号車で運行している
        return f"giga{TRACKNUM_RE.findall(s)[-1].zfill(2)}"
    return s


def parse_track_filter(values) -> set:
    """track_filter 設定を正規化済み号車の集合にする。空 (未指定) なら絞り込み無効。"""
    if not values:
        return set()
    items = values if isinstance(values, list) else [values]
    return {normalize_trackname(v) for v in items if str(v).strip()}


def track_allowed(track: str, allow: set) -> bool:
    """号車が取り込み対象か。allow が空なら常に True (絞り込み無効)。

    号車を読めなかった投稿 (空) ・【重要運行】 (giga番号が未確定) ・【giga05→06】 の
    ような一意に定まらない表記は、ここでは落とさず True を返す。号車が確定していない
    ものを無言で消すと気付けないため。

    run() はこれを **2 回** 呼ぶ: ①Slack本文の号車 (Teams会議を引く前。対象外と分かって
    いる投稿で Graph を叩かないため) ②Teams会議件名で確定した号車 (①を通り抜けた
    【重要運行】・号車なしの投稿を取りこぼさないため)。①だけだと対象外号車が
    legs.json に混ざる。
    """
    if not allow:
        return True
    norm = normalize_trackname(track)
    if norm == "重要運行" or not TRACK_CANONICAL_RE.match(norm):
        return True
    return norm in allow


def set_trackname(meta: dict, raw: str, source: str) -> None:
    """号車表記を正規化して meta に格納し、正規形と違っていれば警告を積む。

    Slack本文と Teams会議件名の両方がここを通る (Teams件名由来の表記も検証される)。
    """
    track = (raw or "").strip()
    if not track:
        return
    norm = normalize_trackname(track)
    meta["Trackname"] = norm
    nm = TRACKNUM_RE.search(norm)
    if nm:
        meta["Track-num"] = nm.group(1)
    if not TRACK_CANONICAL_RE.match(norm):
        meta.setdefault("_warn", []).append(
            f"{source}の号車表記が非正規で自動補正できません: 【{track}】 (正規: 【giga05】形式)")
    elif GIGA_CHANGE_RE.match(unicodedata.normalize("NFKC", track).strip()):
        # 号車変更は「どちらの号車で走ったか」の判断が入るため、補正内容を明示する
        meta.setdefault("_warn", []).append(
            f"{source}の号車変更表記から変更後の号車を採用しました: 【{track}】 → 【{norm}】")
    elif norm != track:
        meta.setdefault("_warn", []).append(
            f"{source}の号車表記を補正しました: 【{track}】 → 【{norm}】")


def extract_tracking_metadata(msg: dict, client: WebClient, user_cache: dict,
                              default_year: int) -> dict:
    """親メッセージ本文から tracking フィールドを抽出 (返信なら _parent_msg を見る)。"""
    source_msg = msg.get("_parent_msg") or msg
    text = normalize_slack_text(source_msg.get("text", "") or "")
    result: dict = {}

    # タイトル行 = 【...】を含む最初の行
    first_line = next((ln for ln in text.split("\n") if "【" in ln and "】" in ln),
                      text.split("\n", 1)[0])
    result["_title_line"] = first_line.strip()

    brackets = TRACKNAME_RE.findall(first_line)
    if brackets:
        set_trackname(result, brackets[0], "Slack本文")
        # 【giga05】【giga06】のような2台併記は先頭しか取り込めないため取りこぼしを警告する
        # (【重要】等の号車番号を含まない併記は対象外)
        others = [b.strip() for b in brackets[1:]
                  if _track_nums(b) - _track_nums(brackets[0])]
        if others:
            result.setdefault("_warn", []).append(
                f"Slack本文に号車が複数あります: 【{brackets[0].strip()}】 のみ取り込み、"
                + "".join(f"【{o}】" for o in others) + " は取りこぼしています")

    # Customer: 】の後 ～ 末尾の方向(X→Y)・装飾を除いた部分
    m = re.search(r'】([^\n<※]+)', first_line)
    if m:
        tokens = m.group(1).strip().split()
        while tokens and CUSTOMER_TAIL_ARROW_RE.search(tokens[-1]):
            tokens.pop()
        while tokens and re.fullmatch(r'[※#\-=●○◯◎★☆]+', tokens[-1]):
            tokens.pop()
        result["Customer"] = " ".join(tokens)

    m = TRACKING_DATE_TITLE_RE.search(first_line)
    if m:
        try:
            result["_title_date"] = date(default_year, int(m.group(1)), int(m.group(2)))
        except ValueError:
            pass

    m = DRIVER_LINE_RE.search(text)
    if m:
        result["Driver-name"] = resolve_user_name(m.group(1), client, user_cache)

    m = OPERATOR_LINE_RE.search(text)
    result["Operator"] = resolve_user_name(m.group(1), client, user_cache) if m else "oneman"

    m = SW_VER_RE.search(text)
    if m:
        result["SW-ver"] = m.group(1)

    m = ROUTE_LINE_RE.search(text)
    if m:
        result["Route"] = m.group(1).strip()

    return result


# ---------- Microsoft Graph / Teams ----------

def get_graph_token(extra_scopes: list[str] | None = None) -> str:
    """Graphトークン取得。キャッシュあれば自動更新、無ければデバイスコード認証。"""
    import msal
    cache = msal.SerializableTokenCache()
    if os.path.exists(MSAL_CACHE_FILE):
        with open(MSAL_CACHE_FILE) as f:
            cache.deserialize(f.read())
    app = msal.PublicClientApplication(
        GRAPH_CLIENT_ID,
        authority="https://login.microsoftonline.com/organizations",
        token_cache=cache)
    scopes = list(extra_scopes or ["OnlineMeetings.Read", "Calendars.Read"])

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
    if not result:
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            raise RuntimeError(f"Device flow開始失敗: {flow}")
        log("\n=== Microsoft認証が必要です ===")
        log(flow["message"])
        log(f"要求スコープ: {scopes}\n")
        result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"トークン取得失敗: {result.get('error_description', result)}")
    if cache.has_state_changed:
        with open(MSAL_CACHE_FILE, "w") as f:
            f.write(cache.serialize())
    return result["access_token"]


def fetch_track_calendars(driver_emails: list[str], graph_token: str,
                          target_date: "date | None", days_window: int = 14) -> dict:
    """各ドライバーの±N日分オンライン会議を JoinUrl キーの dict で返す。"""
    headers = {"Authorization": f"Bearer {graph_token}"}
    target_date = target_date or datetime.now().date()
    base = datetime.combine(target_date, datetime.min.time())
    start = (base - timedelta(days=days_window)).isoformat() + "Z"
    end = (base + timedelta(days=days_window)).isoformat() + "Z"

    join_url_map: dict = {}
    for driver in driver_emails:
        try:
            r = requests.get(
                f"https://graph.microsoft.com/v1.0/users/{driver}/calendarView"
                f"?startDateTime={start}&endDateTime={end}"
                f"&$select=subject,start,end,onlineMeeting,isOnlineMeeting,organizer&$top=500",
                headers=headers, timeout=60)
            if not r.ok:
                log(f"[track-cal] {driver} 取得失敗 status={r.status_code}: {r.text[:200]}")
                continue
            events = r.json().get("value", [])
            n = 0
            for ev in events:
                join = (ev.get("onlineMeeting") or {}).get("joinUrl", "")
                if ev.get("isOnlineMeeting") and join and join not in join_url_map:
                    join_url_map[join] = ev
                    n += 1
            log(f"[track-cal] {driver}: 全{len(events)}件 / オンライン会議 {n}件を取得")
        except requests.RequestException as e:
            log(f"[track-cal] {driver} 例外: {type(e).__name__}: {e}")
    return join_url_map


def get_teams_meeting(url: str, graph_token: str) -> dict | None:
    """OnlineMeetings API (自分主催) → CalendarView API (自分の予定) の順で会議情報を取得。"""
    headers = {"Authorization": f"Bearer {graph_token}"}
    try:
        r = requests.get(
            "https://graph.microsoft.com/v1.0/me/onlineMeetings"
            f"?$filter=JoinWebUrl eq '{url.replace(chr(39), chr(39) * 2)}'",
            headers=headers, timeout=30)
        if r.ok:
            items = r.json().get("value", [])
            if items:
                return items[0]
        elif r.status_code != 403:
            log(f"    [teams API失敗] status={r.status_code} {r.text[:200]}")
    except requests.RequestException as e:
        log(f"    [teams API例外] {type(e).__name__}: {e}")

    # CalendarView フォールバック
    now = datetime.utcnow()
    try:
        r = requests.get(
            "https://graph.microsoft.com/v1.0/me/calendarView"
            f"?startDateTime={(now - timedelta(days=90)).isoformat()}Z"
            f"&endDateTime={(now + timedelta(days=90)).isoformat()}Z"
            "&$select=subject,start,end,onlineMeeting,isOnlineMeeting&$top=500",
            headers=headers, timeout=60)
        if not r.ok:
            log(f"    [calendar API失敗] status={r.status_code} {r.text[:200]}")
            return None
        for ev in r.json().get("value", []):
            join = (ev.get("onlineMeeting") or {}).get("joinUrl", "")
            if ev.get("isOnlineMeeting") and join and (join == url or url in join or join in url):
                log(f"    [calendar] マッチ: {ev.get('subject', '')[:60]}")
                return {"subject": ev.get("subject", ""),
                        "startDateTime": _iso_z(ev["start"]["dateTime"]),
                        "endDateTime": _iso_z(ev["end"]["dateTime"]),
                        "_source": "calendarView"}
        log(f"    [calendar] joinUrl一致せず ({len(r.json().get('value', []))}件中): {url[:80]}")
        return None
    except requests.RequestException as e:
        log(f"    [calendar API例外] {type(e).__name__}: {e}")
        return None


def parse_teams_subject(subject: str) -> dict:
    """会議件名 (例: 【giga■■】○○様 実証 往路 関東→関西) から Trackname/Customer/Route を抽出。"""
    result: dict = {}
    if not subject:
        return result
    m = TRACKNAME_RE.search(subject)
    if m:
        track = m.group(1).strip()
        result["Trackname"] = track
        nm = TRACKNUM_RE.search(track)
        if nm:
            result["Track-num"] = nm.group(1)
    m = re.search(r'】(.+)$', subject)
    if m:
        tokens = m.group(1).strip().split()
        route_tokens: list = []
        while tokens and CUSTOMER_TAIL_ARROW_RE.search(tokens[-1]):
            route_tokens.insert(0, tokens.pop())
        while tokens and re.fullmatch(r'[※#\-=●○◯◎★☆]+', tokens[-1]):
            tokens.pop()
        result["Customer"] = " ".join(tokens)
        result["Route"] = " ".join(route_tokens)
    return result


def _track_nums(brackets: str) -> set:
    """【】内文字列から号車番号を0埋め2桁の集合で返す ('GIGA05→06' → {'05','06'})。"""
    return {n.zfill(2) for n in TRACKNUM_RE.findall(brackets or "")}


def apply_teams_meeting(meta: dict, meeting: dict, debug: bool = False) -> None:
    """Teams会議情報を meta に反映 (Trackname/Track-num/Customer と開始/終了時刻)。
    Route は Slack本文の `自動運転区間：` を優先するため上書きしない。"""
    subject = meeting.get("subject", "") or ""
    tmeta = parse_teams_subject(subject)
    # 他号車の Teams会議室を流用した投稿だと Trackname が Teams側の号車に上書きされてしまう。
    # 自動判定はせず、Slack本文と号車が食い違うことだけを警告する (上書き前に比較)
    s_nums = _track_nums(meta.get("Trackname", ""))
    t_nums = _track_nums("".join(TRACKNAME_RE.findall(subject)))
    if s_nums and t_nums and not (s_nums & t_nums):
        meta.setdefault("_warn", []).append(
            f"Slack本文と Teams会議件名で号車が不一致です: "
            f"Slack=【{meta.get('Trackname', '')}】 / Teams件名=\"{subject}\" "
            f"→ Teams側を採用")
    if tmeta.get("Trackname"):
        # Teams件名由来の表記も正規化・検証する (Track-num も併せて更新される)
        set_trackname(meta, tmeta["Trackname"], "Teams会議件名")
    elif tmeta.get("Track-num"):
        meta["Track-num"] = tmeta["Track-num"]
    if tmeta.get("Customer"):
        meta["Customer"] = tmeta["Customer"]
    start_str, end_str = meeting.get("startDateTime", ""), meeting.get("endDateTime", "")
    if not (start_str and end_str):
        if debug:
            log(f"    [teams診断] 開始/終了時刻が無い: {list(meeting.keys())}")
        return
    try:
        s_dt = datetime.fromisoformat(start_str.replace("Z", "+00:00")).astimezone(JST)
        e_dt = datetime.fromisoformat(end_str.replace("Z", "+00:00")).astimezone(JST)
    except ValueError as e:
        log(f"    [teams] 時刻パース失敗: {e}")
        return
    meta["_teams_start_dt"], meta["_teams_end_dt"] = s_dt, e_dt
    meta["_title_date"] = s_dt.date()  # 日付は Teams会議の開始日を採用
    date_str = s_dt.strftime("%Y/%m/%d")
    meta["_teams_time_range"] = (f"{date_str}T{s_dt.strftime('%H:%M')}+9:00/"
                                 f"{date_str}T{e_dt.strftime('%H:%M')}+9:00")
    log(f"    [teams時刻] {tmeta.get('Trackname') or '?'}: "
        f"{s_dt.strftime('%Y-%m-%d %H:%M')}〜{e_dt.strftime('%H:%M')} JST "
        f"→ {meta['_teams_time_range']}")


# ---------- legs.json レコード ----------

def build_legs_record(meta: dict, url: str) -> list:
    """legs.json レコード [Trackname, "Track-num|YY/MM/DD", "開始ISO/終了ISO",
    {SW-version, selfdrive_section, loaded_luggage, url}] を作る (跨日OK)。"""
    track = meta.get("Trackname", "")
    track_num = meta.get("Track-num", "")
    title_date: "date | None" = meta.get("_title_date")
    date_str = title_date.strftime("%Y/%m/%d") if title_date else ""

    teams_start, teams_end = meta.get("_teams_start_dt"), meta.get("_teams_end_dt")
    time_range = ""
    if teams_start and teams_end:
        time_range = (f"{teams_start.isoformat(timespec='milliseconds')}/"
                      f"{teams_end.isoformat(timespec='milliseconds')}")

    def _giga(v):  # 出力上の "GIGA" は小文字 "giga" に統一
        return v.replace("GIGA", "giga") if isinstance(v, str) else v

    def _strip_giga_tag(v):  # loaded_luggage 先頭の 【gigaXX】 を除去
        if not isinstance(v, str):
            return v
        return re.sub(r"^\s*【\s*giga\d+\s*】\s*", "", v, flags=re.IGNORECASE)

    track = normalize_trackname(track)
    # 「重要運行」は loaded_luggage の【gigaXX】から番号を拾って Trackname=gigaXX に補正する。
    # 「重要運行」であること自体はレコードに残さない (下流 zero-plotter が Trackname を
    # 号車ID としてそのまま使っており、giga03 以外の値にすると Druid のデータソース名・
    # 号車フィルタ・動画ディレクトリ検索が一致しなくなるため)
    if track == "重要運行":
        customer = unicodedata.normalize("NFKC", meta.get("Customer", "") or "")
        m = re.search(r"giga\s*(\d{1,2})", customer, re.IGNORECASE)
        if m:
            track = normalize_trackname(f"giga{m.group(1)}")
            track_num = m.group(1).zfill(2)

    return [
        track,
        f"{track_num}|{date_str}" if date_str else track_num,
        time_range,
        {
            "SW-version": meta.get("SW-ver", ""),
            "selfdrive_section": meta.get("Route", ""),
            "loaded_luggage": _strip_giga_tag(_giga(meta.get("Customer", ""))),
            "url": url,
        },
    ]


# 空欄扱いするプレースホルダ ("---" "未定" "TBD" 等)
_PLACEHOLDERS = {
    "", "-", "--", "---", "----", "−", "—", "ー", "未定", "未取得", "未設定",
    "なし", "無し", "tbd", "TBD", "n/a", "N/A", "na", "NA", "?", "？", "不明",
}


def _is_placeholder(value) -> bool:
    if value is None:
        return True
    s = str(value).strip()
    return s in _PLACEHOLDERS or (bool(s) and all(ch in "-−—ー" for ch in s))


def is_legs_record_complete(rec) -> tuple[bool, list]:
    """全フィールドが埋まっているか。戻り値: (完全か, 空フィールド名リスト)。"""
    if not isinstance(rec, list) or len(rec) < 4:
        return False, ["record_format"]
    missing: list = []
    if _is_placeholder(rec[0]):
        missing.append("Trackname")
    if not rec[1] or "|" not in (rec[1] or ""):
        missing.append("Date")
    if _is_placeholder(rec[2]):
        missing.append("Time-range")
    meta = rec[3] if isinstance(rec[3], dict) else {}
    for key in ("SW-version", "selfdrive_section", "loaded_luggage", "url"):
        if _is_placeholder(meta.get(key)):
            missing.append(key)
    return len(missing) == 0, missing


IMPORTANT_TAG_RE = re.compile(r'(?:^重要運行[_＿]|[_＿]重要運行$)')


def strip_important_tag(value: str) -> str:
    """過去の legs.json に残る「重要運行」マーカーを取り除く (重複判定でのみ使う)。

    旧形式 '重要運行_giga03' (〜2026-07-06) や '2026/06/22_重要運行' のレコードが
    本番に残っているため、マーカーの有無だけで同じ運行が二重登録されないよう、
    dedup キーの比較時に限って外す。出力する表記は変えない。
    merge_legs.py 側にも同じ関数がある。変更する場合は両方を揃えること。
    """
    return IMPORTANT_TAG_RE.sub("", value or "").strip()


def legs_dedup_key(rec) -> tuple:
    """重複判定キー (Trackname, 日付, 往路/復路)。配列/dict 両形式に対応。

    Trackname は normalize_trackname で正規化して比較するため、既存レコードが
    'GIGA05' 'Giga05' のような表記ゆれでも同一運行として重複判定される。
    merge_legs.legs_dedup_key と同一ロジック。変更する場合は両方を揃えること。
    """
    if isinstance(rec, list) and len(rec) >= 4:
        track = rec[0] or ""
        date_part = rec[1] or ""
        d = rec[3] if isinstance(rec[3], dict) else {}
        liggage = d.get("loaded_luggage") or d.get("loaded liggage", "") or ""
    elif isinstance(rec, dict):
        track = rec.get("Trackname", "") or ""
        date_part = rec.get("Track-num|YY/MM/DD", "") or ""
        liggage = rec.get("loaded_luggage") or rec.get("loaded liggage", "") or ""
    else:
        return ("", "", "")
    date_str = date_part.split("|", 1)[1] if "|" in date_part else ""
    direction = next((w for w in ("往路", "復路") if w in liggage), "")
    # 「重要運行」マーカーは比較前に外す (旧形式のレコードとの二重登録を防ぐ)
    return (normalize_trackname(strip_important_tag(track)),
            strip_important_tag(date_str), direction)


# ---------- I/O ----------

class DataFileError(RuntimeError):
    """既存データファイルが壊れていて安全に続行できない (Slack へ通知して中断する)。"""


def read_json_list(path: str) -> list:
    """BOM対応でJSON配列を読む。存在しない/空の場合は []。

    壊れている場合は [] を返して続行せず DataFileError で中断する。空配列で続行すると
    既存レコードを取りこぼしたまま legs.json を上書きしてしまうため。
    """
    if not os.path.exists(path):
        return []
    try:
        with open(path, "rb") as f:
            raw = f.read()
        if raw.startswith(b'\xff\xfe'):
            content = raw.decode("utf-16-le", errors="replace")
        elif raw.startswith(b'\xfe\xff'):
            content = raw.decode("utf-16-be", errors="replace")
        elif raw.startswith(b'\xef\xbb\xbf'):
            content = raw[3:].decode("utf-8", errors="replace")
        else:
            content = raw.decode("utf-8", errors="replace")
        content = content.lstrip("﻿").strip()
        if not content:
            return []
        data = json.loads(content)
    except OSError as e:
        raise DataFileError(f"{path} を読み込めません: {e}") from e
    except json.JSONDecodeError as e:
        raise DataFileError(f"{path} が JSON として壊れています: {e}") from e
    if not isinstance(data, list):
        raise DataFileError(
            f"{path} が JSON 配列ではありません (type={type(data).__name__})")
    return data


def load_logs_json(path: str) -> tuple[list, "date | None"]:
    """logs.json を読み、(entries, 最新投稿日) を返す。"""
    entries = read_json_list(path)
    max_date: "date | None" = None
    for e in entries:
        if not isinstance(e, dict) or not e.get("posted_at"):
            continue
        try:
            d = datetime.fromisoformat(e["posted_at"].replace("Z", "+00:00")).date()
            if max_date is None or d > max_date:
                max_date = d
        except ValueError:
            continue
    return entries, max_date


def atomic_write_json(path: str, data) -> None:
    """一時ファイルに書いてから os.replace で差し替え (書き込み中クラッシュでも元は無傷)。"""
    p = Path(path)
    tmp = p.with_name(p.name + ".tmp")
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            f.write(json.dumps(data, ensure_ascii=False, indent=2))
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, p)
    except Exception:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass
        raise


def format_mentions(values) -> str:
    """❌ failed 通知に付けるメンションを Slack 記法へ変換する。

    'U01ABCDEF' (メンバーID) → '<@U01ABCDEF>' / 'S01ABCDEF' (ユーザーグループID) →
    '<!subteam^S01ABCDEF>' / 'here' 'channel' → '<!here>' '<!channel>'。
    '<@U01ABCDEF>' のように既に記法で書かれていればそのまま使う。

    表示名 (@田中) は Webhook 側で ID に解決されず、ただの文字列として送られて
    通知が飛ばない。黙って無視すると「メンションしたのに気付けない」ので警告を出す。
    """
    if not values:
        return ""
    items = values if isinstance(values, list) else [values]
    out: list = []
    for v in items:
        s = str(v).strip().lstrip("@")
        if not s:
            continue
        if s.startswith("<"):
            out.append(s)
        elif s.lower() in ("here", "channel"):
            out.append(f"<!{s.lower()}>")
        elif re.fullmatch(r'S[A-Z0-9]{6,}', s):
            out.append(f"<!subteam^{s}>")
        elif re.fullmatch(r'[UW][A-Z0-9]{6,}', s):
            out.append(f"<@{s}>")
        else:
            log(f"[警告] notify_mentions の値 '{v}' はメンバーID/グループIDではないため無視します "
                f"(Slackのプロフィール → 「メンバーIDをコピー」の U から始まる文字列を指定)")
    return " ".join(out)


def send_slack_notification(webhook_url: str, text: str) -> None:
    if not webhook_url:
        # 通知先未設定を黙って無視すると「通知が来ない」ことに気付けないため必ず残す
        log("[警告] 通知先 webhook が未設定のため Slack 通知を送れませんでした")
        return
    try:
        r = requests.post(webhook_url, json={"text": text}, timeout=15)
        if not r.ok:
            log(f"[警告] Slack通知失敗 ({r.status_code}): {r.text[:200]}")
    except requests.RequestException as e:
        log(f"[警告] Slack通知送信エラー: {e}")


# ---------- CLI ----------

def parse_since(s: str) -> str:
    return str(datetime.strptime(s, "%Y-%m-%d").timestamp())


def resolve_relative_date(value: str | None) -> "date | None":
    """YYYY-MM-DD / today / yesterday / tomorrow / N_days_ago / N_days_later を date に。"""
    if not value:
        return None
    today = datetime.now().date()
    fixed = {"today": today, "yesterday": today - timedelta(days=1),
             "tomorrow": today + timedelta(days=1)}
    if value in fixed:
        return fixed[value]
    m = re.match(r'(\d+)_days?_ago$', value)
    if m:
        return today - timedelta(days=int(m.group(1)))
    m = re.match(r'(\d+)_days?_later$', value)
    if m:
        return today + timedelta(days=int(m.group(1)))
    return datetime.strptime(value, "%Y-%m-%d").date()


def build_notification(legs_records: list, target_date: "date | None",
                       legs_new_count: int, legs_skipped_count: int,
                       format_warns: list = (), errors: list = (),
                       filtered_count: int = 0, existing_dropped: int = 0,
                       mentions: str = "") -> str:
    """完了通知テキスト (✅ success / ❌ failed / 🔁 skipped) を組み立てる。
    legs_skipped_count は URL重複・dedupキー重複の両方を合算した件数。
    format_warns は人が確認すべき事象 [(警告文, URL), ...] (号車表記の補正・非正規表記・
    日付が読めずスキップした投稿など)。ステータスは変えず ⚠️ 要確認として列挙する。
    errors は処理中に発生した異常の一覧。1 件でもあれば ✅ success は出さない。
    filtered_count は track_filter で対象外号車として除外した投稿数 (件数のみ報告する。
    毎回同じ号車が並んで ⚠️ 要確認 が埋まるのを避けるため明細は出さない)。
    existing_dropped は append 時に既存 legs.json から取り除いた対象外号車のレコード数
    (件数だけだと消えたことに気付けないので、0 件でなければ必ず行を出す)。
    mentions は ❌ failed のときだけ先頭に付けるメンション (format_mentions の戻り値)。
    ✅ success / 🔁 skipped では付けない (毎日メンションが飛ぶと見なくなるため)。"""
    incomplete: set = set()
    for i, rec in enumerate(legs_records):
        ok, _ = is_legs_record_complete(rec)
        track = rec[0] if isinstance(rec, list) and rec else ""
        if not ok or track == "重要運行":  # giga番号未確定も failed 扱い
            incomplete.add(i)

    if errors:
        status = f"❌ failed: 処理中にエラーが発生しました ({len(errors)} 件)"
    elif legs_new_count == 0 and legs_skipped_count > 0:
        # URL重複でレコード化前に落ちた分も含むため legs_records は空になり得る
        status = f"🔁 skipped: 全て既存レコードと重複のため追加なし ({legs_skipped_count} 件スキップ)"
    elif incomplete:
        status = (f"❌ failed: 運行記録に不完全なレコードがあります "
                  f"({len(incomplete)}/{len(legs_records)} 件)")
    elif not legs_records:
        status = "✅ success: (今回追加されたレコードはありません)"
    elif legs_skipped_count > 0:
        status = f"✅ success: {legs_new_count} 件追加 (重複スキップ {legs_skipped_count} 件)"
    else:
        status = "✅ success: 運行記録の読み込みが完了しました"

    lines = [f"{mentions} 対応をお願いします", status] if (mentions and status.startswith("❌")) \
        else [status]
    if target_date:
        lines.append(f"対象日: {target_date.isoformat()}")
    if filtered_count:
        lines.append(f"対象外号車のためスキップ: {filtered_count} 件 (track_filter)")
    if existing_dropped:
        lines.append(f"対象外号車のため既存 legs.json から除外: {existing_dropped} 件 (track_filter)")
    MAX = 30
    for i, rec in enumerate(legs_records[:MAX]):
        track = rec[0] or "(Trackname不明)"
        meta = rec[3] if isinstance(rec[3], dict) else {}
        marker = ""
        if i in incomplete:
            _, missing = is_legs_record_complete(rec)
            parts = (["Trackname未確定(giga番号取得失敗)"] if track == "重要運行" else []) \
                + ([f"未取得: {', '.join(missing)}"] if missing else [])
            if parts:
                marker = " ⚠️ " + " / ".join(parts)
        lines += ["", f"{track}{marker}",
                  f'運行名: "{meta.get("loaded_luggage", "")}"',
                  f'自動運転区間: "{meta.get("selfdrive_section", "")}"']
        if meta.get("url"):
            lines.append(f'URL: <{meta["url"]}|スレッドを開く>')
    if len(legs_records) > MAX:
        lines += ["", f"…他 {len(legs_records) - MAX} 件"]
    warns = list(dict.fromkeys(tuple(w) for w in format_warns))  # 同一警告の重複を除く
    if warns:
        lines[0] += f" / ⚠️ 要確認 {len(warns)} 件"
        lines += ["", f"⚠️ 要確認 ({len(warns)} 件)"]
        for w, u in warns:
            lines.append(f"・{w}" + (f" <{u}|スレッドを開く>" if u else ""))
    if errors:
        lines += ["", f"❌ エラー ({len(errors)} 件)"]
        for e in list(errors)[:MAX]:
            lines.append(f"・{e}")
        if len(errors) > MAX:
            lines.append(f"…他 {len(errors) - MAX} 件")
    return "\n".join(lines)


def build_crash_notification(exc: BaseException, target_date: "date | None",
                             mentions: str = "") -> str:
    """異常終了時の通知テキスト。完了通知が出せないまま落ちた場合に使う。
    異常終了は常に ❌ failed なので、mentions があれば必ず先頭に付ける。"""
    tb = traceback.format_exception(type(exc), exc, exc.__traceback__)
    tail = "".join(tb[-3:]).strip()
    lines = ([f"{mentions} 対応をお願いします"] if mentions else []) + \
            ["❌ failed: zp_summary が異常終了しました "
             "(運行記録が更新されていない可能性があります)"]
    if target_date:
        lines.append(f"対象日: {target_date.isoformat()}")
    lines += [f"実行終了: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
              "", f"{type(exc).__name__}: {exc}", "", "```", tail[:1500], "```"]
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    # --config を先読みして既定値に流し込む
    pre_parser = argparse.ArgumentParser(add_help=False)
    pre_parser.add_argument("--config", default="config.json")
    config_path = pre_parser.parse_known_args()[0].config

    config: dict = {}
    if os.path.exists(config_path):
        try:
            with open(config_path, encoding="utf-8") as f:
                config = json.load(f)
            log(f"[設定] {config_path} を読み込み ({len(config)}項目)")
        except Exception as e:
            sys.exit(f"設定ファイル読み込み失敗 ({config_path}): {e}")

    p = argparse.ArgumentParser(
        description="Slackの運行記録チャンネルから Teams会議情報を集約して legs.json を生成")
    p.add_argument("--config", default="config.json", help="設定ファイル(JSON)のパス")
    p.add_argument("--channel", nargs="+",
                   help="チャンネルID/名前。複数可。config では文字列/配列")
    p.add_argument("--limit", type=int, default=20, help="取得する親メッセージ件数")
    p.add_argument("--since", help="この日付以降の投稿のみ (YYYY-MM-DD)")
    p.add_argument("--content-date",
                   help="Teams会議の実日付がこの日の投稿のみ (YYYY-MM-DD/today/yesterday/N_days_ago 等)")
    p.add_argument("--track-calendars", nargs="+", default=None,
                   help="共有予定表を引くドライバーのメール/UPN (Calendars.Read.Shared)")
    p.add_argument("--track-filter", nargs="+", default=None,
                   help="取り込む号車のホワイトリスト (例: giga03 giga04)。未指定なら全号車")
    p.add_argument("--append", action="store_true",
                   help="既存 legs.json に追記 (Trackname+日付+往路/復路 で重複スキップ)")
    p.add_argument("--out", default=None, help="結果JSONの出力先 (省略時は標準出力)")
    p.add_argument("--legs-out", default=None, help="legs.json の出力先")
    p.add_argument("--logs-out", default=None,
                   help="処理済み投稿の投稿日履歴。次回はこの最新日以降のみ取得")
    p.add_argument("--notify-webhook-url", default=None, help="完了通知用 Slack Webhook")
    p.add_argument("--notify-mentions", nargs="+", default=None,
                   help="❌ failed のときにメンションする相手 (メンバーID U... / "
                        "ユーザーグループID S... / here / channel)")
    p.add_argument("--start-notify-webhook-url", default=None, help="開始通知用 Slack Webhook")
    p.add_argument("--debug", action="store_true", help="詳細ログを表示")

    known = {a.dest for a in p._actions}
    for key in config:
        if key not in known:
            log(f"[警告] 設定ファイルの未知のキー: {key}")
    p.set_defaults(**{k: v for k, v in config.items() if k in known})
    return p.parse_args()


def run(args: argparse.Namespace, state: dict) -> int:
    """本体。戻り値は errors の件数 (0 以外なら main() が非ゼロ終了する)。

    state は異常終了時の通知に使う情報 (対象日) を main() へ渡すための箱。
    errors に積んだ異常は完了通知で ❌ failed として報告する。1 件でも積まれたら
    ✅ success は出さない (保存に失敗しているのに成功通知が飛ぶのを防ぐ)。
    """
    errors: list = []
    if not args.channel:
        sys.exit("エラー: --channel もしくは config.json の \"channel\" 指定が必要です")
    channels = [c for c in (args.channel if isinstance(args.channel, list)
                            else [args.channel]) if c]

    mentions = state.get("mentions", "")
    track_allow = parse_track_filter(getattr(args, "track_filter", None))
    # 正規形 (gigaNN) でない指定値はどの投稿とも一致せず、全件が対象外として
    # 消える。設定ミスに気付けるよう通知にも出す (打ち間違い・全角の打ち漏らし)
    invalid_filter = sorted(t for t in track_allow if not TRACK_CANONICAL_RE.match(t))
    if track_allow:
        log(f"[情報] track_filter = {', '.join(sorted(track_allow))} (対象外の号車は取り込まない)")
    filter_warns: list = []
    for t in invalid_filter:
        log(f"[警告] track_filter の値 '{t}' は号車として解釈できません (どの投稿とも一致しません)")
        filter_warns.append((
            f"config の track_filter の値 '{t}' は号車として解釈できません "
            f"(【giga05】形式で指定してください)。この値に一致する投稿はありません", ""))

    client = WebClient(token=get_token())
    from slack_sdk.http_retry.builtin_handlers import RateLimitErrorRetryHandler
    client.retry_handlers.append(RateLimitErrorRetryHandler(max_retry_count=5))

    since_ts = parse_since(args.since) if args.since else None
    target_date = resolve_relative_date(args.content_date)
    state["target_date"] = target_date
    if target_date:
        log(f"[情報] content_date = {target_date.isoformat()}")

    if args.start_notify_webhook_url:
        start = ["🚀 zp_summary の実行を開始しました"]
        if target_date:
            start.append(f"対象日: {target_date.isoformat()}")
        start.append(f"実行開始: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        send_slack_notification(args.start_notify_webhook_url, "\n".join(start))

    # logs.json の最新投稿日を since に反映して取得範囲を絞る
    existing_logs: list = []
    if args.logs_out:
        existing_logs, log_max_date = load_logs_json(args.logs_out)
        if log_max_date:
            # 対象日は投稿日 ±POST_WINDOW_DAYS で拾うので、logs.json 由来の since が
            # その下限より新しいと対象投稿が取得段階で落ちる (過去日を再実行すると
            # 通知の明細が減る) 。下限側にクランプして取り逃しを防ぐ
            since_date = log_max_date
            if target_date:
                since_date = min(since_date,
                                 target_date - timedelta(days=POST_WINDOW_DAYS))
            log_ts = datetime.combine(since_date, datetime.min.time()).timestamp()
            if since_ts is None or float(since_ts) < log_ts:
                since_ts = str(log_ts)
                log(f"[logs] since={since_date} を使用 "
                    f"(logs.json 最新投稿日 {log_max_date} / 既存 {len(existing_logs)} 件)")

    # Teams URL を含む親メッセージを収集
    posts: list = []
    for ch_spec in channels:
        try:
            ch_id = resolve_channel_id(client, ch_spec)
        except SystemExit as e:
            log(f"[警告] チャンネル '{ch_spec}' をスキップ: {e}")
            errors.append(f"チャンネル '{ch_spec}' を解決できずスキップしました: {e}")
            continue
        log(f"\n=== チャンネル {ch_spec} ({ch_id}) を処理 ===")
        try:
            messages = fetch_messages(client, ch_id, args.limit, since_ts)
        except SlackApiError as e:
            log(f"[警告] {ch_spec}: Slack API error: {e.response['error']}")
            errors.append(f"チャンネル '{ch_spec}' のメッセージ取得に失敗しました: "
                          f"Slack API error: {e.response['error']}")
            continue
        ch_posts = collect_teams_posts(messages, target_date=target_date)
        for post in ch_posts:
            post["_channel_id"] = ch_id
        log(f"  → {len(ch_posts)} 件のTeams投稿候補")
        posts.extend(ch_posts)

    if not posts:
        # 「本当に0件」と「取得に失敗して0件」を Slack 上で区別できるよう必ず通知する
        log("Teams URL を含む投稿が見つかりませんでした。")
        if args.notify_webhook_url:
            send_slack_notification(
                args.notify_webhook_url,
                build_notification([], target_date, 0, 0, filter_warns, errors,
                                   mentions=mentions))
        return len(errors)
    log(f"\n合計 {len(posts)} 件のTeams投稿を処理します...")

    extra_scopes = ["OnlineMeetings.Read", "Calendars.Read"]
    if args.track_calendars:
        extra_scopes.append("Calendars.Read.Shared")
    graph_token = get_graph_token(extra_scopes=extra_scopes)

    track_event_map: dict = {}
    if graph_token and args.track_calendars:
        drivers = args.track_calendars if isinstance(args.track_calendars, list) \
            else [args.track_calendars]
        log(f"\n=== ドライバー予定表をプリフェッチ ({len(drivers)}人) ===")
        track_event_map = fetch_track_calendars(drivers, graph_token, target_date,
                                                POST_WINDOW_DAYS)
        log(f"[track-cal] 合計 {len(track_event_map)} 件を JoinUrl で索引化")

    user_cache: dict = {}
    default_year = target_date.year if target_date else datetime.now().year

    # append: 既存 legs.json から dedupキー & URL集合を構築
    existing_legs: list = []
    existing_keys: set = set()
    existing_urls: set = set()
    legs_skipped_count = 0
    existing_dropped = 0
    if args.append and args.legs_out:
        for rec in read_json_list(args.legs_out):
            # 既存レコードにも track_filter を掛ける。掛けないと、絞り込みを始める前に
            # 取り込んだ対象外号車が legs.json に残り続ける (追記のたびに再出力される)
            if track_allow:
                rec_track = rec[0] if isinstance(rec, list) and rec \
                    else (rec.get("Trackname", "") if isinstance(rec, dict) else "")
                # 旧形式 '重要運行_giga06' はマーカーを外さないと号車が読めない
                if not track_allowed(strip_important_tag(rec_track), track_allow):
                    existing_dropped += 1
                    log(f"[legs] 対象外号車のため既存レコードを除外: 【{rec_track}】")
                    continue
            k = legs_dedup_key(rec)
            if any(k):
                if k in existing_keys:
                    legs_skipped_count += 1
                    continue
                existing_keys.add(k)
            existing_legs.append(rec)
            if isinstance(rec, list) and len(rec) >= 4 and isinstance(rec[3], dict):
                if rec[3].get("url"):
                    existing_urls.add(rec[3]["url"])
        if existing_legs:
            log(f"[legs] 既存 {len(existing_legs)} 件を読み込み")

    legs_records: list = []
    url_to_posted_iso: dict = {}
    format_warns: list = list(filter_warns)
    skipped_count = 0
    filtered_count = 0
    total = len(posts)
    for idx, f in enumerate(posts, 1):
        log(f"  [{idx}/{total}] {f.get('name', '?')[:60]}")
        msg = f.get("_msg") or {}
        posted_iso = (datetime.fromtimestamp(float(f["_posted_ts"])).isoformat(timespec="seconds")
                      if f.get("_posted_ts") else "")
        try:
            meta = extract_tracking_metadata(msg, client, user_cache, default_year)
            # ①Slack本文の号車で判定。Graph を叩く前なので、他号車の投稿が多い
            # チャンネルを追加しても Teams会議の取得回数と実行時間が増えない
            slack_track = meta.get("Trackname", "")
            if not track_allowed(slack_track, track_allow):
                filtered_count += 1
                log(f"    [filter] 対象外号車のためスキップ: 【{slack_track}】")
                continue
            teams_urls = extract_teams_urls(msg)
            if teams_urls:
                meeting = get_teams_meeting(teams_urls[0], graph_token)
                if not meeting and teams_urls[0] in track_event_map:  # 予定表フォールバック
                    ev = track_event_map[teams_urls[0]]
                    meeting = {"subject": ev.get("subject", ""),
                               "startDateTime": _iso_z(ev["start"]["dateTime"]),
                               "endDateTime": _iso_z(ev["end"]["dateTime"]),
                               "_source": "driverCalendar"}
                    log(f"    [track-cal] マッチ: {ev.get('subject', '')[:60]}")
                if meeting:
                    apply_teams_meeting(meta, meeting, args.debug)
            else:
                log("    [teams診断] Teams URL なし")

            # ②Teams会議件名で号車が確定した後にもう一度判定する。【重要運行】や
            # 号車なしの投稿は①を通り抜けるため、ここで見ないと対象外号車が混ざる
            teams_track = meta.get("Trackname", "")
            if not track_allowed(teams_track, track_allow):
                filtered_count += 1
                log(f"    [filter] 対象外号車のためスキップ (Teams会議件名で確定): 【{teams_track}】")
                if normalize_trackname(slack_track) in track_allow:
                    # Slack本文は対象号車なのに Teams側が対象外 = 会議室の流用や
                    # 投稿ミスの可能性。件数だけだと気付けないので明細を出す
                    format_warns.append((
                        f"Slack本文は対象号車 【{slack_track}】 ですが、Teams会議件名が "
                        f"対象外の 【{teams_track}】 のためスキップしました "
                        f"(track_filter): {meta.get('_title_line', '')[:40]}",
                        get_message_permalink(client, f.get("_channel_id", ""), f["_posted_ts"])
                        if f.get("_posted_ts") and f.get("_channel_id") else ""))
                continue

            if target_date is not None and meta.get("_title_date") != target_date:
                if meta.get("_title_date") is None:
                    # 日付が読めない投稿は対象日と必ず不一致になり丸ごと落ちる。無言で
                    # 消えると気付けないので ⚠️ 要確認として通知に載せる
                    log("    [警告] タイトル行から日付を読めずスキップ")
                    format_warns.append((
                        f"タイトル行の日付を読めずスキップしました "
                        f"(「8月12日」形式のみ対応): {meta.get('_title_line', '')[:60]}", ""))
                elif args.debug:
                    log(f"    [filter] 日付不一致でスキップ: {meta.get('_title_date')} != {target_date}")
                continue

            url = get_message_permalink(client, f.get("_channel_id", ""), f["_posted_ts"]) \
                if f.get("_posted_ts") and f.get("_channel_id") else ""
            if args.append and url and url in existing_urls:
                skipped_count += 1
                continue
            for w in meta.get("_warn", []):
                log(f"    [警告] {w}")
                format_warns.append((w, url))
            legs_records.append(build_legs_record(meta, url))
            if url:
                existing_urls.add(url)
                if posted_iso:
                    url_to_posted_iso[url] = posted_iso
        except Exception as e:
            # 投稿単位の失敗は他の投稿の処理を止めないが、無言で件数が減らないよう
            # 必ずログと通知に残す
            log(f"    [エラー] {type(e).__name__}: {e}")
            if args.debug:
                log(traceback.format_exc())
            errors.append(f"投稿の処理に失敗しました ({f.get('name', '?')[:40]}): "
                          f"{type(e).__name__}: {e}")
            continue

    # dedup してマージ
    new_legs: list = []
    for rec in legs_records:
        k = legs_dedup_key(rec)
        if any(k) and k in existing_keys:
            legs_skipped_count += 1
            continue
        new_legs.append(rec)
        if any(k):
            existing_keys.add(k)
    legs_new_count = len(new_legs)
    final_legs = existing_legs + new_legs
    if args.append:
        log(f"[append] 新規 {legs_new_count} 件、重複スキップ {skipped_count + legs_skipped_count} 件")

    # 出力: --out / 標準出力 / legs.json (いずれも final_legs)
    if args.out:
        atomic_write_json(args.out, final_legs)
        log(f"結果を {args.out} に保存しました。")
    elif not args.legs_out:
        print(json.dumps(final_legs, ensure_ascii=False, indent=2))
    if args.legs_out:
        try:
            atomic_write_json(args.legs_out, final_legs)
            log(f"[legs] {args.legs_out} に保存 (新規 {legs_new_count} / 重複スキップ "
                f"{legs_skipped_count} / 累計 {len(final_legs)} 件)")
        except Exception as e:
            # 保存できていないのに ✅ success が飛ばないよう errors に積む
            log(f"[エラー] legs.json 出力エラー: {e}")
            errors.append(f"{args.legs_out} の保存に失敗しました "
                          f"(運行記録 {legs_new_count} 件が未保存): {type(e).__name__}: {e}")

    # logs.json: 今回処理した投稿の投稿日履歴を追記
    if args.logs_out:
        try:
            seen_urls = {e.get("url") for e in existing_logs
                         if isinstance(e, dict) and e.get("url")}
            new_entries: list = []
            for rec in legs_records:
                if not (isinstance(rec, list) and len(rec) >= 4 and isinstance(rec[3], dict)):
                    continue
                url = rec[3].get("url", "")
                posted_at = url_to_posted_iso.get(url, "")
                if not url or url in seen_urls or not posted_at:
                    continue
                new_entries.append({"trackname": rec[0] if rec else "",
                                    "url": url, "posted_at": posted_at})
                seen_urls.add(url)
            final_logs = existing_logs + new_entries
            atomic_write_json(args.logs_out, final_logs)
            log(f"[logs] {args.logs_out} に保存 (新規 {len(new_entries)} / 累計 {len(final_logs)} 件)")
        except Exception as e:
            # 保存できないと次回の since がずれて取得漏れになるため errors に積む
            log(f"[エラー] logs.json 出力エラー: {e}")
            errors.append(f"{args.logs_out} の保存に失敗しました "
                          f"(次回実行の取得範囲がずれます): {type(e).__name__}: {e}")

    if args.notify_webhook_url:
        send_slack_notification(
            args.notify_webhook_url,
            build_notification(legs_records, target_date, legs_new_count,
                               legs_skipped_count + skipped_count, format_warns, errors,
                               filtered_count, existing_dropped, mentions))
    return len(errors)


def fallback_notify_webhook() -> str:
    """parse_args が失敗しても通知先を得るための最終手段。

    環境変数 ZP_NOTIFY_WEBHOOK_URL → config.json の生読み の順に探す。
    """
    url = os.environ.get("ZP_NOTIFY_WEBHOOK_URL", "")
    if url:
        return url
    path = "config.json"
    if "--config" in sys.argv:
        i = sys.argv.index("--config")
        if i + 1 < len(sys.argv):
            path = sys.argv[i + 1]
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f).get("notify_webhook_url", "") or ""
    except Exception:
        return ""


def fallback_notify_mentions() -> str:
    """parse_args が失敗しても メンション先を得るための最終手段。

    config.json 自体が壊れて落ちたときこそ人が気付く必要があるため、
    環境変数 ZP_NOTIFY_MENTIONS (空白/カンマ区切り) → config.json の生読み の順に探す。
    config.json が壊れて読めない場合に効くのは環境変数の方だけ。
    """
    env = os.environ.get("ZP_NOTIFY_MENTIONS", "")
    if env:
        return format_mentions(env.replace(",", " ").split())
    path = "config.json"
    if "--config" in sys.argv:
        i = sys.argv.index("--config")
        if i + 1 < len(sys.argv):
            path = sys.argv[i + 1]
    try:
        with open(path, encoding="utf-8") as f:
            return format_mentions(json.load(f).get("notify_mentions"))
    except Exception:
        return ""


def main() -> None:
    """引数解析 → run()。どこで落ちても Slack に ❌ failed を通知して非ゼロ終了する。

    通知先が分かる前 (config.json の読み込み失敗など) に落ちた場合は
    fallback_notify_webhook() / fallback_notify_mentions() で通知先を探す。
    """
    webhook, state, error_count = "", {}, 0
    try:
        args = parse_args()
        webhook = args.notify_webhook_url or ""
        # run() より前に確定させる。run() の途中で落ちてもクラッシュ通知でメンションできる
        state["mentions"] = format_mentions(getattr(args, "notify_mentions", None))
        error_count = run(args, state)
    except KeyboardInterrupt:
        raise
    except SystemExit as e:
        if not e.code:  # 正常終了 (exit 0 / exit None)
            raise
        _notify_crash(e, state, webhook)
        sys.exit(1)
    except Exception as e:
        _notify_crash(e, state, webhook)
        sys.exit(2)
    if error_count:
        # 完了通知で ❌ failed を報告済みなので、ここでは終了コードだけ立てる
        log(f"[終了] エラー {error_count} 件")
        sys.exit(1)


def _notify_crash(exc: BaseException, state: dict, webhook: str) -> None:
    mentions = state.get("mentions") or fallback_notify_mentions()
    text = build_crash_notification(exc, state.get("target_date"), mentions)
    log(text)
    send_slack_notification(webhook or fallback_notify_webhook(), text)


if __name__ == "__main__":
    main()

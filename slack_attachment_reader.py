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
# 正規の号車表記は 【giga05】(giga+2桁) と 【重要運行】のみ。
# 【GIGA05→06】(号車変更) 【GIGA6】【GIG05】等は自動判定せず Slack通知で警告するだけに留める
TRACK_CANONICAL_RE = re.compile(r'^(?:giga\d{2}|重要運行)$', re.IGNORECASE)
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
        sys.exit("環境変数 SLACK_BOT_TOKEN (xoxb-...) が設定されていません。")
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


def extract_tracking_metadata(msg: dict, client: WebClient, user_cache: dict,
                              default_year: int) -> dict:
    """親メッセージ本文から tracking フィールドを抽出 (返信なら _parent_msg を見る)。"""
    source_msg = msg.get("_parent_msg") or msg
    text = normalize_slack_text(source_msg.get("text", "") or "")
    result: dict = {}

    # タイトル行 = 【...】を含む最初の行
    first_line = next((ln for ln in text.split("\n") if "【" in ln and "】" in ln),
                      text.split("\n", 1)[0])

    m = TRACKNAME_RE.search(first_line)
    if m:
        track = m.group(1).strip()
        result["Trackname"] = track
        nm = TRACKNUM_RE.search(track)
        if nm:
            result["Track-num"] = nm.group(1)
        if not TRACK_CANONICAL_RE.match(track):
            result.setdefault("_warn", []).append(
                f"Slack本文の号車表記が非正規です: 【{track}】 (正規: 【giga05】形式)")

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
        # 非対話 (systemd / cron) でデバイスコード認証に入ると、誰もコードを入力
        # しないまま有効期限の 900 秒までポーリングし続ける。再認証が必要になった
        # 日から毎日ハングして無言で失敗するので、先に打ち切って原因を明示する
        if not sys.stdin.isatty():
            raise RuntimeError(
                "Microsoft の再認証が必要ですが、非対話実行のため完了できません。"
                "端末から ./run_zp_summary.sh を実行してデバイスコード認証を通し直して"
                f"ください (キャッシュ: {MSAL_CACHE_FILE})")
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
        # 中身はリフレッシュトークン。druid は sudo 可能な利用者が複数いるため、
        # .env と同じく本人のみ読める権限で作る (O_CREAT のモードは新規作成時
        # にしか効かないので、既存ファイル向けに chmod も行う)
        fd = os.open(MSAL_CACHE_FILE, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with open(fd, "w") as f:
            f.write(cache.serialize())
        os.chmod(MSAL_CACHE_FILE, 0o600)
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
    for k in ("Trackname", "Track-num", "Customer"):
        if tmeta.get(k):
            meta[k] = tmeta[k]
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

def _direction_of(luggage: str) -> str:
    """loaded_luggage から運行方向 (往路/復路) を取り出す。無ければ空文字。

    legs_dedup_key が重複判定に使うのと同じ語を見る。
    """
    return next((w for w in ("往路", "復路") if w in (luggage or "")), "")


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

    track = _giga(track)
    luggage = _strip_giga_tag(_giga(meta.get("Customer", "")))

    # 同日 2 便は leg[1] が同一値になり、csv_exported/x/main.js が leg[1] をキーに
    # 一意化するため UI 上で片方が消える。区別できる情報を leg[1] に含める
    direction = _direction_of(luggage)
    if date_str and direction:
        date_str = f"{date_str}({direction})"
    elif date_str and teams_start:
        # 方向が取れない同日 2 便 (日勤と夜勤など) は開始時刻で分ける。開始時刻は
        # Teams 会議から取るので再実行しても同じ値になり、重複スキップは効いたままになる
        date_str = f"{date_str}({teams_start.strftime('%H:%M')})"

    # 「重要運行」は loaded_luggage の【gigaXX】から番号を拾って Trackname=gigaXX 等に補正
    if track == "重要運行":
        m = re.search(r"giga(\d+)", _giga(meta.get("Customer", "")), re.IGNORECASE)
        if m:
            track = f"giga{m.group(1)}"
            track_num = m.group(1)
            if date_str:
                date_str = f"{date_str}_重要運行"

    return [
        track,
        f"{track_num}|{date_str}" if date_str else track_num,
        time_range,
        {
            "SW-version": meta.get("SW-ver", ""),
            "selfdrive_section": meta.get("Route", ""),
            "loaded_luggage": luggage,
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


def legs_dedup_key(rec) -> tuple:
    """重複判定キー (Trackname, 日付, 往路/復路)。配列/dict 両形式に対応。"""
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
    return (track, date_str, direction)


# ---------- I/O ----------

def read_json_list(path: str) -> list:
    """BOM対応でJSON配列を読む。存在しない/壊れている場合は []。"""
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
        return data if isinstance(data, list) else []
    except Exception as e:
        log(f"[警告] JSON読み込み失敗 ({path}): {e}")
        return []


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
    """一時ファイルに書いてから os.replace で差し替え (書き込み中クラッシュでも元は無傷)。

    出力先の親ディレクトリが無ければ作る (merge_legs.make_backup と同じ流儀)。
    """
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
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


def send_slack_notification(webhook_url: str, text: str) -> None:
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
                       format_warns: list = ()) -> str:
    """完了通知テキスト (✅ success / ❌ failed / 🔁 skipped) を組み立てる。
    legs_skipped_count は URL重複・dedupキー重複の両方を合算した件数。
    format_warns は号車表記の形式ずれ [(警告文, URL), ...] (自動補正はしない)。"""
    incomplete: set = set()
    if legs_new_count == 0 and legs_skipped_count > 0:
        # URL重複でレコード化前に落ちた分も含むため legs_records は空になり得る
        status = f"🔁 skipped: 全て既存レコードと重複のため追加なし ({legs_skipped_count} 件スキップ)"
    elif not legs_records:
        status = "✅ success: (今回追加されたレコードはありません)"
    else:
        for i, rec in enumerate(legs_records):
            ok, _ = is_legs_record_complete(rec)
            track = rec[0] if isinstance(rec, list) and rec else ""
            if not ok or track == "重要運行":  # giga番号未確定も failed 扱い
                incomplete.add(i)
        if incomplete:
            status = (f"❌ failed: 運行記録に不完全なレコードがあります "
                      f"({len(incomplete)}/{len(legs_records)} 件)")
        elif legs_skipped_count > 0:
            status = f"✅ success: {legs_new_count} 件追加 (重複スキップ {legs_skipped_count} 件)"
        else:
            status = "✅ success: 運行記録の読み込みが完了しました"

    lines = [status]
    if target_date:
        lines.append(f"対象日: {target_date.isoformat()}")
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
    if format_warns:
        lines[0] += f" / ⚠️ 号車表記の形式ずれ {len(format_warns)} 件"
        lines += ["", f"⚠️ 号車表記の形式ずれ ({len(format_warns)} 件) — Trackname を要確認"]
        for w, u in format_warns:
            lines.append(f"・{w}" + (f" <{u}|スレッドを開く>" if u else ""))
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
    p.add_argument("--append", action="store_true",
                   help="既存 legs.json に追記 (Trackname+日付+往路/復路 で重複スキップ)")
    p.add_argument("--out", default=None, help="結果JSONの出力先 (省略時は標準出力)")
    p.add_argument("--legs-out", default=None, help="legs.json の出力先")
    p.add_argument("--logs-out", default=None,
                   help="処理済み投稿の投稿日履歴。次回はこの最新日以降のみ取得")
    p.add_argument("--notify-webhook-url", default=None, help="完了通知用 Slack Webhook")
    p.add_argument("--start-notify-webhook-url", default=None, help="開始通知用 Slack Webhook")
    p.add_argument("--debug", action="store_true", help="詳細ログを表示")

    known = {a.dest for a in p._actions}
    for key in config:
        if key.startswith("_"):  # "_comment..." は雛形の注釈なので警告しない
            continue
        if key not in known:
            log(f"[警告] 設定ファイルの未知のキー: {key}")
    p.set_defaults(**{k: v for k, v in config.items() if k in known})
    return p.parse_args()


def main() -> None:
    args = parse_args()
    if not args.channel:
        sys.exit("エラー: --channel もしくは config.json の \"channel\" 指定が必要です")
    channels = [c for c in (args.channel if isinstance(args.channel, list)
                            else [args.channel]) if c]

    client = WebClient(token=get_token())
    from slack_sdk.http_retry.builtin_handlers import RateLimitErrorRetryHandler
    client.retry_handlers.append(RateLimitErrorRetryHandler(max_retry_count=5))

    since_ts = parse_since(args.since) if args.since else None
    target_date = resolve_relative_date(args.content_date)
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
            continue
        log(f"\n=== チャンネル {ch_spec} ({ch_id}) を処理 ===")
        try:
            messages = fetch_messages(client, ch_id, args.limit, since_ts)
        except SlackApiError as e:
            log(f"[警告] {ch_spec}: Slack API error: {e.response['error']}")
            continue
        ch_posts = collect_teams_posts(messages, target_date=target_date)
        for post in ch_posts:
            post["_channel_id"] = ch_id
        log(f"  → {len(ch_posts)} 件のTeams投稿候補")
        posts.extend(ch_posts)

    if not posts:
        log("Teams URL を含む投稿が見つかりませんでした。")
        return
    log(f"\n合計 {len(posts)} 件のTeams投稿を処理します...")

    extra_scopes = ["OnlineMeetings.Read", "Calendars.Read"]
    if args.track_calendars:
        extra_scopes.append("Calendars.Read.Shared")
    try:
        graph_token = get_graph_token(extra_scopes=extra_scopes)
    except Exception as e:
        # 開始通知だけ届いて以降が無音になると気づきにくいので、ここで失敗を伝える
        log(f"[異常] Microsoft Graph のトークン取得に失敗: {e}")
        if args.notify_webhook_url:
            send_slack_notification(
                args.notify_webhook_url,
                f"❌ zp_summary: Microsoft Graph の認証に失敗しました: {e}")
        sys.exit(1)

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
    if args.append and args.legs_out:
        for rec in read_json_list(args.legs_out):
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
    format_warns: list = []
    skipped_count = 0
    total = len(posts)
    for idx, f in enumerate(posts, 1):
        log(f"  [{idx}/{total}] {f.get('name', '?')[:60]}")
        msg = f.get("_msg") or {}
        posted_iso = (datetime.fromtimestamp(float(f["_posted_ts"])).isoformat(timespec="seconds")
                      if f.get("_posted_ts") else "")
        try:
            meta = extract_tracking_metadata(msg, client, user_cache, default_year)
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

            if target_date is not None and meta.get("_title_date") != target_date:
                if args.debug:
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
            if args.debug:
                log(f"[debug] エラー ({f.get('name')}): {type(e).__name__}: {e}")
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
            # 生成物を作るのがこのツールの目的なので、ここは成功扱いにしない。
            # exit 0 で続けると systemd の 2 段目 (merge_legs) が「生成物が
            # 存在しません」で落ち、原因から離れた場所にエラーが出る
            log(f"[異常] legs.json 出力エラー: {e}")
            if args.notify_webhook_url:
                send_slack_notification(
                    args.notify_webhook_url,
                    f"❌ zp_summary: 生成物 {args.legs_out} の書き込みに失敗しました: {e}")
            sys.exit(1)

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
            log(f"[警告] logs.json 出力エラー: {e}")

    if args.notify_webhook_url:
        send_slack_notification(
            args.notify_webhook_url,
            build_notification(legs_records, target_date, legs_new_count,
                               legs_skipped_count + skipped_count, format_warns))


if __name__ == "__main__":
    main()

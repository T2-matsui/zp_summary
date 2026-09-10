#!/usr/bin/env python3
"""
zp_summary の生成物を本番 legs.json へ安全にマージする。

設計方針 (安全側):
  - 既存レコードは一切変更・削除しない。順序も保持する。追記のみ。
  - 生成物のうち dedup キー (Trackname, 日付, 往路/復路) が既存に無いものだけを末尾に追加。
  - 出力は「1 行 1 レコード」形式。legs_tools/legs_server.py の行ベースパーサと互換を保つ。
  - 書き込み前に世代バックアップ、書き込みはアトミック (tmp + os.replace)。
  - 書き込み後に読み直して検証し、異常なら自動ロールバック。

差分チェック (いずれかに該当したら中断し、本番を書き換えない):
  1. 本番ファイルがパースできない
  2. 出力件数が既存件数を下回る
  3. 既存レコードの内容が 1 件でも変化している
  4. 追加件数が --max-add を超える (暴走検知)
  5. 生成物がパースできない / 配列でない

依存は標準ライブラリのみ。
"""

import argparse
import json
import os
import re
import shutil
import sys
import traceback
import unicodedata
import urllib.request
from datetime import datetime
from pathlib import Path

DEFAULT_TARGET = "/mnt/disks/dropoff/global_data/legs.json"
DEFAULT_SOURCE = "/mnt/disks/dropoff/zp_summary/legs_generated.json"
DEFAULT_BACKUP_DIR = "/mnt/disks/dropoff/backups/global_data"
DEFAULT_MAX_ADD = 50
INDENT = "    "  # 現行 legs.json のインデント (4 スペース) を踏襲


def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}", flush=True)


# ---------- legs 形式 ----------

GIGA_NORMALIZABLE_RE = re.compile(r'^giga\s*(\d{1,2})$', re.IGNORECASE)
# 号車変更表記 (giga05→06 等)。変更後 = 最後の号車を採用する
GIGA_CHANGE_RE = re.compile(
    r'^giga\s*\d{1,2}(?:\s*(?:→|⇒|➡|=>|->)\s*(?:giga\s*)?\d{1,2})+$', re.IGNORECASE)


def normalize_trackname(track: str) -> str:
    """号車表記を正規形 (半角小文字 giga + 2桁) に揃える。

    'GIGA05' 'ｇｉｇａ０５' 'giga5' → 'giga05' / '重要運行' → '重要運行'
    号車変更表記 'giga05→06' は変更後 (最後) の号車 'giga06' を採用する。
    slack_attachment_reader.normalize_trackname と同一ロジック。両方を揃えること。
    """
    s = unicodedata.normalize("NFKC", track or "").strip()
    m = GIGA_NORMALIZABLE_RE.match(s)
    if m:
        return f"giga{m.group(1).zfill(2)}"
    if GIGA_CHANGE_RE.match(s):
        return f"giga{re.findall(r'(\d{1,2})', s)[-1].zfill(2)}"
    return s


IMPORTANT_TAG_RE = re.compile(r'(?:^重要運行[_＿]|[_＿]重要運行$)')
# rec[1] 末尾の "(往路)" / "(復路)" (同日2便を UI 上で区別するために付けている)
DIRECTION_SUFFIX_RE = re.compile(r'\((?:往路|復路)\)\s*$')


def strip_important_tag(value: str) -> str:
    """過去の legs.json に残る「重要運行」マーカーを取り除く (重複判定でのみ使う)。

    旧形式 '重要運行_giga03' (〜2026-07-06) や '2026/06/22_重要運行' のレコードが
    本番に残っているため、マーカーの有無だけで同じ運行が二重登録されないよう、
    dedup キーの比較時に限って外す。出力する表記は変えない。
    slack_attachment_reader.strip_important_tag と同一ロジック。両方を揃えること。
    """
    return IMPORTANT_TAG_RE.sub("", value or "").strip()



def strip_direction_suffix(value: str) -> str:
    """rec[1] 末尾の "(往路)" / "(復路)" を取り除く (重複判定でのみ使う)。

    build_legs_record が同日2便を UI 上で区別するために付けているサフィックス。
    付ける前に本番へ入ったレコードと別キーになって二重登録されるため、比較時に外す。
    方向はキーの3要素目で持っているので情報は落ちない。
    時刻サフィックス "(HH:MM)" は外さない。方向が取れない同日2便 (日勤と夜勤など) は
    それが唯一の区別材料で、外すと2便目が重複扱いでスキップされる。
    slack_attachment_reader / merge_legs の両方に同じ関数がある。変更する場合は両方を揃えること。
    """
    return DIRECTION_SUFFIX_RE.sub("", value or "").strip()


def legs_dedup_key(rec) -> tuple:
    """重複判定キー (Trackname, 日付, 往路/復路)。

    Trackname は正規化して比較するため、本番の既存レコードが 'GIGA05' のような
    表記ゆれでも同一運行として重複判定され、二重登録されない。
    slack_attachment_reader.legs_dedup_key と同一ロジック。変更する場合は両方を揃えること。
    """
    if isinstance(rec, list) and len(rec) >= 4:
        track = rec[0] or ""
        date_part = rec[1] or ""
        d = rec[3] if isinstance(rec[3], dict) else {}
        luggage = d.get("loaded_luggage") or d.get("loaded liggage", "") or ""
    elif isinstance(rec, dict):
        track = rec.get("Trackname", "") or ""
        date_part = rec.get("Track-num|YY/MM/DD", "") or ""
        luggage = rec.get("loaded_luggage") or rec.get("loaded liggage", "") or ""
    else:
        return ("", "", "")
    date_str = date_part.split("|", 1)[1] if "|" in date_part else ""
    direction = next((w for w in ("往路", "復路") if w in luggage), "")
    # 「重要運行」マーカーと "(往路)" サフィックスは比較前に外す
    # (旧形式・サフィックス付与前のレコードとの二重登録を防ぐ)
    return (normalize_trackname(strip_important_tag(track)),
            strip_direction_suffix(strip_important_tag(date_str)), direction)


def dumps_legs(records: list) -> str:
    """1 行 1 レコード形式で直列化する (legs_server.py の行パーサ互換)。"""
    if not records:
        return "[]\n"
    body = ",\n".join(f"{INDENT}{json.dumps(r, ensure_ascii=False)}" for r in records)
    return "[\n" + body + "\n]\n"


def parse_as_legs_server(content: str) -> list:
    """legs_tools/legs_server.py の parse_legs_data と同じ行ベースパーサ。互換性検証用。"""
    import ast
    legs = []
    for line in content.strip().split("\n"):
        line = line.strip()
        if not line:
            continue
        if line.startswith("[") and (line.endswith("],") or line.endswith("]")):
            try:
                value = ast.literal_eval(line.rstrip(","))
            except (ValueError, SyntaxError):
                continue
            if isinstance(value, list):
                legs.append(value)
    return legs


def read_json_list(path: str, what: str) -> list:
    p = Path(path)
    if not p.exists():
        raise RuntimeError(f"{what} が存在しません: {path}")
    text = p.read_text(encoding="utf-8")
    if not text.strip():
        raise RuntimeError(f"{what} が空です: {path}")
    try:
        data = json.loads(text)
    except json.JSONDecodeError as e:
        raise RuntimeError(f"{what} が JSON としてパースできません: {path}: {e}") from e
    if not isinstance(data, list):
        raise RuntimeError(f"{what} が配列ではありません: {path} (type={type(data).__name__})")
    return data


# ---------- 書き込み ----------

def make_backup(target: str, backup_dir: str) -> str:
    Path(backup_dir).mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    dest = os.path.join(backup_dir, f"{Path(target).name}.{ts}")
    shutil.copy2(target, dest)
    return dest


def atomic_write(path: str, content: str) -> None:
    p = Path(path)
    tmp = p.with_name(p.name + ".merge.tmp")
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            f.write(content)
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


def prune_backups(backup_dir: str, target: str, keep: int) -> None:
    """このスクリプトが作った世代バックアップを keep 世代だけ残す。

    自分が作った "<target名>.YYYYMMDD-HHMMSS" だけを対象にする。手動で退避した
    "*.safe" のような別名のファイルは、名前が似ていても削除しない。
    """
    d = Path(backup_dir)
    if not d.is_dir() or keep <= 0:
        return
    pattern = re.compile(rf"^{re.escape(Path(target).name)}\.\d{{8}}-\d{{6}}$")
    files = sorted((f for f in d.iterdir() if f.is_file() and pattern.match(f.name)),
                   key=lambda f: f.name, reverse=True)
    for f in files[keep:]:
        try:
            f.unlink()
            log(f"[backup] 古い世代を削除: {f.name}")
        except OSError as e:
            log(f"[警告] バックアップ削除失敗 {f.name}: {e}")


def notify(webhook_url: str, text: str) -> None:
    if not webhook_url:
        # 未設定を黙って無視すると「通知が来ない」ことに気付けないため必ずログに残す
        log("[警告] 通知先 webhook が未設定です "
            "(--webhook-url もしくは MERGE_LEGS_WEBHOOK_URL を設定してください)")
        return
    try:
        req = urllib.request.Request(
            webhook_url,
            data=json.dumps({"text": text}).encode("utf-8"),
            headers={"Content-Type": "application/json"})
        urllib.request.urlopen(req, timeout=15).read()
    except Exception as e:  # 通知失敗でマージ自体は失敗させない
        log(f"[警告] Slack 通知失敗: {e}")


# ---------- 本体 ----------

def main() -> int:
    ap = argparse.ArgumentParser(description="zp_summary の生成物を本番 legs.json へ安全にマージする")
    ap.add_argument("--target", default=DEFAULT_TARGET, help="本番 legs.json")
    ap.add_argument("--source", default=DEFAULT_SOURCE, help="zp_summary の生成物")
    ap.add_argument("--backup-dir", default=DEFAULT_BACKUP_DIR, help="世代バックアップの保存先")
    ap.add_argument("--keep-backups", type=int, default=30, help="残すバックアップ世代数 (0 で無制限)")
    ap.add_argument("--max-add", type=int, default=DEFAULT_MAX_ADD,
                    help=f"1 回で追加を許容する最大件数 (既定 {DEFAULT_MAX_ADD})")
    ap.add_argument("--dry-run", action="store_true", help="本番を書き換えず結果だけ表示")
    ap.add_argument("--normalize-only", action="store_true",
                    help="マージせず、本番を 1 行 1 レコード形式に整形し直すだけ")
    ap.add_argument("--webhook-url", default=os.environ.get("MERGE_LEGS_WEBHOOK_URL", ""),
                    help="Slack Incoming Webhook URL (省略可)")
    args = ap.parse_args()

    mode = "normalize" if args.normalize_only else "merge"
    log(f"=== merge_legs 開始 (mode={mode}{', dry-run' if args.dry_run else ''}) ===")

    # --- 読み込み (チェック 1 / 5) ---
    try:
        existing = read_json_list(args.target, "本番 legs.json")
    except RuntimeError as e:
        log(f"[中断] {e}")
        notify(args.webhook_url, f"❌ legs.json マージ中断: {e}")
        return 1
    log(f"本番: {len(existing)} 件")

    added: list = []
    if not args.normalize_only:
        try:
            generated = read_json_list(args.source, "zp_summary 生成物")
        except RuntimeError as e:
            log(f"[中断] {e}")
            notify(args.webhook_url, f"❌ legs.json マージ中断: {e}")
            return 1
        log(f"生成物: {len(generated)} 件")

        existing_keys = {legs_dedup_key(r) for r in existing}
        skipped_dup = 0
        skipped_nokey = 0
        for rec in generated:
            key = legs_dedup_key(rec)
            if not any(key):
                # キーが立たないレコードは毎回重複追加されるため取り込まない
                skipped_nokey += 1
                continue
            if key in existing_keys:
                skipped_dup += 1
                continue
            existing_keys.add(key)
            added.append(rec)
        log(f"追加候補: {len(added)} 件 / 既存重複スキップ: {skipped_dup} 件 / "
            f"キー欠落スキップ: {skipped_nokey} 件")
        if skipped_nokey:
            log(f"[警告] Trackname・日付が取れないレコードが {skipped_nokey} 件あります。"
                f"{args.source} を確認してください")

        # --- チェック 4: 暴走検知 ---
        if len(added) > args.max_add:
            msg = (f"追加件数 {len(added)} 件が上限 {args.max_add} 件を超えています。"
                   f"生成物が異常な可能性があるため中断します")
            log(f"[中断] {msg}")
            notify(args.webhook_url, f"❌ legs.json マージ中断: {msg}")
            return 1

    merged = existing + added

    # --- チェック 2: 件数が減っていないか ---
    if len(merged) < len(existing):
        msg = f"出力件数 {len(merged)} が既存件数 {len(existing)} を下回りました"
        log(f"[中断] {msg}")
        notify(args.webhook_url, f"❌ legs.json マージ中断: {msg}")
        return 1

    # --- チェック 3: 既存レコードが 1 件も変化していないか ---
    for i, (before, after) in enumerate(zip(existing, merged)):
        if before != after:
            msg = f"既存レコード #{i} が変化しています (追記のみのはず)"
            log(f"[中断] {msg}")
            log(f"  before: {json.dumps(before, ensure_ascii=False)[:200]}")
            log(f"  after : {json.dumps(after, ensure_ascii=False)[:200]}")
            notify(args.webhook_url, f"❌ legs.json マージ中断: {msg}")
            return 1

    content = dumps_legs(merged)

    # --- 出力前に legs_server.py 互換を検証 ---
    reparsed = parse_as_legs_server(content)
    if len(reparsed) != len(merged):
        msg = (f"出力が legs_server.py のパーサと非互換です "
               f"(読めた {len(reparsed)} / 全 {len(merged)} 件)")
        log(f"[中断] {msg}")
        notify(args.webhook_url, f"❌ legs.json マージ中断: {msg}")
        return 1
    log(f"legs_server.py 互換チェック: {len(reparsed)}/{len(merged)} 件 OK")

    current = Path(args.target).read_text(encoding="utf-8")
    if content == current:
        log("変更なし。本番はそのままです")
        if not args.normalize_only:
            notify(args.webhook_url, f"🔁 legs.json 変更なし (既存 {len(existing)} 件)")
        return 0

    for rec in added:
        log(f"  + {json.dumps(rec, ensure_ascii=False)[:160]}")
    if args.normalize_only:
        log(f"整形のみ: {len(merged)} 件を 1 行 1 レコード形式に書き直します")

    if args.dry_run:
        log(f"[dry-run] 書き込みは行いません (既存 {len(existing)} → 出力 {len(merged)} 件)")
        return 0

    # --- 世代バックアップ → アトミック書き込み ---
    backup = make_backup(args.target, args.backup_dir)
    log(f"バックアップ: {backup}")
    atomic_write(args.target, content)

    # --- 書き込み後の検証。失敗したらロールバック ---
    try:
        verify = read_json_list(args.target, "書き込み後の legs.json")
        if len(verify) != len(merged):
            raise RuntimeError(f"件数不一致 (期待 {len(merged)} / 実際 {len(verify)})")
        if verify[:len(existing)] != existing:
            raise RuntimeError("既存レコードが保持されていません")
    except RuntimeError as e:
        log(f"[異常] 書き込み後の検証に失敗: {e}")
        shutil.copy2(backup, args.target)
        log(f"[ロールバック] {backup} から復元しました")
        notify(args.webhook_url, f"❌ legs.json マージ失敗・ロールバック済み: {e}")
        return 1

    prune_backups(args.backup_dir, args.target, args.keep_backups)

    if args.normalize_only:
        log(f"整形完了: {len(merged)} 件")
        notify(args.webhook_url, f"🧹 legs.json を 1 行 1 レコード形式に整形 ({len(merged)} 件)")
    else:
        log(f"マージ完了: {len(existing)} → {len(merged)} 件 (+{len(added)})")
        notify(args.webhook_url,
               f"✅ legs.json マージ完了: +{len(added)} 件 (累計 {len(merged)} 件)")
    return 0


def _fallback_webhook() -> str:
    """引数解析前に落ちても通知先を得るための最終手段 (環境変数 → sys.argv)。"""
    url = os.environ.get("MERGE_LEGS_WEBHOOK_URL", "")
    if url:
        return url
    if "--webhook-url" in sys.argv:
        i = sys.argv.index("--webhook-url")
        if i + 1 < len(sys.argv):
            return sys.argv[i + 1]
    return ""


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:  # 想定外は必ず非ゼロで落とす
        # バックアップ作成失敗・書き込みの I/O エラー等はここに来る。一番気付きたい
        # 障害なので、ログだけで終わらせず必ず Slack へ流す
        tail = "".join(traceback.format_exc().splitlines(keepends=True)[-3:]).strip()
        log(f"[異常終了] {type(e).__name__}: {e}")
        log(tail)
        notify(_fallback_webhook(),
               f"❌ legs.json マージが異常終了しました: {type(e).__name__}: {e}\n"
               f"```\n{tail[:1500]}\n```")
        sys.exit(2)

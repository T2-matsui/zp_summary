# zp_summary

Slack の運行記録チャンネルから Teams 会議URL を含む親メッセージを読み取り、Microsoft Teams 会議情報（および共有ドライバー予定表）と組み合わせて、運行記録 (Trackname / 顧客 / 区間 / 時刻 / ドライバー / SW-ver など) を構造化 JSON (`legs.json`) として出力する社内ツール。

systemd timer で日次実行し、Slack の Incoming Webhook で開始・完了通知を送る運用想定。

---

## 目次

1. [機能概要](#機能概要)
2. [出力ファイル](#出力ファイル)
3. [セットアップ](#セットアップ)
4. [運用コマンド](#運用コマンド)
5. [時間変更手順](#時間変更手順)
6. [新しいチャンネルを追加するとき](#新しいチャンネルを追加するとき)
7. [新しいドライバーを追加するとき](#新しいドライバーを追加するとき)
8. [取り込む号車を変えるとき](#取り込む号車を変えるとき)
9. [Slack通知を変えるとき](#slack通知を変えるとき)
10. [logs.json をリセットして再取得したいとき](#logsjson-をリセットして再取得したいとき)
11. [本番へ取り込む (merge_legs.py)](#本番へ取り込む-merge_legspy)
12. [トラブルシューティング](#トラブルシューティング)
13. [メンテナンス](#メンテナンス)

---

## 機能概要

### 動作フロー

```
1. systemd timer 起動 (例: 毎日 12:00)
   ↓
2. run_zp_summary.sh 実行
   ↓
3. Slack 開始通知 → 個人DM (🚀 実行開始)
   ↓
4. logs.json から最新投稿日を読み込み (since として使用)
   ↓
5. Slack の各チャンネルから Teams URL を含む親メッセージを取得
   ↓
6. ドライバーの共有予定表をプリフェッチ (JoinUrl Map 構築)
   ↓
7. 各投稿について:
   - Slack本文 → Driver / SW-ver / 区間 / 顧客 を抽出
   - track_filter 指定時: Slack本文の号車が対象外ならここで除外 (会議を引かない)
   - Teams 会議URL → 開始/終了時刻を取得
   - track_filter 指定時: 会議件名で確定した号車が対象外ならここでも除外
   ↓
8. 出力:
   - legs.json: 配列フォーマット (重複スキップ)
   - logs.json: 投稿日履歴の追記
   ↓
9. Slack 完了通知 → 運行チャンネル (✅ success / ❌ failed / 🔁 skipped)
```

### 完了通知のステータス

| ステータス | 条件 |
|---|---|
| `✅ success` | 新規レコードを追加 (問題なし)。一部重複した場合は「N 件追加 (重複スキップ M 件)」と併記 |
| `❌ failed` | **処理中にエラーが発生した** (優先)、または新規追加分に不完全なレコード (未取得項目 / giga番号未確定の「重要運行」) がある |
| `🔁 skipped` | 今回分が **全て既存と重複** で追加なし。success でも failed でもない中立表示。既存レコードの不完全さは failed 扱いにしない |

`notify_mentions` を設定すると、**`❌ failed` のときだけ**通知の先頭にメンションが付く (`✅ success` / `🔁 skipped` には付かない)。異常終了して完了通知が出せなかった場合の通知にも付く。
| `✅ success (対象なし)` | 対象日に該当投稿が 0 件 |

いずれのステータスでも `対象日: YYYY-MM-DD` の行が付く。エラーが 1 件でもあれば `✅ success` は出さず、末尾に `❌ エラー (N 件)` の明細が付く。人の確認が必要なだけの事象 (号車表記の補正など) はステータスを変えず `⚠️ 要確認 (N 件)` として列挙される。

重複には **URL一致** (同じスレッドの再取得) と **(Trackname, 日付, 往路/復路) 一致** (同じ運行の別スレッド再投稿) の 2 経路があり、どちらも `🔁 skipped` のスキップ件数に合算される。前者はレコード化前に落ちるため、`🔁 skipped` 通知に運行の明細行が並ばないことがある。

### 失敗したときに Slack へ流れるもの

トラブルに気付けることを優先し、**異常は全て運行チャンネル (`notify_webhook_url`) に流れる**。

| 失敗内容 | 通知 |
|---|---|
| 途中で異常終了した (Graph トークン期限切れ、Slack API 障害、想定外の例外) | `❌ failed: zp_summary が異常終了しました` + 例外名・メッセージ・traceback 末尾 |
| 起動時の設定不備 (`SLACK_TOKEN` 未設定 / config.json 読み込み失敗 / channel 未指定) | 同上 (通知先が分かる前に落ちた場合は `ZP_NOTIFY_WEBHOOK_URL` → `config.json` の生読みで通知先を探す) |
| `legs.json` / `logs.json` が壊れて読めない | `❌ failed` で**中断**する。空配列で続行して既存レコードを取りこぼしたまま上書きするのを防ぐため |
| `legs.json` の保存失敗 (ディスクフル・権限など) | `❌ failed` + 「運行記録 N 件が未保存」。**保存できていないのに `✅ success` は出ない** |
| `logs.json` の保存失敗 | `❌ failed` + 「次回実行の取得範囲がずれます」 |
| 投稿単位の処理失敗 | `❌ failed` + どの投稿で何の例外が出たか (他の投稿の処理は続行する) |
| チャンネルの解決失敗 / チャンネル単位の Slack API エラー | `❌ failed` + 対象チャンネル名 |
| 対象投稿が 0 件 | `✅ success (対象なし)` を必ず通知する。「本当に 0 件」と「取得に失敗して 0 件」を Slack 上で区別できるようにするため |
| 号車表記の補正・非正規表記・2台併記の取りこぼし | `⚠️ 要確認` (ステータスは変えない) |
| タイトル行の日付が読めずスキップした投稿 | `⚠️ 要確認` (日付は `8月12日` 形式のみ解釈できる。`8/12` は読めない) |

エラーがあった場合はプロセスも**非ゼロ終了**するため、systemd の `OnFailure=` でも検知できる ([systemd ユニット作成](#9-systemd-ユニット作成) 参照)。

通知経路そのものが落ちている場合 (webhook URL 未設定・Slack 側障害) は Slack へ出せないため、ログに `[警告] 通知先 webhook が未設定` / `[警告] Slack通知失敗` が残る。**通知が来ない日が続いたら、まず journalctl を確認する。**

### 号車表記 (Trackname) の扱い

正規の表記は `【giga05】` (半角小文字 giga + 2桁) と `【重要運行】` のみ。

| 入力 | 出力 | 通知 |
|---|---|---|
| `【giga05】` | `giga05` | なし |
| `【GIGA05】` `【Giga05】` `【ｇｉｇａ０５】` `【giga5】` | `giga05` に**正規化** | `⚠️ 要確認` に「補正しました」 |
| `【giga05→06】` `【GIGA05→06】` (号車変更) | **変更後 (最後) の号車** `giga06` を採用 | `⚠️ 要確認` に「号車変更表記から変更後の号車を採用しました」 |
| `【giga100】` `【giga05-06】` `【giga05〜06】` | 補正せずそのまま | `⚠️ 要確認` に「自動補正できません」 |
| `【giga05】【giga06】` (2台併記) | 先頭の `giga05` のみ | `⚠️ 要確認` に「取りこぼしています」 |

号車変更として扱うのは矢印表記 (`→` `⇒` `➡` `->` `=>`) のみ。`giga05-06` `giga05〜06` は「変更」か「範囲」か判別できないため補正せず警告する。

正規化は **Slack本文と Teams会議件名の両方**に掛かる。重複判定キー (`legs_dedup_key`) も正規化後の値で比較するため、`GIGA05` と `giga05` は同一運行として扱われ二重登録されない (`merge_legs.py` 側も同一ロジック。**片方だけ変更しないこと**)。

#### `【重要運行】` の扱い

`【重要運行】` は号車名ではなく「まだ号車が決まっていない」という印。会議件名の `【gigaXX】`、それも無ければ運行名に含まれる `gigaXX` から号車を拾って `giga03` などに解決する。

**「重要運行だった」という情報はレコードに残さない。** 下流の zero-plotter が Trackname を号車ID としてそのまま使っており (`legs_record_table.js` は `${vehicleId}_${dataSource}` で Druid のデータソース名を組み立て、号車フィルタと動画ディレクトリ検索も完全一致で引く)、`giga03_重要運行` のような値にすると号車として認識されなくなるため。

どこからも号車を拾えなかった場合だけ Trackname が `重要運行` のまま残り、`❌ failed` (Trackname未確定) になる。

過去の legs.json には `重要運行_giga03` (〜2026-07-06 の形式) や日付末尾の `_重要運行` が残っている。重複判定 (`legs_dedup_key`) はこのマーカーを外してから比較するため、同じ運行が二重登録されることはない (`merge_legs.py` 側も同一ロジック)。

Slack本文が `【重要運行】` で Teams会議件名に号車がある場合は、Teams 側の号車が採用される (これは正常フロー)。**他号車の会議室を流用して投稿された場合は誤った号車が確定するが、判別する材料がないため検知できない。**

### 号車の絞り込み (`track_filter`)

チャンネルに複数号車の投稿が流れる場合、`track_filter` で**取り込む号車を限定**できる (未指定なら全号車が対象で、挙動は従来どおり)。

```json
"track_filter": ["giga03", "giga04"]
```

指定値も投稿側の表記も `normalize_trackname()` で正規化してから比較するため、`GIGA03` `ｇｉｇａ０３` `giga3` はすべて `giga03` として一致する。

#### 判定は 2 回行う

号車は Slack本文と Teams会議件名の2箇所から決まり、**会議件名の方が強い** (「号車表記 (Trackname) の扱い」参照)。そのため判定も 2 段階になっている。

| | いつ | 何のため |
|---|---|---|
| ① | Slack本文を読んだ直後 (会議を引く前) | 対象外と分かっている投稿で Teams会議を取得しない (他号車の投稿が多いチャンネルを追加しても API 呼び出しと実行時間が増えない) |
| ② | 会議件名で号車が確定した後 | `【重要運行】`・`【giga100】`・号車なしの投稿は①を素通りするため。**②が無いと対象外号車が legs.json に混ざる** |

| Slack本文の号車 | 会議件名の号車 | `["giga03","giga04"]` 指定時 |
|---|---|---|
| `【giga03】` `【GIGA03】` `【giga3】` | `【giga03】` | ✅ 取り込む |
| `【giga06】` | (引かない) | **①で除外** |
| `【giga05→06】` (号車変更) | `【giga06】` | ✅ 取り込む (`giga06` として判定) |
| `【giga05→06】` (号車変更) | — | **①で除外** (`giga06` は対象外のため) |
| `【重要運行】` `【giga100】` `【　】`なし | `【giga03】` | ✅ 取り込む (号車は会議件名で確定) |
| `【重要運行】` `【giga100】` `【　】`なし | `【giga06】` | **②で除外** |
| `【重要運行】` `【giga100】` `【　】`なし | 会議が引けない | ✅ 除外しない。従来どおり `⚠️ 要確認` / `❌ failed` に出す |
| `【giga03】` (対象) | `【giga06】` (対象外) | **②で除外**し、`⚠️ 要確認` に明細を出す (会議室の流用・投稿ミスの可能性) |

**号車が確定していない投稿は絞り込みで消さない**方針。無言で件数が減るより、`⚠️ 要確認` / `❌ failed` として人が見て判断する方を優先している。

#### 通知の出方

除外した投稿は完了通知に件数だけ出る (毎回同じ号車が並んで `⚠️ 要確認` が埋まるのを避けるため明細は出さない)。

```
対象日: 2026-09-09
対象外号車のためスキップ: 5 件 (track_filter)
```

例外は上表の最終行 (Slack本文は対象号車なのに会議件名が対象外) で、これだけは `⚠️ 要確認` に明細と URL が出る。

#### 既存 legs.json に残っている対象外号車も取り除く

`--append` (config の `"append": true`) で既存 legs.json に追記する場合、**既存レコードにも `track_filter` を掛ける**。絞り込みを始める前に取り込んだ対象外号車が残り続けないようにするため。取り除いた件数は通知に出る。

```
対象外号車のため既存 legs.json から除外: 3 件 (track_filter)
```

`append` が false のときは毎回 legs.json を作り直すため、そもそも対象外号車は残らない。**本番 legs.json (merge_legs.py の取り込み先) からは削除しない** — merge_legs.py は追記専用で、既存レコードを消さない設計のため。

`track_filter` に `gigaNN` として解釈できない値 (`giga3O` のような打ち間違い等) を書くと、**その値はどの投稿とも一致せず全件が除外され得る**。気付けるよう、起動時に警告し `⚠️ 要確認` にも出す。

```
・config の track_filter の値 'giga3O' は号車として解釈できません (【giga05】形式で指定してください)。この値に一致する投稿はありません
```

#### 取り込み直すとき

除外された投稿は logs.json にも記録されない。そのため `track_filter` から外したうえで、その投稿日を含む `content_date` を指定して再実行すれば取り込み直せる (取得範囲は対象日の投稿日 ±14 日)。

### 認証

| 認証情報 | 用途 | scope/権限 |
|---|---|---|
| Slack Bot Token (`xoxb-`) | チャンネル読み取り | `channels:history` `channels:read` `groups:history` `groups:read` `users:read` `incoming-webhook` |
| Slack Webhook URL (開始) | 個人DM通知 | scope不要 (URL自体が認証情報) |
| Slack Webhook URL (完了) | 運行チャンネル通知 | scope不要 |
| Microsoft Graph (MSAL) | Teams 会議情報 / ドライバー予定表 | `OnlineMeetings.Read` `Calendars.Read` `Calendars.Read.Shared` |

書き込み・削除権限は一切持たない (Read only)。

---

## 出力ファイル

### `legs.json`
往路/復路ごとに区切った配列フォーマット。**本ツールの成果物**。

```json
[
  ["GIGA03", "03|2026/06/02",
   "2026-06-02T13:00:00.000+09:00/2026-06-02T17:00:00.000+09:00",
   {"SW-version": "v1.2.3", "selfdrive_section": "東京-大阪",
    "loaded_luggage": "QP様", "url": "https://..."}]
]
```

**append と重複スキップの挙動**

| `config.json` の `append` | 動作 |
|---|---|
| `false` (既定) | 既存 legs.json を読まず、今回分だけで **上書き** |
| `true` | 既存 legs.json に **追記**。重複は自動スキップし、既存内の重複も掃除する |

重複判定キーは **(Trackname, 日付, 往路/復路)**。同じ運行を再取得しても `append: true` なら二重登録されない。過去分から作り直したい場合は下記「[logs.json をリセット](#logsjson-をリセットして再取得したいとき)」を参照。

### `logs.json`
既処理スレッドの投稿日履歴。次回実行時の高速化に使用。

```json
[
  {"trackname": "GIGA03",
   "url": "https://...",
   "posted_at": "2026-05-29T14:00:00+09:00"}
]
```

---

## セットアップ

### 1. ディレクトリ構成

```
~/Downloads/zp_summary/      ← 運用ディレクトリ
├── slack_attachment_reader.py   ← メインスクリプト
├── merge_legs.py                ← 生成物を本番 legs.json へマージ (手動実行)
├── calendar_probe.py            ← 動作確認用
├── run_zp_summary.sh            ← systemd から呼ばれる起動スクリプト
├── config.json                  ← 設定 (秘密、git除外)
├── config.json.example          ← 雛形 (git管理)
├── .env                         ← 認証情報 (秘密、git除外)
├── .env.example                 ← 雛形 (git管理)
├── legs.json                    ← 出力 (git除外)
├── logs.json                    ← 出力 (git除外)
├── README.md
└── .gitignore
```

### 2. 依存ライブラリ

```bash
pip3 install slack_sdk requests msal python-dotenv --break-system-packages
```

### 3. `.env` 作成

```bash
cp .env.example .env
nano .env
```

```env
SLACK_BOT_TOKEN=xoxb-...実際のBot Token...
```

### 4. `config.json` 作成

```bash
cp config.json.example config.json
nano config.json
```

```json
{
  "channel": ["C0XXX", "C0YYY"],
  "limit": 200,
  "content_date": "yesterday",
  "track_calendars": [
    "driver-a@example.com",
    "driver-b@example.com"
  ],
  "track_filter": ["giga03", "giga04"],
  "legs_out": "legs.json",
  "logs_out": "logs.json",
  "notify_webhook_url": "https://hooks.slack.com/services/XXX/YYY/ZZZ",
  "start_notify_webhook_url": "https://hooks.slack.com/services/AAA/BBB/CCC"
}
```

| キー | 内容 |
|---|---|
| `channel` | 対象 Slack チャンネル ID の配列 |
| `content_date` | 対象日 (`today` `yesterday` `tomorrow` `YYYY-MM-DD` `N_days_ago` 等) |
| `track_calendars` | ドライバーのメール/UPN (カレンダー共有が必要) |
| `track_filter` | 取り込む号車のホワイトリスト (省略時は全号車) |
| `notify_webhook_url` | 完了通知用 Webhook URL (運行チャンネル) |
| `notify_mentions` | `❌ failed` のときにメンションする相手 (メンバーID `U...` の配列) |
| `start_notify_webhook_url` | 開始通知用 Webhook URL (個人DM) |

### 5. Bot をチャンネルに招待

各 channel ID で:

```
/invite @アプリ名
```

Slack のメッセージ欄から実行。

### 6. ドライバーカレンダー共有

各ドライバーが Outlook で「予定表の共有」→ 「閲覧可能」または「編集可能」で実行ユーザーに共有しておく。

### 7. Microsoft Graph 認証

初回実行時にデバイスコード認証が起動:

```bash
python3 slack_attachment_reader.py --config config.json
```

ターミナルに表示される URL を開き、コードを入力 → 同意画面で承諾。
キャッシュは `~/.slack_attachment_reader_msal_cache.bin` に保存され、約90日間は自動更新。

### 8. run_zp_summary.sh 作成

```bash
nano ~/Downloads/zp_summary/run_zp_summary.sh
```

```bash
#!/bin/bash
cd /home/$USER/Downloads/zp_summary
python3 slack_attachment_reader.py --config config.json
```

実行権限付与:

```bash
chmod +x ~/Downloads/zp_summary/run_zp_summary.sh
```

### 9. systemd ユニット作成

#### service ファイル

```bash
nano ~/.config/systemd/user/zp-summary.service
```

```ini
[Unit]
Description=Run zp_summary slack attachment reader
OnFailure=zp-summary-failure.service

[Service]
Type=oneshot
ExecStart=%h/Downloads/zp_summary/run_zp_summary.sh
```

#### 失敗通知ユニット (OnFailure)

スクリプト自体が起動しない・OOM kill された等、**本体が Slack 通知を出せずに落ちた場合の最後の砦**。
`ZP_NOTIFY_WEBHOOK_URL` に運行チャンネルの webhook URL を入れておく。

```bash
nano ~/.config/systemd/user/zp-summary-failure.service
```

```ini
[Unit]
Description=Notify Slack when zp-summary.service fails

[Service]
Type=oneshot
Environment=ZP_NOTIFY_WEBHOOK_URL=https://hooks.slack.com/services/XXX/YYY/ZZZ
ExecStart=/bin/bash -c 'curl -sS -X POST -H "Content-type: application/json" \
  --data "{\"text\":\"❌ failed: zp-summary.service が異常終了しました (systemd 検知)\\njournalctl --user -u zp-summary.service -n 50 で確認してください\"}" \
  "$ZP_NOTIFY_WEBHOOK_URL"'
```

本体がエラーを検知した場合はプロセスが非ゼロ終了するので、この経路でも通知が飛ぶ (Slack 側では本体の `❌ failed` と 2 通並ぶ)。

#### timer ファイル

```bash
nano ~/.config/systemd/user/zp-summary.timer
```

```ini
[Unit]
Description=Schedule zp_summary

[Timer]
OnCalendar=*-*-* 12:00:00
Persistent=true

[Install]
WantedBy=timers.target
```

#### 有効化

```bash
systemctl --user daemon-reload
systemctl --user enable --now zp-summary.timer
```

#### 次回予定確認

```bash
systemctl --user list-timers --no-pager | grep zp-summary
```

`NEXT` 列に時刻が出ていればOK。

---

## 運用コマンド

### 状態確認

```bash
# タイマー次回実行予定
systemctl --user list-timers --no-pager | grep zp-summary

# タイマー設定確認
systemctl --user cat zp-summary.timer --no-pager

# サービス状態
systemctl --user status zp-summary.service --no-pager
```

### ログ確認

```bash
# 最新50行
journalctl --user -u zp-summary.service -n 50 --no-pager

# 指定時刻以降
journalctl --user -u zp-summary.service --since "12:00" --no-pager

# リアルタイム監視
journalctl --user -u zp-summary.service -f

# 警告・エラーだけ抽出
journalctl --user -u zp-summary.service -n 200 --no-pager | grep -E "警告|エラー|Error|Traceback"
```

### 手動実行

```bash
# サービスを今すぐ起動 (timer を介さず)
systemctl --user start zp-summary.service

# シェルから直接 (デバッグ向け)
cd ~/Downloads/zp_summary
python3 slack_attachment_reader.py --config config.json 2>&1 | tee run.log
```

### 停止・再開

```bash
# 一時停止 (タイマー OFF、設定は残る)
systemctl --user stop zp-summary.timer

# 再開
systemctl --user start zp-summary.timer

# 自動起動も無効化
systemctl --user disable zp-summary.timer

# 再有効化 + 即時起動
systemctl --user enable --now zp-summary.timer
```

---

## 時間変更手順

### Step 1: timer ファイル編集

```bash
nano ~/.config/systemd/user/zp-summary.timer
```

`OnCalendar=` 行を希望時刻に変更:

```ini
[Timer]
OnCalendar=*-*-* 09:00:00      ← ここを編集 (例: 朝9時)
Persistent=true
```

`Ctrl+O` → Enter → `Ctrl+X` で保存。

### Step 2: 構文チェック (任意)

```bash
systemd-analyze calendar "*-*-* 09:00:00"
```

`Next elapse: ...` が表示されればOK。`Failed to parse calendar specification` が出たら書式エラー。

### Step 3: 反映

```bash
systemctl --user daemon-reload
systemctl --user restart zp-summary.timer
```

**両方必要**。`daemon-reload` だけだとタイマーは古い設定のまま。

### Step 4: 確認

```bash
systemctl --user list-timers --no-pager | grep zp-summary
```

`NEXT` 列が新しい時刻になっていれば完了。

### 時刻指定パターン

```ini
# 毎日 朝9時
OnCalendar=*-*-* 09:00:00

# 平日 (月〜金) のみ 朝7時
OnCalendar=Mon..Fri *-*-* 07:00:00

# 平日 朝7時 と 夕方17時 (2回)
OnCalendar=Mon..Fri *-*-* 07:00:00
OnCalendar=Mon..Fri *-*-* 17:00:00

# 月曜だけ朝9時 (週次)
OnCalendar=Mon *-*-* 09:00:00

# 1時間ごと (毎時00分)
OnCalendar=hourly

# 30分ごと
OnCalendar=*-*-* *:00,30:00

# 15分ごと
OnCalendar=*-*-* *:0/15:00
```

---

## 新しいチャンネルを追加するとき

### Step 1: Bot をチャンネルに招待

Slack でそのチャンネルを開き、メッセージ欄で:

```
/invite @アプリ名
```

### Step 2: config.json の channel 配列に追加

```bash
nano ~/Downloads/zp_summary/config.json
```

```json
"channel": ["C0XXX", "C0YYY", "C0ZZZ"],
```

新しい channel ID をリストに追加 (Slack のチャンネル設定画面の下部に表示される `Channel ID`)。

### Step 3: JSON 構文チェック

```bash
python3 -m json.tool ~/Downloads/zp_summary/config.json > /dev/null && echo "JSON OK" || echo "JSON NG"
```

### Step 4: 動作確認

```bash
systemctl --user start zp-summary.service
sleep 5
journalctl --user -u zp-summary.service -n 30 --no-pager
```

ログに新しいチャンネルが処理されていればOK。

---

## 新しいドライバーを追加するとき

### Step 1: そのドライバーから予定表共有を受ける

ドライバーが Outlook で「予定表の共有」→ 「閲覧可能」 (または編集可能) で実行ユーザーに共有。

### Step 2: 共有が反映されたか確認 (任意)

```bash
cd ~/Downloads/zp_summary
python3 calendar_probe.py 新しいドライバー@example.com
```

`Test 2: 新しいドライバー@example.com のカレンダー読み取り` が `status: 200` ならOK。

### Step 3: config.json の track_calendars に追加

```bash
nano config.json
```

```json
"track_calendars": [
  "driver-a@example.com",
  "driver-b@example.com",
  "新しいドライバー@example.com"
],
```

### Step 4: 動作確認

```bash
systemctl --user start zp-summary.service
journalctl --user -u zp-summary.service --since "1min ago" --no-pager | grep "track-cal"
```

新しいドライバーの行が出ていればOK。

---

## 取り込む号車を変えるとき

チャンネルに他号車の投稿も流れていて、**特定の号車だけを legs.json に入れたい**場合に設定する (仕組みは [号車の絞り込み](#号車の絞り込み-track_filter))。設定しなければ全号車が対象。

### Step 1: config.json に `track_filter` を書く

```bash
nano ~/Downloads/zp_summary/config.json
```

```json
"track_filter": ["giga03", "giga04"],
```

- **正規表記 (`gigaNN`) で書く。** `giga3O` (数字の 0 と英字の O) のような値はどの投稿とも一致せず、全件が除外され得る
- 絞り込みをやめるときは行ごと消すか `"track_filter": []` にする

### Step 2: JSON 構文チェック

```bash
python3 -m json.tool ~/Downloads/zp_summary/config.json > /dev/null && echo "JSON OK" || echo "JSON NG"
```

### Step 3: 本番を汚さずに動作確認

`--legs-out` `--logs-out` を一時ファイルに向け、`--notify-webhook-url` を付けなければ、**本番の legs.json / logs.json と Slack 通知に影響せず**に確認できる。

```bash
cd ~/Downloads/zp_summary
.venv/bin/python slack_attachment_reader.py \
  --content-date yesterday \
  --track-filter giga03 giga04 \
  --legs-out /tmp/legs_test.json --logs-out /tmp/logs_test.json
```

ログの `[filter] 対象外号車のためスキップ` 行と、`/tmp/legs_test.json` に入った号車を確認する。

```bash
python3 -c "import json;print([r[0] for r in json.load(open('/tmp/legs_test.json'))])"
```

### Step 4: 反映

config.json を保存すれば次回の timer 実行から効く (systemd の再読み込みは不要)。すぐ試すなら:

```bash
systemctl --user start zp-summary.service
journalctl --user -u zp-summary.service --since "1min ago" --no-pager | grep -E "track_filter|filter\]"
```

### 除外しすぎた / 除外できていないとき

| 症状 | 原因 | 対処 |
|---|---|---|
| 対象の号車が入らない | `track_filter` の打ち間違い | 通知の `⚠️ 要確認` に「号車として解釈できません」が出ていないか確認 |
| 対象外の号車が入る | 号車が `【重要運行】` で会議件名も対象号車になっている | 仕様どおり (会議件名が正) 。会議件名を直す |
| 対象の号車が「会議件名が対象外」でスキップされる | 他号車の会議室を流用している | 会議件名を正しい号車に直して再実行 |
| 除外した分を取り込み直したい | — | `track_filter` から外し、`--content-date YYYY-MM-DD` を指定して再実行 (±14日以内の投稿) |

---

## Slack通知を変えるとき

### Webhook URL を新規発行する場合

1. https://api.slack.com/apps → 該当アプリ → **Incoming Webhooks**
2. **Add New Webhook to Workspace**
3. 通知先 (チャンネル or 個人DM) を選択
4. 「許可する」
5. リストに追加された URL をコピー

### config.json で URL を切り替え

```bash
nano ~/Downloads/zp_summary/config.json
```

- `notify_webhook_url` を変更 → 完了通知先を変更
- `start_notify_webhook_url` を変更 → 開始通知先を変更

### 失敗時にメンバーをメンションする

`❌ failed` が普通の投稿に埋もれて気付かれないのを防ぐため、失敗時だけメンションを飛ばせる。

**Step 1: メンバーIDを調べる**

Slack でその人のプロフィールを開く → 「その他」(⋮) → **メンバーIDをコピー**。`U01ABCDEF` のような文字列。

**Step 2: config.json に書く**

```json
"notify_mentions": ["U01ABCDEF", "U02GHIJKL"],
```

| 書ける値 | 変換後 |
|---|---|
| `U01ABCDEF` (メンバーID) | `<@U01ABCDEF>` |
| `S01XYZABC` (ユーザーグループID) | `<!subteam^S01XYZABC>` |
| `here` / `channel` | `<!here>` / `<!channel>` |

**表示名 (`@田中`) は書かない。** Webhook 側で ID に解決されず通知が飛ばないため、`[警告]` を出して無視する (本文にも出さない)。

**Step 3: 動作確認**

```bash
cd ~/Downloads/zp_summary
MENTION=$(python3 -c "import json; print(' '.join('<@%s>' % m for m in json.load(open('config.json')).get('notify_mentions', [])))")
END_URL=$(python3 -c "import json; print(json.load(open('config.json')).get('notify_webhook_url', ''))")
curl -X POST -H 'Content-type: application/json' \
  --data "{\"text\":\"$MENTION メンションテスト\"}" "$END_URL"
```

対象の人に通知が飛べばOK (名前が青くリンクになっていればID解決に成功している)。

> config.json 自体が壊れて起動時に落ちた場合は config を読めないため、メンションも飛ばせない。
> その状況でも飛ばしたいときは環境変数 `ZP_NOTIFY_MENTIONS="U01ABCDEF,U02GHIJKL"` を
> systemd ユニットに設定しておく (`ZP_NOTIFY_WEBHOOK_URL` と同じ用途)。

### Webhook 単独テスト

```bash
cd ~/Downloads/zp_summary

START_URL=$(python3 -c "import json; print(json.load(open('config.json')).get('start_notify_webhook_url', ''))")
END_URL=$(python3 -c "import json; print(json.load(open('config.json')).get('notify_webhook_url', ''))")

# 開始通知
curl -X POST -H 'Content-type: application/json' --data '{"text":"DM テスト"}' "$START_URL"
echo ""

# 完了通知
curl -X POST -H 'Content-type: application/json' --data '{"text":"チャンネル テスト"}' "$END_URL"
echo ""
```

両方とも `ok` が返り、Slack側に届けば設定完了。

---

## logs.json をリセットして再取得したいとき

logs.json があると `since` フィルタが効いて古い投稿は取得されません。過去分も含めて再取得したい場合:

```bash
# バックアップ
cp ~/Downloads/zp_summary/logs.json ~/Downloads/zp_summary/logs.json.bak

# 削除
rm ~/Downloads/zp_summary/logs.json

# 実行
systemctl --user start zp-summary.service
```

次回実行で logs.json が新規作成され、`content_date` の対象範囲全部が処理される。

### legs.json も全部やり直したいなら

```bash
cp ~/Downloads/zp_summary/legs.json ~/Downloads/zp_summary/legs.json.bak
rm ~/Downloads/zp_summary/legs.json ~/Downloads/zp_summary/logs.json
systemctl --user start zp-summary.service
```

---

## 本番へ取り込む (merge_legs.py)

`merge_legs.py` は zp_summary の生成物を本番 `legs.json` へマージする。**手動実行**であり、`run_zp_summary.sh` からは呼ばれない。

### 設計 (安全側)

- **追記のみ。** 既存レコードは変更・削除せず、順序も保持する
- 追加するのは dedup キー `(Trackname, 日付, 往路/復路)` が既存に無いものだけ (Trackname は正規化して比較)
- 書き込み前に**世代バックアップ** → アトミック書き込み (tmp + `os.replace`) → 書き込み後に読み直して検証 → 異常なら**自動ロールバック**
- 出力は「1 行 1 レコード」形式。`legs_tools/legs_server.py` の行ベースパーサ互換を書き込み前に検証する

以下のいずれかに該当したら**本番を書き換えずに中断**し、Slack へ `❌ legs.json マージ中断` を通知する。

1. 本番ファイルがパースできない / 空 / 配列でない
2. 生成物がパースできない / 空 / 配列でない
3. 出力件数が既存件数を下回る
4. 既存レコードの内容が 1 件でも変化している
5. 追加件数が `--max-add` (既定 50) を超える (暴走検知)
6. 出力が `legs_server.py` のパーサと非互換

バックアップ作成失敗・書き込みの I/O エラー等の想定外の例外も `❌ legs.json マージが異常終了しました` として Slack へ流れる。

### 手順

```bash
export MERGE_LEGS_WEBHOOK_URL="https://hooks.slack.com/services/XXX/YYY/ZZZ"

# 1. 何が追加されるか確認 (本番は書き換えない)
python3 merge_legs.py --dry-run

# 2. 問題なければ実行
python3 merge_legs.py
```

既定のパスは以下。変える場合は `--target` `--source` `--backup-dir` で指定する。

| 対象 | 既定パス |
|---|---|
| 本番 legs.json | `/mnt/disks/dropoff/global_data/legs.json` |
| 生成物 | `/mnt/disks/dropoff/zp_summary/legs_generated.json` |
| バックアップ | `/mnt/disks/dropoff/backups/global_data/` (30 世代) |

### リバート手順

書き込み後の検証に失敗した場合は**自動でロールバックされる** (Slack に `❌ legs.json マージ失敗・ロールバック済み` が飛ぶ)。取り込んだ内容自体を後から戻したいときは、世代バックアップから復元する。

```bash
# 1. バックアップ世代を確認 (ファイル名の末尾がマージ実行時刻)
ls -lt /mnt/disks/dropoff/backups/global_data/

# 2. 戻す前に現状を退避 (このスクリプトは *.safe を削除しない)
cp /mnt/disks/dropoff/global_data/legs.json \
   /mnt/disks/dropoff/backups/global_data/legs.json.before-revert.safe

# 3. 復元
cp /mnt/disks/dropoff/backups/global_data/legs.json.20260812-120000 \
   /mnt/disks/dropoff/global_data/legs.json

# 4. 件数とパース互換を確認
python3 -c "import json;print(len(json.load(open('/mnt/disks/dropoff/global_data/legs.json'))))"
python3 merge_legs.py --dry-run
```

`--keep-backups` の世代管理は `<ファイル名>.YYYYMMDD-HHMMSS` 形式のみを対象にするため、手動で退避した `*.safe` は自動削除されない。

本番を 1 行 1 レコード形式に整形し直すだけなら `--normalize-only` を使う (マージはしない)。

---

## トラブルシューティング

### Slack API error: missing_scope

Bot Token に必要 scope が無い:

```bash
cd ~/Downloads/zp_summary
TOKEN=$(grep '^SLACK_BOT_TOKEN' .env | cut -d= -f2 | tr -d '"' | tr -d ' ' | tr -d "'" | tr -d '\r')
curl -s -X POST "https://slack.com/api/auth.test" -H "Authorization: Bearer $TOKEN" -i | grep -i "x-oauth-scopes"
```

必要なscope: `channels:history` `channels:read` `groups:history` `groups:read` `users:read` `incoming-webhook`

- public チャンネルは `channels:history` / `channels:read`
- **private チャンネルは `groups:history` / `groups:read` が必須**。`conversations.info` すら `missing_scope` を返す場合は対象が private の可能性が高い

不足があれば Slack App画面で **Bot Token Scopes** に追加 → **reinstall your app** → 新トークンを `.env` に反映。private チャンネルでは Bot がメンバーである必要もある (`/invite @アプリ名`)。

### Slack API error: not_in_channel

Bot がチャンネルに招待されていない:

```
/invite @アプリ名
```

をそのチャンネルで実行。

### Microsoft Graph 403: User does not have access to lookup meeting

会議の主催者ではないため。track_calendars でドライバー予定表を指定し、共有を受けることで解消。

### 「管理者の承認が必要です」画面

`Calendars.Read.Shared` 等が admin consent 必須のテナント設定。

- **Request approval** を押して IT 管理者に依頼
- 承認されるまで Teams 会議情報は取得できない (このツールは Teams 会議情報が前提)

### MSAL トークンキャッシュをリセット

scope を変更した場合や認証エラー時:

```bash
rm ~/.slack_attachment_reader_msal_cache.bin
```

次回実行時にデバイスコード認証が再起動。

### systemd `status=203/EXEC` エラー

`run_zp_summary.sh` が存在しない or 実行権限が無い:

```bash
ls -la ~/Downloads/zp_summary/run_zp_summary.sh
```

実行権限が無い場合:

```bash
chmod +x ~/Downloads/zp_summary/run_zp_summary.sh
```

### JSON 構文エラー

```bash
python3 -m json.tool ~/Downloads/zp_summary/config.json
```

エラー位置 (line/column) が表示される。よくあるミス:
- 末尾カンマ (最後のキーの後に `,` がある)
- カンマ忘れ (キーとキーの間)
- クォート不一致

### DM に開始通知が来ない

1. `config.json` に `start_notify_webhook_url` が記載されているか確認:
   ```bash
   grep "start_notify_webhook_url" ~/Downloads/zp_summary/config.json
   ```
2. URL が生きているか curl でテスト ([Slack通知を変えるとき](#slack通知を変えるとき) 参照)
3. ログを確認:
   ```bash
   journalctl --user -u zp-summary.service --since "10min ago" --no-pager | grep -i "webhook\|notif"
   ```

### 完了通知が `✅ success` だが実際は不完全データ

`---` 等のプレースホルダ文字を「未取得」扱いするロジックは入っているが、独自の表記があれば `is_legs_record_complete` 関数の `PLACEHOLDER_PATTERNS` に追加可能 (slack_attachment_reader.py を編集)。

---

## メンテナンス

### MSAL トークンの90日制限

`Calendars.Read.Shared` 等のリフレッシュトークンは約90日有効。期限切れになると認証エラー。

対策: **月に1回くらい手動で実行** してリフレッシュトークンを更新:

```bash
cd ~/Downloads/zp_summary
python3 slack_attachment_reader.py --config config.json
```

通常通り動けば自動でリフレッシュされる。デバイスコード認証画面が出たら、URLを開いて再認証。

### PC がオフでも動かしたい場合

systemd user services はログイン中のみ動作。PC放置時にも実行したい場合:

```bash
sudo loginctl enable-linger $USER
```

これでログアウト中・スクリーンロック中でもユーザーサービスが走る。

### Persistent=true の効果

`zp-summary.timer` の `Persistent=true` により、PCがオフだった時間帯の実行も、次回起動時に1回だけ取り戻す。

確認:

```bash
grep "Persistent" ~/.config/systemd/user/zp-summary.timer
```

### Git に変更を反映

```bash
cd ~/Downloads/zp_summary
git status                # 変更点確認
git add slack_attachment_reader.py
git commit -m "変更内容"
git push
```

`.gitignore` で除外されているため、`config.json` `.env` `*.bin` 等の秘密情報は誤って push されない。

### 定期的な動作確認

毎日のSlack通知 (`✅ success` / `❌ failed` / `🔁 skipped`) で結果を確認するのが基本。
通知が来ない日があれば:

```bash
# 直近のtimer起動を確認
systemctl --user list-timers --no-pager | grep zp-summary

# 直近の実行ログ
journalctl --user -u zp-summary.service --since "today" --no-pager | tail -50
```

---

## 関連リソース

- [Slack API: Incoming Webhooks](https://api.slack.com/messaging/webhooks)
- [Slack API: OAuth scopes](https://api.slack.com/scopes)
- [Microsoft Graph: Calendar API](https://docs.microsoft.com/en-us/graph/api/resources/calendar)
- [systemd.timer manual](https://www.freedesktop.org/software/systemd/man/systemd.timer.html)
- [systemd calendar event 構文](https://www.freedesktop.org/software/systemd/man/systemd.time.html#Calendar%20Events)

---

## ライセンス

社内ツール (非公開)

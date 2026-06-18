# zp_summary

Slack 運行記録チャンネルから親メッセージを読み取り、Microsoft Teams 会議情報および Excel 添付ファイルを組み合わせて、運行記録(Trackname / 顧客 / 区間 / 時刻 / ドライバー / SW-ver など)を構造化 JSON として出力するツール。

## 機能

- **Slack の親メッセージ取得**: 複数チャンネル対応・日付絞り込み・スレッド取得
- **Microsoft Graph 連携**:
  - Teams 会議情報 (件名 / 開始終了時刻) を取得
  - 自分が招待されていない会議も、ドライバーから共有された予定表 (`Calendars.Read.Shared`) を経由して取得可能
- **Excel パース**: `運行ダイヤ.xlsx` 等から AD開始/AD終了 を抽出 (`tracking_source: "excel"` モード)
- **出力**:
  - `result.json`: ラベル付きビュー (`__labeled_view__`)
  - `legs.json`: 往路/復路ごとに区切られた配列フォーマット (重複スキップ)
  - `logs.json`: 既処理スレッドの投稿日履歴 (次回実行時の高速化に使用)
- **Slack 通知**: Incoming Webhook で完了通知 (各レコードの完全性に基づく `success`/`failed` 表示)
- **増分実行**: `logs.json` の最新投稿日より新しい投稿のみ取得

## セットアップ

### 1. 依存ライブラリ

```bash
pip install slack_sdk requests openpyxl python-docx pandas xlrd msal python-dotenv
```

(Ubuntu 24 等で外部管理 Python の場合は `--break-system-packages` を付与)

### 2. 環境変数

`.env` ファイルを作成 (リポジトリには含めない):

```env
SLACK_BOT_TOKEN=xoxb-...
SLACK_USER_TOKEN=xoxp-...
```

### 3. 設定ファイル

`config.json.example` を参考に `config.json` を作成:

```json
{
  "channel": ["C0123ABC", "C0456DEF"],
  "limit": 200,
  "content_date": "today",
  "extract": "tracking",
  "tracking_source": "teams",
  "track_calendars": [
    "driver-a@yourcompany.com",
    "driver-b@yourcompany.com"
  ],
  "name_pattern": "^運行",
  "out": "result.json",
  "legs_out": "legs.json",
  "logs_out": "logs.json",
  "notify_webhook_url": "https://hooks.slack.com/services/..."
}
```

### 4. Microsoft Graph 認証

初回実行時、デバイスコード認証が起動する:

1. ターミナルに表示された URL を開く
2. 表示されたコードを入力
3. 同意画面で要求スコープを承諾

`tracking_source: "teams"` では以下のスコープが必要:
- `Files.Read.All`
- `OnlineMeetings.Read`
- `Calendars.Read`
- `Calendars.Read.Shared` (`track_calendars` 指定時)

`Calendars.Read.Shared` は管理者承認が必要なテナントもある。

### 5. ドライバーカレンダー共有

`track_calendars` を使う場合は、各ドライバーが Outlook で
**「予定表の共有」→ 「編集可能」または「閲覧可能」** で実行ユーザーに共有しておく必要がある。

## 使い方

### 通常実行 (config経由)

```bash
python slack_attachment_reader.py --config config.json
```

### CLI 引数で実行

```bash
python slack_attachment_reader.py \
  --channel C0123ABC C0456DEF \
  --content-date 2026-06-02 \
  --extract tracking \
  --tracking-source teams \
  --track-calendars driver-a@yourcompany.com driver-b@yourcompany.com \
  --out result.json \
  --legs-out legs.json \
  --logs-out logs.json
```

### Excelモード (Teams API 不要)

```bash
python slack_attachment_reader.py --config config.json
# config.json で "tracking_source": "excel" に変更
```

### 補助スクリプト

`calendar_probe.py`: Microsoft Graph の予定表アクセス可否を確認:

```bash
python calendar_probe.py driver-a@yourcompany.com
```

## 出力フォーマット

### `result.json`
```json
[
  ["GIGA03", "03|2026/06/02", "13:00/17:00", {
    "Driver": "U12345", "operator": "oneman",
    "SW-version": "v1.2.3", "selfdrive section": "東京-大阪",
    "loaded liggage": "XX様", "url": "https://..."
  }],
  ...,
  {"__labeled_view__": {...}}
]
```

### `legs.json`
```json
[
  ["GIGA03", "03|2026/06/02",
   "2026-06-02T13:00:00.000+09:00/2026-06-02T17:00:00.000+09:00",
   {"SW-version": "v1.2.3", "selfdrive_section": "東京-大阪",
    "loaded_luggage": "XX様", "url": "https://..."}],
  ...
]
```

### `logs.json`
```json
[
  {"trackname": "GIGA03",
   "url": "https://...",
   "posted_at": "2026-05-29T14:00:00+09:00"},
  ...
]
```

## アーキテクチャ

### 抽出モード (`--extract`)
- `raw`: 全添付ファイルをそのまま保存
- `driving`: 運転日報用に「自動運転」関連フィールドを抽出
- `tracking`: 運行記録用に時系列レコードを構築 (本ツールの主用途)

### `tracking_source`
- `excel`: Excel の AD開始/AD終了 行を Time-range として使用
- `teams`: Teams 会議の開始終了時刻を Time-range として使用
  - 招待されている会議: `OnlineMeetings.Read`
  - ドライバー予定表から: `Calendars.Read.Shared` 経由

### 増分実行 (`logs.json`)
1. 起動時 `logs.json` を読み、最新投稿日を取得
2. その日付以降の Slack メッセージのみ取得
3. 処理完了後、新規追加レコードの投稿日を追記

## デプロイ例 (cron)

```cron
# 毎朝 7時 に前日分を取得
0 7 * * * cd /path/to/zp_summary && python slack_attachment_reader.py --config config.json --content-date yesterday >> stdout.log 2>> stderr.log
```

## トラブルシューティング

### Teams API が `403 User does not have access to lookup meeting`
- 実行ユーザーが会議の主催者ではないため
- 対処: `track_calendars` でドライバーの予定表を指定し、共有設定を確認

### Microsoft の `管理者の承認が必要です` 画面
- `Calendars.Read.Shared` 等が管理者承認必須のテナント設定
- 対処: 「承認をリクエスト」を押す → IT管理者に連絡

### MSAL トークンキャッシュをリセット
スコープを変更した場合は、再認証のため:
```bash
# Windows
Remove-Item $env:USERPROFILE\.slack_attachment_reader_msal_cache.bin
# Linux / WSL
rm ~/.slack_attachment_reader_msal_cache.bin
```

## ライセンス

社内ツール (非公開)

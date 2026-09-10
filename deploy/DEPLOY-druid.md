# druid への配備手順

`zp_summary` を druid（`t2-integration` / `asia-northeast1-a`）で日次実行し、
生成した運行記録を本番 `legs.json` へ**追記のみ**で反映するための手順。

> このファイルの内容はローカルで検証済みだが、**druid 上ではまだ一切実行していない**。
> 実行者が各ステップの内容を確認したうえで進めること。

---

## 1. 設計方針 — なぜ `global_data` に直接書かないか

```
Slack + Teams
     │
     ▼
slack_attachment_reader.py   (append: true)
     │
     ▼
/mnt/disks/dropoff/zp_summary/legs_generated.json     ← 生成物の置き場 (専用)
     │
     │  merge_legs.py   追記のみ / 差分チェック / 世代バックアップ / アトミック書き込み
     ▼
/mnt/disks/dropoff/global_data/legs.json              ← 本番 (nginx が配信)
     │
     ├─→ nginx :8080 /global_data/legs.json  →  main.js / legs_record_table.js / bug_tickets.html
     └─→ legs_server.py :9090                →  legs エディタ
```

`legs_out` を本番 `legs.json` に直接向けてはいけない。理由は 3 つある。

1. **`append: false` だと全消しになる。** `slack_attachment_reader.py` は
   `final_legs = existing_legs + new_legs` で出力し、`append=false` のとき
   `existing_legs` は空。1 回実行するだけで既存レコードが当日抽出分だけに置き換わる。
2. **メタデータの互換性がない。** `build_legs_record()` が出力するのは
   `SW-version` / `selfdrive_section` / `loaded_luggage` / `url` の 4 つだけ。
   本番の全レコードが持つ `driver` / `operator` / `estimated_load` は出力されないため、
   上書きするとこれらが消える。
3. **`legs_server.py` の `/global_data/` 判定に巻き込まれる。** サブディレクトリを
   切っても前方一致に引っかかる（このバグは修正済みだが、経路自体を分けておくのが安全）。

---

## 2. 事前確認

```bash
# 本番データの現状
jq 'length' /mnt/disks/dropoff/global_data/legs.json          # 期待: 360 前後
md5sum /mnt/disks/dropoff/global_data/legs.json

# 依存パッケージ
python3 -c "import slack_sdk, requests, msal, dotenv; print('ok')"
```

未インストールなら:

```bash
pip3 install slack_sdk requests msal python-dotenv --break-system-packages
```

---

## 3. コードの配置

`legs_tools/zp_summary/` は zero-plotter リポジトリに含まれる。既存のクローンを更新する。

```bash
cd /home/integration-user/zero-plotter
git fetch origin
git status                     # 未コミット変更がないか必ず確認
git checkout <zp_summary を含むブランチ>
git pull
ls legs_tools/zp_summary/      # slack_attachment_reader.py と merge_legs.py があること
chmod +x legs_tools/zp_summary/run_zp_summary.sh legs_tools/zp_summary/merge_legs.py
```

> ⚠️ nginx コンテナ `zero-plotter` が `csv_exported/` をマウントしている。
> `git checkout` / `git pull` はそのディレクトリも書き換えうるので、
> 差分の内容を確認してから実行すること。

---

## 4. 認証情報の配置

**共有 VM に秘密情報を置くことになる。** druid には `sudo` 可能な利用者が複数いるため、
配置してよいか判断したうえで進めること。

```bash
# Slack トークン
install -m 600 -o integration-user -g integration-user /dev/null \
    /home/integration-user/zero-plotter/legs_tools/zp_summary/.env
# → SLACK_BOT_TOKEN / SLACK_USER_TOKEN を記入
#   merge_legs.py の通知を使うなら MERGE_LEGS_WEBHOOK_URL も追記

# Microsoft Graph のトークンキャッシュ
#   ローカルの ~/.slack_attachment_reader_msal_cache.bin をコピーするか、
#   druid 上で対話的にデバイスコード認証を 1 回通す:
cd /home/integration-user/zero-plotter/legs_tools/zp_summary
sudo -u integration-user HOME=/home/integration-user \
    python3 slack_attachment_reader.py --config config.json
chmod 600 /home/integration-user/.slack_attachment_reader_msal_cache.bin
```

> キャッシュの有効期間は約 90 日で、定期実行していれば自動更新される。
> 失効すると systemd 実行時にデバイスコード認証へ進もうとして失敗するため、
> 完了通知が来なくなったらこの再認証を行う。

---

## 5. 設定ファイル

```bash
cd /home/integration-user/zero-plotter/legs_tools/zp_summary
cp deploy/config.druid.json.example config.json
# webhook URL を実値に置き換える
```

`legs_out` / `logs_out` / `append: true` は雛形のまま変更しないこと。

---

## 6. ディレクトリ作成

```bash
sudo mkdir -p /mnt/disks/dropoff/zp_summary /mnt/disks/dropoff/backups/global_data
sudo chown integration-user:integration-user \
    /mnt/disks/dropoff/zp_summary /mnt/disks/dropoff/backups/global_data
sudo chmod 750 /mnt/disks/dropoff/zp_summary /mnt/disks/dropoff/backups/global_data
```

`backups/` は `global_data/` の外に置く。`global_data/` 直下だと nginx が配信してしまう。

---

## 7. 本番 legs.json の整形（初回のみ）

本番には 1 件、複数行に展開されたレコードが混入しており、`legs_server.py` の
行ベースパーサから見えない（360 件中 359 件しか読めない）。
**legs エディタで開いて保存すると、この 1 件が無言で消える。**

```bash
# まず退避
sudo -u integration-user cp -a /mnt/disks/dropoff/global_data/legs.json \
    /mnt/disks/dropoff/backups/global_data/legs.json.$(date +%Y%m%d-%H%M%S).safe

# 差分を確認（書き込まない）
sudo -u integration-user python3 merge_legs.py --normalize-only --dry-run

# 実行
sudo -u integration-user python3 merge_legs.py --normalize-only
```

JSON としてのデータは 1 バイトも変わらず、行の折り方だけが揃う（ローカル検証済み）。
実行後は `legs_server.py` が全件読めるようになる。

---

## 8. 動作確認（本番を書き換えない）

```bash
cd /home/integration-user/zero-plotter/legs_tools/zp_summary

# 収集だけ実行して生成物を確認
sudo -u integration-user HOME=/home/integration-user ./run_zp_summary.sh
jq 'length' /mnt/disks/dropoff/zp_summary/legs_generated.json

# マージを dry-run（何件追加されるか、既存が変化しないか）
sudo -u integration-user python3 merge_legs.py --dry-run
```

`--dry-run` は本番に一切書き込まない。追加件数と追加内容がログに出るので、
妥当であることを確認してから次へ進む。

---

## 9. systemd の設置

```bash
sudo cp deploy/zp-summary.service /etc/systemd/system/
sudo cp deploy/zp-summary.timer   /etc/systemd/system/
sudo systemctl daemon-reload

# まず手動で 1 回流す
sudo systemctl start zp-summary.service
sudo journalctl -u zp-summary.service -n 100 --no-pager

# 問題なければタイマーを有効化
sudo systemctl enable --now zp-summary.timer
systemctl list-timers --no-pager | grep zp-summary
```

実行時刻は `zp-summary.timer` の `OnCalendar=*-*-* 12:00:00` を編集して変更する。

---

## 10. 安全機構のまとめ

`merge_legs.py` は以下のいずれかに該当すると **本番を書き換えずに中断**する（終了コード 1）。

| # | 条件 |
|---|---|
| 1 | 本番 `legs.json` がパースできない |
| 2 | 出力件数が既存件数を下回る |
| 3 | 既存レコードが 1 件でも変化している |
| 4 | 追加件数が `--max-add`（既定 50）を超える |
| 5 | 生成物がパースできない / 空 / 配列でない |

書き込みは「世代バックアップ → 一時ファイル → `os.replace`」の順で、
書き込み後に読み直して検証し、異常ならバックアップから自動ロールバックする。

`legs_server.py` 側にも以下のガードを追加済み。

| リクエスト | 挙動 |
|---|---|
| `GET /global_data/legs.json` | 200（従来どおり） |
| `GET /global_data/<その他>` | 404（従来は legs.json を返していた） |
| `PUT /global_data/<その他>` | 405（従来は legs.json を上書きしていた） |
| `PUT /global_data/legs.json` に空配列 | 409（全消しを拒否） |
| `PUT /global_data/legs.json` に配列以外 | 400 |

---

## 11. ロールバック

```bash
ls -lt /mnt/disks/dropoff/backups/global_data/
sudo -u integration-user cp -a \
    /mnt/disks/dropoff/backups/global_data/legs.json.<タイムスタンプ> \
    /mnt/disks/dropoff/global_data/legs.json
jq 'length' /mnt/disks/dropoff/global_data/legs.json
```

タイマーを止める場合:

```bash
sudo systemctl disable --now zp-summary.timer
```

---

## 12. 運用コマンド

```bash
systemctl list-timers --no-pager | grep zp-summary   # 次回実行予定
sudo systemctl status zp-summary.service             # 直近の実行結果
sudo journalctl -u zp-summary.service -n 200 --no-pager
sudo journalctl -u zp-summary.service --since today | grep -E '中断|警告|ERROR'
```

---

## 13. 未対応の課題

配備とは独立して残っている問題。

- `/global_data/epic_list.json` が存在せず HTTP 404。
  `bug_tickets.html:164` と `geo-plotter.js:5` が参照している。
- `/mnt/disks/dropoff/global_data/` のパーミッションが `777`（`root:root` 所有）。
- nginx コンテナへの `global_data` マウントが `rw`。nginx は書かないので `:ro` でよい。
- `default.conf` の `location /global_data`（末尾スラッシュなし）+ `alias` の組み合わせ。
  `/global_dataXXX` が `/usr/share/nginx/global_dataXXX` に解決される。
- NFS が `/mnt/disks/dropoff` を exporter VM（10.146.0.13）へ `rw,no_root_squash` で
  エクスポートしている。

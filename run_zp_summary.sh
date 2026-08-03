#!/bin/bash
# systemd から呼ばれる起動スクリプト。
# 配置場所を移動しても動くよう、スクリプト自身の位置を基準にする。
set -eu
cd "$(dirname "$(realpath "$0")")"
python3 slack_attachment_reader.py --config config.json

#!/bin/sh
# 安裝或更新 systemd 服務（以 root 執行）。
set -eu
APP=/volume1/docker/image-classification
HERE=$(cd "$(dirname "$0")" && pwd)
mkdir -p "$APP/bin" "$APP/logs"
[ -f "$APP/sd-import.ini" ] || cp "$HERE/sd-import.ini.example" "$APP/sd-import.ini"
cp "$HERE/sd-import.service" /etc/systemd/system/sd-import.service
systemctl daemon-reload
systemctl enable sd-import.service
systemctl restart sd-import.service
systemctl --no-pager status sd-import.service | head -5

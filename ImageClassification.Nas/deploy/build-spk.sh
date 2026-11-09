#!/bin/sh
# 打包 DSM 套件（只放主選單圖示，點了開 :8390 進度頁）。輸出 SdImport.spk。
# 用法：sh build-spk.sh [輸出目錄]
set -eu
HERE=$(cd "$(dirname "$0")" && pwd)
OUT=${1:-$HERE}
W=$(mktemp -d)
trap 'rm -rf "$W"' EXIT
P=$HERE/dsm-package
A=$HERE/../assets

mkdir -p "$W/pkg/ui/images" "$W/spk"
cp "$P/ui-config.json" "$W/pkg/ui/config"
for n in 16 24 32 48 64 72 256; do cp "$A/dsm_$n.png" "$W/pkg/ui/images/sdimport_$n.png"; done
COPYFILE_DISABLE=1 tar -C "$W/pkg" -czf "$W/spk/package.tgz" ui

cp -R "$P/scripts" "$P/conf" "$W/spk/"
cp "$P/INFO" "$W/spk/INFO"
cp "$A/dsm_72.png" "$W/spk/PACKAGE_ICON.PNG"
cp "$A/dsm_256.png" "$W/spk/PACKAGE_ICON_256.PNG"
COPYFILE_DISABLE=1 tar -C "$W/spk" --format=ustar -cf "$OUT/SdImport.spk" INFO PACKAGE_ICON.PNG PACKAGE_ICON_256.PNG package.tgz scripts conf
echo "$OUT/SdImport.spk"

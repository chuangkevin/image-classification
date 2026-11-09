#!/bin/sh
# 把匯入結果備援到另一台機器（例如另一台 NAS）的照片資料夾。
# 主機、帳號、路徑放在 $APP/backup.env（範例見 backup.env.example，不進版控）。
#
# 不用 rsync：DSM 目的地要開「rsync 服務」才允許 rsync over SSH，那會改系統設定
# 並開 873 埠。改成「比對清單 → tar 串流缺的檔案」：
#   1. 列出本機檔案，與目的地同名資料夾底下的檔案比對，只挑目的地沒有的
#   2. tar 串流到目的地的暫存資料夾 .sd-import-incoming
#   3. 全部收完才逐一 mv 到正式位置；目的地已有同路徑檔案就不動（不覆蓋、不刪除）
# 中途斷線只會留下暫存資料夾，下次執行先清掉再重傳，正式資料夾不會有半個檔案。
#
# 由 sd-import 在每次匯入後呼叫（BackupScript），失敗會定時重試。
# DRY_RUN=1 只列出會傳哪些檔案；LIMIT=n 只傳前 n 個（測試用）。
set -eu
APP=${APP:-/volume1/docker/image-classification}
# shellcheck disable=SC1091
. "$APP/backup.env"
SRC=${BACKUP_SRC:-/volume1/photo/ImageClassification}
KEY=${BACKUP_KEY:-$APP/backup_ed25519}
KNOWN=$APP/known_hosts
DEST_USER=$BACKUP_USER
DEST_DIR=$BACKUP_DEST
HOSTS=$BACKUP_HOSTS          # 空白分隔，依序嘗試（例如先區網、再 Tailscale）
STG=".sd-import-incoming"
export LC_ALL=C

SSH="ssh -i $KEY -o UserKnownHostsFile=$KNOWN -o StrictHostKeyChecking=accept-new -o BatchMode=yes -o ConnectTimeout=10 -o ServerAliveInterval=30"
HOST=""
for h in $HOSTS; do
  if $SSH "$DEST_USER@$h" true 2>/dev/null; then HOST=$h; break; fi
  echo "連不到 $h" >&2
done
[ -n "$HOST" ] || exit 255
R="$SSH $DEST_USER@$HOST"

TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT
cd "$SRC"

find . -path '*/@eaDir' -prune -o -type f ! -name '*.part' -printf '%P\n' | sort > "$TMP/local"
ls -1 | grep -v '^@' > "$TMP/tops" || true

# 只列目的地上「本機也有」的品牌資料夾，避免掃整棵樹（舊 NAS 很慢）
$R "cd \"$DEST_DIR\" && while IFS= read -r d; do [ -d \"\$d\" ] && find \"\$d\" -path '*/@eaDir' -prune -o -type f -printf '%p\n'; done; true" \
  < "$TMP/tops" | sort > "$TMP/remote"

comm -23 "$TMP/local" "$TMP/remote" > "$TMP/missing.all"
if [ -n "${LIMIT:-}" ]; then head -n "$LIMIT" "$TMP/missing.all"; else cat "$TMP/missing.all"; fi > "$TMP/missing"
N=$(wc -l < "$TMP/missing")
MB=0
[ "$N" -gt 0 ] && MB=$(tr '\n' '\0' < "$TMP/missing" | xargs -0 stat -c %s | awk '{s+=$1} END {printf "%d", s/1048576}')
echo "本機 $(wc -l < "$TMP/local") 個，目的地已有 $(comm -12 "$TMP/local" "$TMP/remote" | wc -l) 個，要傳 $N 個 ${MB} MB（經 $HOST）"
[ "$N" -gt 0 ] || exit 0

if [ "${DRY_RUN:-0}" = 1 ]; then
  head -20 "$TMP/missing"
  exit 0
fi

tar -cf - --no-recursion --mode='u+rwX,go+rX' -T "$TMP/missing" | $R "
  set -e
  cd \"$DEST_DIR\"
  rm -rf \"$STG\" && mkdir \"$STG\"
  tar -xf - -C \"$STG\"
  cd \"$STG\"
  moved=0; kept=0
  find . -type f > ../$STG.list
  while IFS= read -r f; do
    f=\${f#./}
    if [ -e \"$DEST_DIR/\$f\" ]; then kept=\$((kept+1)); continue; fi
    mkdir -p \"$DEST_DIR/\$(dirname \"\$f\")\"
    mv \"\$f\" \"$DEST_DIR/\$f\"
    moved=\$((moved+1))
  done < ../$STG.list
  cd .. && rm -rf \"$STG\" \"$STG.list\"
  echo \"目的地新增 \$moved 個，已存在未動 \$kept 個\"
"

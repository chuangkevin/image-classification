#!/bin/sh
# 把匯入結果備援到另一台機器（例如另一台 NAS）的照片資料夾。
# 主機、帳號、路徑放在 $APP/backup.env（範例見 backup.env.example，不進版控）。
#
# 不用 rsync：DSM 目的地要開「rsync 服務」才允許 rsync over SSH，那會改系統設定
# 並開 873 埠。改成「比對清單 → tar 串流缺的檔案」：
#   1. 列出本機檔案，與目的地同名資料夾底下的檔案比對，只挑目的地沒有的
#   2. 每 BACKUP_BATCH 個檔案一批，連同「路徑＋大小」清單 tar 串流到目的地的暫存資料夾
#   3. 目的地只搬清單上、而且大小完全相符的檔案；已有同路徑檔案就不動（不覆蓋、不刪除）
# 中斷只會留下暫存資料夾，下次執行先清掉；正式資料夾不會有半個檔案。
# 同路徑但大小不同的檔案只列在日誌（stderr），不自動處理。
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

# 同時只能跑一份（共用目的地暫存資料夾）
exec 9>"$APP/backup.lock"
flock -n 9 || { echo "另一個備援正在執行，略過" >&2; exit 0; }

HOST=""
for h in $HOSTS; do
  if $SSH "$DEST_USER@$h" true 2>/dev/null; then HOST=$h; break; fi
  echo "連不到 $h" >&2
done
[ -n "$HOST" ] || exit 255
R="$SSH $DEST_USER@$HOST"
TAB=$(printf '\t')

TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT
cd "$SRC"

# 路徑<TAB>大小
find . -path '*/@eaDir' -prune -o -type f ! -name '*.part' -printf '%P\t%s\n' | sort > "$TMP/local"
ls -1 | grep -v '^@' > "$TMP/tops" || true

# 只列目的地上「本機也有」的品牌資料夾，避免掃整棵樹（舊 NAS 很慢）
$R "cd \"$DEST_DIR\" && while IFS= read -r d; do [ -d \"\$d\" ] && find \"\$d\" -path '*/@eaDir' -prune -o -type f -printf '%p\t%s\n'; done; true" \
  < "$TMP/tops" | sort > "$TMP/remote"

# missing：目的地沒有這個路徑；mismatch：路徑相同但大小不同（不處理，只回報）
awk -F'\t' 'NR==FNR { r[$1]=$2; next } !($1 in r) { print > "'"$TMP/missing.all"'" } ($1 in r) && r[$1] != $2 { print $1 "\t本機 " $2 "\t目的地 " r[$1] > "'"$TMP/mismatch"'" }' \
  "$TMP/remote" "$TMP/local"
touch "$TMP/missing.all" "$TMP/mismatch"
if [ -n "${LIMIT:-}" ]; then head -n "$LIMIT" "$TMP/missing.all"; else cat "$TMP/missing.all"; fi > "$TMP/missing"

N=$(wc -l < "$TMP/missing")
MB=$(awk -F'\t' '{s+=$2} END {printf "%d", s/1048576}' "$TMP/missing")
BAD=$(wc -l < "$TMP/mismatch")
if [ "$BAD" -gt 0 ]; then
  echo "注意：目的地有 $BAD 個同名但大小不同的檔案（未處理）：" >&2
  head -20 "$TMP/mismatch" >&2
fi
echo "本機 $(wc -l < "$TMP/local") 個，目的地已有 $(( $(wc -l < "$TMP/local") - $(wc -l < "$TMP/missing.all") )) 個，要傳 $N 個 ${MB} MB（經 $HOST）"
[ "$N" -gt 0 ] || exit 0

if [ "${DRY_RUN:-0}" = 1 ]; then
  head -20 "$TMP/missing"
  exit 0
fi

BATCH=${BACKUP_BATCH:-100}
split -l "$BATCH" "$TMP/missing" "$TMP/batch."
total_batches=$(ls "$TMP"/batch.* | wc -l)
RUN=$(date +%s)$$
i=0; moved_all=0; bad_all=0
for b in "$TMP"/batch.*; do
  i=$((i+1))
  cp "$b" "$TMP/sd-import-manifest"
  cut -f1 "$b" > "$TMP/paths"
  # 清單與檔案放在同一個 tar：先清單，再切到來源目錄放檔案
  res=$(tar -cf - --mode='u+rwX,go+rX' -C "$TMP" sd-import-manifest --no-recursion -C "$SRC" -T "$TMP/paths" | $R "
    set -e
    cd \"$DEST_DIR\"
    rm -rf .sd-import-incoming .sd-import-incoming.*
    STG=.sd-import-incoming.$RUN
    mkdir \"\$STG\"
    tar -xf - -C \"\$STG\"
    moved=0; kept=0; bad=0
    while IFS='$TAB' read -r f size; do
      if [ ! -f \"\$STG/\$f\" ] || [ \"\$(stat -c %s \"\$STG/\$f\")\" != \"\$size\" ]; then bad=\$((bad+1)); continue; fi
      if [ -e \"\$f\" ]; then kept=\$((kept+1)); continue; fi
      mkdir -p \"\$(dirname \"\$f\")\"
      mv \"\$STG/\$f\" \"\$f\"
      moved=\$((moved+1))
    done < \"\$STG/sd-import-manifest\"
    rm -rf \"\$STG\"
    echo \"\$moved \$kept \$bad\"
  ")
  set -- $res
  moved_all=$((moved_all + $1)); bad_all=$((bad_all + $3))
  echo "第 $i/$total_batches 批完成，累計新增 $moved_all 個"
done
echo "目的地新增 $moved_all 個（共 $N 個 ${MB} MB）"
[ "$bad_all" -eq 0 ] || { echo "有 $bad_all 個檔案傳送後大小不符，已略過，下次重試" >&2; exit 3; }

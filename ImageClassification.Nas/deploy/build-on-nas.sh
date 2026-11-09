#!/bin/sh
# 在 NAS 上用 .NET SDK 容器編譯成 linux-x64 單一執行檔，輸出到 $APP/bin。
# 換新執行檔後要 systemctl restart sd-import 才會生效。
# 用法：sh build-on-nas.sh   （原始碼需已放在 $APP/src）
set -eu
APP=/volume1/docker/image-classification
rm -rf "$APP/bin.new"
mkdir -p "$APP/bin.new" "$APP/.nuget"
docker run --rm \
  -v "$APP/src:/src" -v "$APP/bin.new:/out" -v "$APP/.nuget:/root/.nuget" \
  -w /src/ImageClassification.Nas \
  mcr.microsoft.com/dotnet/sdk:8.0 \
  dotnet publish -c Release -r linux-x64 --self-contained true \
    -p:PublishSingleFile=true -p:DebugType=None -o /out
# 先建到 bin.new 再換掉：執行中的檔案不能直接覆寫（text file busy），舊行程繼續用舊 inode。
mv "$APP/bin.new/sd-import" "$APP/bin/sd-import.new"
mv -f "$APP/bin/sd-import.new" "$APP/bin/sd-import"
rm -rf "$APP/bin.new"
ls -la "$APP/bin"

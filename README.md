# image-classification
Auto detect storage and classify.

插入記憶卡後，依照片 EXIF 的「品牌 / 型號 / 拍攝日期」自動分類複製。

| 專案 | 執行環境 | 說明 |
|---|---|---|
| `ImageClassification` | Windows 桌面（WinForms） | 原版。偵測新磁碟機，複製到 `config.ini` 的 `OutputPath`。 |
| `ImageClassification.Nas` | Synology DSM（x86_64） | NAS 常駐服務 `sd-import`。讀卡機插在 NAS 上，插卡即自動匯入。 |

## NAS 版（sd-import）

插卡後的流程：

1. 服務每 5 秒檢查一次讀卡機。DSM 只在讀卡機插上 USB 時才掛載卡片，換卡不會，所以有新卡沒被掛載時由服務呼叫 `synousbdisk -mount`。卡片以檔案系統的磁碟區序號辨識，退出後還留在讀卡機裡的同一張卡不會被重掛。看到 `DCIM` 就開始。exFAT 卡需要 DSM 的 exFAT Access 套件。
2. 暫停 `PausePackages`（預設 Synology Photos），匯入時不跟縮圖、人臉辨識搶硬碟。原本就關著的不會被打開。
3. Telegram 通知「開始匯入 N 個檔案，請先不要拔卡」。
4. 複製到 `{OutputPath}/{品牌}/{型號}/{yyyy-MM-dd}/`。
5. 全部成功才自動退出卡片，Telegram 通知結果與「可以拔卡」。
6. 備援丟到背景執行，可以馬上插下一張卡。
7. 最後一張卡處理完、閒置 `ResumeAfterIdleMinutes` 分鐘（預設 10）沒有新卡，才恢復暫停的套件。連續換卡不會反覆開關。
8. 通知 Synology Photos 登錄新照片（`PhotosIndexCommand`）。服務直接寫檔不會觸發 DSM 的檔案事件，不登錄的話 Photos 看不到新照片、也不會做縮圖。

和 Windows 版的差異：

- **不會重複匯入**：目的地已有同名、同大小、同修改時間的檔案就略過，同一張卡再插只會補新照片。
- **同型號多台機身不會漏檔**：同名同大小但時間不同時，比對 SHA-256 才算同一個檔案，否則另存 `_1`、`_2`。
- **不覆蓋**：同名但內容不同一律改名。
- **拔卡不留半成品**：先寫 `.part`，完成才改名。中斷後重插會從中斷處繼續。
- **日期修正**：照片優先用 DateTimeOriginal，影片用 QuickTime 建立時間。原版會被 DateTimeDigitized 蓋掉。
- **影片歸類**：影片沒有相機資訊時，沿用同檔名照片，否則沿用前一張照片的相機。
- **可選刪卡**：`DeleteAfterCopy = true` 時，SHA-256 驗證兩邊一致才刪卡上原檔。預設關閉。

### 看進度

- **進度網頁**：`http://<NAS>:8390`，每秒更新目前的卡、速度、剩餘時間、背景備援、最近處理的卡與日誌。歷史存在 `history.json`，服務重啟不會清空。
- **DSM 主選單圖示**：`deploy/build-spk.sh` 打包 `SdImport.spk`，以 root 執行 `synopkg install SdImport.spk` 後，DSM 主選單會多一個「記憶卡匯入」，點了開進度網頁。套件只有捷徑，不含服務本身。
- **手機 App**：進度網頁支援「加入主畫面」。iPhone Safari 會變成全螢幕獨立 App；Android Chrome 在 http 下只會建立捷徑，要完整安裝需要 https。

### 通知

Telegram bot。token 與 chat id 在每次發送時從 `TelegramEnvFile` 讀取（`TELEGRAM_BOT_TOKEN`、`TELEGRAM_HOME_CHANNEL`），不另存副本。DSM 7 的 `synodsmnotify` 標題與內容都只接受內建 i18n 字串，送不了自訂文字，所以不用它。

### 備援

`deploy/backup-push.sh` 把匯入結果推到另一台機器。目的地主機、帳號、路徑放在 NAS 上的 `backup.env`，範例見 `deploy/backup.env.example`，不進版控。

- **背景執行**：匯入完成只送出請求，不擋下一張卡。執行中又有新請求，跑完會再補一輪。
- **只新增**：比對兩邊的「路徑＋大小」，只傳目的地沒有的檔案。不覆蓋、不刪除。同名但大小不同的只寫進日誌，不自動處理。
- **分批**：每 `BACKUP_BATCH`（預設 100）個檔案一批，每批附上「路徑＋大小」清單；目的地只搬清單上、大小完全相符的檔案。中斷最多損失一批。
- **不用 rsync**：DSM 目的地要開「rsync 服務」才允許 rsync over SSH，會改系統設定並開 873 埠。改用 tar 串流。
- **斷線安全**：每批先收進目的地的暫存資料夾 `.sd-import-incoming.*`，收完才搬到正式位置。下次執行先清掉中斷留下的暫存。`flock` 保證同時只跑一份。
- **重試**：失敗每 `BackupRetryMinutes` 分鐘重試，第一次失敗與補完各通知一次。服務啟動時也會先補跑一次。
- **金鑰**：建議用專用金鑰，目的地 `authorized_keys` 以 `from=` 限定來源並加 `no-pty`。

手動演練（只列出會傳哪些檔案）：

```bash
DRY_RUN=1 sh /volume1/docker/image-classification/src/ImageClassification.Nas/deploy/backup-push.sh
```

### 為什麼不用 Docker

卡片插入時 DSM 才建立 `/volumeUSB{n}`，容器只能掛已存在的路徑。改掛 `/` 加 rslave 也不行，因為 DSM 的根目錄是 private mount，要每次開機先 `mount --make-rshared /`。退出卡片的 `synousbdisk`、暫停套件的 `synopkg` 也只在主機上，容器要開 privileged＋pid host 才能用。所以編譯用容器，執行用原生 systemd 服務。

### 部署

NAS 上的位置：

```
/volume1/docker/image-classification/
  src/                原始碼
  bin/sd-import       編譯結果（單一執行檔）
  sd-import.ini       設定（第一次從 deploy/sd-import.ini.example 複製）
  backup.env          備援目的地（不進版控）
  logs/sd-import.log  日誌
```

以 root 在 NAS 上執行：

```bash
sh /volume1/docker/image-classification/src/ImageClassification.Nas/deploy/build-on-nas.sh
```

```bash
sh /volume1/docker/image-classification/src/ImageClassification.Nas/deploy/install.sh
```

手動試跑某張卡（`--dry-run` 只列出不複製）：

```bash
/volume1/docker/image-classification/bin/sd-import import /volumeUSB1/usbshare --dry-run --config /volume1/docker/image-classification/sd-import.ini
```

測試通知：

```bash
/volume1/docker/image-classification/bin/sd-import notify-test --config /volume1/docker/image-classification/sd-import.ini
```

看狀態與日誌：

```bash
systemctl status sd-import
```

```bash
tail -f /volume1/docker/image-classification/logs/sd-import.log
```

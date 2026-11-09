using IniParser;
using IniParser.Model;

namespace ImageClassification.Nas;

/// <summary>sd-import.ini 的設定。檔案不存在時寫出預設值。</summary>
public sealed class Settings
{
    public string OutputPath { get; init; } = "/volume1/photo/ImageClassification";
    /// <summary>以 glob 樣式比對的卡片掛載點（DSM 會掛在 /volumeUSB{n}/usbshare 或 usbshare{n}）。</summary>
    public string MountGlob { get; init; } = "/volumeUSB*/usbshare*";
    public int ScanIntervalSeconds { get; init; } = 5;
    /// <summary>複製並以 SHA-256 驗證後刪除卡上原檔。預設關閉。</summary>
    public bool DeleteAfterCopy { get; init; }
    /// <summary>完成後自動退出卡片，可直接拔卡。</summary>
    public bool EjectWhenDone { get; init; } = true;
    /// <summary>以 Telegram 告知開始與完成。</summary>
    public bool Notify { get; init; } = true;
    /// <summary>含 TELEGRAM_BOT_TOKEN、TELEGRAM_HOME_CHANNEL 的 env 檔。</summary>
    public string TelegramEnvFile { get; init; } = "";
    public string LogPath { get; init; } = "";
    /// <summary>匯入後執行的備援腳本（空白＝不備援）。非 0 結束視為失敗，定時重試。</summary>
    public string BackupScript { get; init; } = "";
    public int BackupRetryMinutes { get; init; } = 60;
    /// <summary>匯入期間暫停的 DSM 套件（逗號分隔，空白＝不暫停）。</summary>
    public string PausePackages { get; init; } = "SynologyPhotos";
    /// <summary>最後一張卡處理完後閒置幾分鐘才恢復暫停的套件。</summary>
    public int ResumeAfterIdleMinutes { get; init; } = 10;
    /// <summary>
    /// 匯入後讓 Synology Photos 登錄新檔的指令（空白＝不登錄）。直接寫檔不會觸發 DSM 的檔案事件，
    /// 不跑這個 Photos 看不到新照片、也不會做縮圖。{0} 會換成 OutputPath。
    /// </summary>
    public string PhotosIndexCommand { get; init; } =
        "/var/packages/SynologyPhotos/target/usr/bin/synofoto-bin-index-tool -t basic -i {0}";
    /// <summary>進度網頁的 HttpListener prefix（空白＝不開）。</summary>
    public string WebPrefix { get; init; } = "http://*:8390/";

    public static Settings Load(string path)
    {
        var parser = new FileIniDataParser();
        if (!File.Exists(path))
        {
            var d = new Settings();
            var data = new IniData();
            var s = data["Settings"];
            s["OutputPath"] = d.OutputPath;
            s["MountGlob"] = d.MountGlob;
            s["ScanIntervalSeconds"] = d.ScanIntervalSeconds.ToString();
            s["DeleteAfterCopy"] = d.DeleteAfterCopy ? "true" : "false";
            s["EjectWhenDone"] = d.EjectWhenDone ? "true" : "false";
            s["Notify"] = d.Notify ? "true" : "false";
            s["TelegramEnvFile"] = d.TelegramEnvFile;
            s["LogPath"] = d.LogPath;
            s["BackupScript"] = d.BackupScript;
            s["BackupRetryMinutes"] = d.BackupRetryMinutes.ToString();
            s["PausePackages"] = d.PausePackages;
            s["ResumeAfterIdleMinutes"] = d.ResumeAfterIdleMinutes.ToString();
            s["WebPrefix"] = d.WebPrefix;
            s["PhotosIndexCommand"] = d.PhotosIndexCommand;
            parser.WriteFile(path, data);
            return d;
        }

        var ini = parser.ReadFile(path)["Settings"];
        string Get(string key, string fallback) =>
            string.IsNullOrWhiteSpace(ini[key]) ? fallback : ini[key].Trim();
        bool GetBool(string key, bool fallback) =>
            ini[key] is { } v && !string.IsNullOrWhiteSpace(v)
                ? v.Trim().ToLowerInvariant() is "true" or "1" or "yes"
                : fallback;

        var def = new Settings();
        return new Settings
        {
            OutputPath = Get("OutputPath", def.OutputPath),
            MountGlob = Get("MountGlob", def.MountGlob),
            ScanIntervalSeconds = int.TryParse(ini["ScanIntervalSeconds"], out var n) && n > 0 ? n : def.ScanIntervalSeconds,
            DeleteAfterCopy = GetBool("DeleteAfterCopy", def.DeleteAfterCopy),
            EjectWhenDone = GetBool("EjectWhenDone", def.EjectWhenDone),
            Notify = GetBool("Notify", def.Notify),
            TelegramEnvFile = Get("TelegramEnvFile", def.TelegramEnvFile),
            LogPath = ini["LogPath"]?.Trim() ?? "",
            BackupScript = ini["BackupScript"]?.Trim() ?? "",
            BackupRetryMinutes = int.TryParse(ini["BackupRetryMinutes"], out var m) && m > 0 ? m : def.BackupRetryMinutes,
            PausePackages = ini["PausePackages"] is { } pp ? pp.Trim() : def.PausePackages,
            ResumeAfterIdleMinutes = int.TryParse(ini["ResumeAfterIdleMinutes"], out var r) && r > 0 ? r : def.ResumeAfterIdleMinutes,
            WebPrefix = ini["WebPrefix"] is { } wp ? wp.Trim() : def.WebPrefix,
            PhotosIndexCommand = ini["PhotosIndexCommand"] is { } pic ? pic.Trim() : def.PhotosIndexCommand,
        };
    }
}

using System.Text.RegularExpressions;

namespace ImageClassification.Nas;

/// <summary>
/// 輪詢 /proc/mounts，找到符合 MountGlob 且含 DCIM 的新掛載點就匯入。
/// 同一個掛載（裝置＋掛載點）只處理一次，拔掉後再插才會重新處理。
/// </summary>
public sealed class CardWatcher
{
    private readonly Settings _settings;
    private readonly Regex _mountPattern;
    private readonly HashSet<string> _handled = new();
    private readonly Notifier _notifier;
    private readonly Backup _backup;
    private readonly PreviewPause _pause;
    private readonly AutoMounter _mounter;

    public CardWatcher(Settings settings, string appDir)
    {
        _settings = settings;
        _notifier = new Notifier(settings);
        _backup = new Backup(settings, _notifier);
        _pause = new PreviewPause(settings, appDir);
        Status.Instance.LoadHistory(Path.Combine(appDir, "history.json"));
        _mounter = new AutoMounter(msg => _notifier.Send(msg));
        _mountPattern = GlobToRegex(settings.MountGlob);
    }

    public void Run(CancellationToken ct)
    {
        Log.Info($"監看 {_settings.MountGlob}，每 {_settings.ScanIntervalSeconds} 秒檢查一次；輸出 {_settings.OutputPath}");
        _backup.Start(ct);
        WebUi.Start(_settings.WebPrefix, ct);
        while (!ct.IsCancellationRequested)
        {
            try
            {
                Tick(ct);
            }
            catch (OperationCanceledException) { break; }
            catch (Exception ex)
            {
                Log.Error($"監看迴圈錯誤: {ex}");
            }
            ct.WaitHandle.WaitOne(TimeSpan.FromSeconds(_settings.ScanIntervalSeconds));
        }
        Log.Info("服務停止。");
    }

    private void Tick(CancellationToken ct)
    {
        var all = ReadMounts().ToList();
        _mounter.Tick(all.Select(m => m.Device).ToHashSet());
        var mounts = all.Where(m => _mountPattern.IsMatch(m.MountPoint)).ToList();
        var present = mounts.Select(m => m.Key).ToHashSet();
        _handled.RemoveWhere(k => !present.Contains(k));

        foreach (var m in mounts)
        {
            if (ct.IsCancellationRequested) return;
            if (_handled.Contains(m.Key)) continue;
            _handled.Add(m.Key);

            if (Importer.FindDcim(m.MountPoint) is null)
            {
                Log.Info($"{m.MountPoint}（{m.Device}）沒有 DCIM，略過。");
                continue;
            }
            _pause.BeforeImport();
            try { HandleCard(m, ct); }
            finally { _pause.AfterImport(); }
        }

        _pause.ResumeIfIdle();
    }

    private void HandleCard(Mount m, CancellationToken ct)
    {
        var label = $"{m.MountPoint}（{m.Device}）";
        Log.Info($"偵測到記憶卡 {label}");
        var count = Importer.ListMedia(Importer.FindDcim(m.MountPoint)!).Count;
        _notifier.Send($"📷 偵測到記憶卡，開始匯入 {count} 個檔案，請先不要拔卡。");

        var started = DateTime.Now;
        ImportResult r;
        try
        {
            Status.Instance.CardScanning(label);
            r = new Importer(_settings, status: Status.Instance).Import(m.MountPoint, ct);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            Log.Error($"{label} 匯入失敗: {ex}");
            Status.Instance.CardFinished("失敗：" + ex.Message, false);
            _notifier.Send($"❌ 記憶卡匯入失敗：{ex.Message}");
            return;
        }

        var secs = Math.Max(1, (DateTime.Now - started).TotalSeconds);
        var summary = $"新增 {r.Copied} 個（{r.BytesCopied / 1048576.0:F0} MB，{r.BytesCopied / 1048576.0 / secs:F1} MB/s），" +
                      $"已存在略過 {r.Skipped} 個，失敗 {r.Failed} 個" +
                      (_settings.DeleteAfterCopy ? $"，已刪卡上 {r.Deleted} 個" : "");
        Log.Info($"{label} {(r.Aborted ? "中止" : "完成")}：{summary}");
        foreach (var f in r.Folders) Log.Info($"  資料夾 {f}");

        var ejected = false;
        if (!r.Aborted && r.Failed == 0 && _settings.EjectWhenDone)
        {
            Status.Instance.CardEjecting();
            ejected = Dsm.Eject(m.Disk);
            Log.Info(ejected ? $"已退出 {m.Disk}，可以拔卡。" : $"退出 {m.Disk} 失敗，請從 DSM 手動退出。");
        }

        var head = r.Aborted ? "⚠️ 記憶卡匯入中止" : r.Failed > 0 ? "⚠️ 記憶卡匯入完成，但有失敗" : "✅ 記憶卡匯入完成";
        var tail = ejected ? "已退出，可以拔卡。" : r.Aborted ? "未退出卡片，重插會從中斷處繼續。" : "未自動退出，請從 DSM 退出後再拔卡。";
        var dates = r.Folders.Count == 0 ? "" : $"\n日期：{string.Join("、", r.Folders.Select(Path.GetFileName).Distinct().Take(10))}{(r.Folders.Count > 10 ? " 等" : "")}";
        _notifier.Send($"{head}\n{summary}{dates}\n{tail}");
        Status.Instance.CardFinished(r.Aborted ? "中止" : r.Failed > 0 ? "完成，有失敗" : "完成", ejected);

        // 備援丟到背景：馬上可以插下一張卡。中止的也要備援已複製的部分。
        if (r.Copied > 0) _backup.Request();
    }

    public sealed record Mount(string Device, string MountPoint)
    {
        public string Key => Device + "|" + MountPoint;
        /// <summary>/dev/sdr1 → sdr（synousbdisk 要的是整顆磁碟名）。</summary>
        public string Disk => Regex.Replace(Path.GetFileName(Device), @"p?\d+$", "");
    }

    public static IEnumerable<Mount> ReadMounts()
    {
        foreach (var line in File.ReadLines("/proc/mounts"))
        {
            var parts = line.Split(' ');
            if (parts.Length < 2 || !parts[0].StartsWith("/dev/")) continue;
            yield return new Mount(parts[0], Unescape(parts[1]));
        }
    }

    // /proc/mounts 會把空白等字元寫成 \040 這種八進位跳脫。
    private static string Unescape(string s) =>
        Regex.Replace(s, @"\\([0-7]{3})", m => ((char)Convert.ToInt32(m.Groups[1].Value, 8)).ToString());

    private static Regex GlobToRegex(string glob) =>
        new("^" + Regex.Escape(glob).Replace(@"\*", "[^/]*").Replace(@"\?", "[^/]") + "$");
}

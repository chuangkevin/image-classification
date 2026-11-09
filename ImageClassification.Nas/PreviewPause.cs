using System.Diagnostics;

namespace ImageClassification.Nas;

/// <summary>
/// 匯入期間暫停會搶硬碟的 DSM 套件（預設 Synology Photos：縮圖、人臉、物件辨識）。
/// 插卡時若套件開著就停掉並記在 marker 檔；最後一張卡處理完、閒置 ResumeAfterIdleMinutes
/// 分鐘都沒新卡才恢復，連續換卡不會反覆開關。原本就關著的套件不會被打開。
/// marker 檔讓服務重啟後仍記得要恢復哪些套件。
/// </summary>
public sealed class PreviewPause
{
    private const string SynoPkg = "/usr/syno/bin/synopkg";
    private readonly Settings _settings;
    private readonly string _marker;
    private readonly HashSet<string> _pausedByUs;
    private DateTime _lastActivity = DateTime.Now;

    public PreviewPause(Settings settings, string appDir)
    {
        _settings = settings;
        _marker = Path.Combine(appDir, "paused-packages");
        _pausedByUs = File.Exists(_marker)
            ? File.ReadAllLines(_marker).Select(l => l.Trim()).Where(l => l.Length > 0).ToHashSet()
            : new HashSet<string>();
        if (_pausedByUs.Count > 0) Log.Info($"上次暫停、待恢復的套件：{string.Join(", ", _pausedByUs)}");
        Status.Instance.SetPaused(_pausedByUs);
    }

    private IEnumerable<string> Packages =>
        _settings.PausePackages.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    public void BeforeImport()
    {
        _lastActivity = DateTime.Now;
        if (!File.Exists(SynoPkg)) return;
        foreach (var pkg in Packages)
        {
            if (_pausedByUs.Contains(pkg) || Run("is_onoff", pkg) != 0) continue; // 已由我們暫停，或本來就關著
            Log.Info($"匯入期間暫停套件 {pkg}");
            if (Run("stop", pkg) == 0)
            {
                _pausedByUs.Add(pkg);
                Save();
            }
        }
    }

    public void AfterImport() => _lastActivity = DateTime.Now;

    /// <summary>監看迴圈每輪呼叫：閒置夠久就恢復。</summary>
    public void ResumeIfIdle()
    {
        if (_pausedByUs.Count == 0) return;
        if (DateTime.Now - _lastActivity < TimeSpan.FromMinutes(_settings.ResumeAfterIdleMinutes)) return;
        foreach (var pkg in _pausedByUs.ToList())
        {
            Log.Info($"閒置 {_settings.ResumeAfterIdleMinutes} 分鐘沒有新卡，恢復套件 {pkg}");
            if (Run("start", pkg) == 0) _pausedByUs.Remove(pkg);
        }
        Save();
        // 啟動失敗就等下一個閒置週期再試，不要每 5 秒狂打。
        _lastActivity = DateTime.Now;
    }

    private void Save()
    {
        Status.Instance.SetPaused(_pausedByUs);
        try
        {
            if (_pausedByUs.Count == 0) File.Delete(_marker);
            else File.WriteAllLines(_marker, _pausedByUs);
        }
        catch (Exception ex)
        {
            Log.Warn($"無法寫入 {_marker}: {ex.Message}");
        }
    }

    private static int Run(string verb, string pkg)
    {
        try
        {
            var psi = new ProcessStartInfo(SynoPkg) { RedirectStandardOutput = true, RedirectStandardError = true };
            psi.ArgumentList.Add(verb);
            psi.ArgumentList.Add(pkg);
            using var p = Process.Start(psi)!;
            p.StandardOutput.ReadToEnd();
            p.StandardError.ReadToEnd();
            if (!p.WaitForExit(180_000)) { p.Kill(); return -1; }
            return p.ExitCode;
        }
        catch (Exception ex)
        {
            Log.Warn($"synopkg {verb} {pkg} 失敗: {ex.Message}");
            return -1;
        }
    }
}

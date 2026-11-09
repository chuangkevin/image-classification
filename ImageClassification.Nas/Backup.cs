using System.Diagnostics;

namespace ImageClassification.Nas;

/// <summary>
/// 背景執行 BackupScript（例如推到另一台 NAS），不擋下一張卡的匯入。
/// - Request()：匯入完成後呼叫，立即返回。
/// - 單一工作執行緒；執行中又收到請求，就在這輪結束後再補跑一輪（新檔一個都不漏）。
/// - 失敗每 BackupRetryMinutes 分鐘重試，第一次失敗與補完各通知一次。
/// - 服務啟動時先補跑一次（腳本只新增，重跑無害）。
/// </summary>
public sealed class Backup
{
    private readonly Settings _settings;
    private readonly Notifier _notifier;
    private readonly AutoResetEvent _signal = new(false);
    private readonly object _gate = new();
    private bool _requested;
    private bool _failing;
    private DateTime _nextRetry = DateTime.MaxValue;
    private Thread? _worker;

    public Backup(Settings settings, Notifier notifier)
    {
        _settings = settings;
        _notifier = notifier;
    }

    public bool Enabled => !string.IsNullOrWhiteSpace(_settings.BackupScript);

    /// <summary>背景是否正在跑或排著要跑（給狀態顯示用）。</summary>
    public bool Busy { get; private set; }

    public void Start(CancellationToken ct)
    {
        if (!Enabled) return;
        _worker = new Thread(() => Loop(ct)) { IsBackground = true, Name = "backup" };
        _worker.Start();
        Request();
    }

    public void Request()
    {
        if (!Enabled) return;
        lock (_gate) _requested = true;
        _signal.Set();
    }

    private void Loop(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            var wait = _nextRetry == DateTime.MaxValue
                ? Timeout.InfiniteTimeSpan
                : TimeSpan.FromMilliseconds(Math.Max(0, (_nextRetry - DateTime.Now).TotalMilliseconds));
            WaitHandle.WaitAny(new[] { _signal, ct.WaitHandle }, wait);
            if (ct.IsCancellationRequested) return;

            lock (_gate)
            {
                var due = _nextRetry != DateTime.MaxValue && DateTime.Now >= _nextRetry;
                if (!_requested && !due) continue;
                _requested = false;
            }

            Busy = true;
            var ok = RunOnce(ct);
            Busy = false;
            if (ct.IsCancellationRequested) return;

            if (ok)
            {
                _nextRetry = DateTime.MaxValue;
                if (_failing) _notifier.Send("✅ 先前失敗的備援已補完。");
                _failing = false;
                // 這輪執行中又有新請求：_requested 仍為 true，下一圈立刻再跑。
                lock (_gate) if (_requested) _signal.Set();
            }
            else
            {
                _nextRetry = DateTime.Now.AddMinutes(_settings.BackupRetryMinutes);
                if (!_failing) _notifier.Send($"⚠️ 備援失敗，會每 {_settings.BackupRetryMinutes} 分鐘重試，詳情看日誌。");
                _failing = true;
            }
        }
    }

    private bool RunOnce(CancellationToken ct)
    {
        Log.Info($"背景備援開始 {_settings.BackupScript}");
        Status.Instance.BackupStarted();
        var started = DateTime.Now;
        string? lastLine = null;
        int code;
        try
        {
            // nice 19：讓備援讓出 CPU 給下一張卡的匯入。DSM 沒有 ionice。
            var psi = new ProcessStartInfo("/usr/bin/nice") { RedirectStandardOutput = true, RedirectStandardError = true };
            foreach (var a in new[] { "-n", "19", "/bin/sh", _settings.BackupScript }) psi.ArgumentList.Add(a);
            using var p = Process.Start(psi)!;
            p.OutputDataReceived += (_, e) =>
            {
                if (string.IsNullOrWhiteSpace(e.Data)) return;
                lastLine = e.Data;
                Status.Instance.BackupLine(e.Data);
                Log.Info($"  backup: {e.Data}");
            };
            p.ErrorDataReceived += (_, e) => { if (!string.IsNullOrWhiteSpace(e.Data)) Log.Warn($"  backup: {e.Data}"); };
            p.BeginOutputReadLine();
            p.BeginErrorReadLine();
            using (ct.Register(() => { try { p.Kill(true); } catch { /* 已結束 */ } }))
            {
                p.WaitForExit();
            }
            code = p.ExitCode;
        }
        catch (Exception ex)
        {
            Log.Error($"備援無法執行: {ex.Message}");
            Status.Instance.BackupFinished(false, DateTime.Now.AddMinutes(_settings.BackupRetryMinutes));
            return false;
        }

        var secs = (DateTime.Now - started).TotalSeconds;
        if (code != 0)
        {
            Log.Warn($"背景備援失敗，結束碼 {code}");
            Status.Instance.BackupFinished(false, DateTime.Now.AddMinutes(_settings.BackupRetryMinutes));
            return false;
        }
        Status.Instance.BackupFinished(true, null);

        Log.Info($"背景備援完成（{secs:F0} 秒）");
        // 只有真的傳了東西才通知；腳本最後一行是「新增 N 個…」，沒東西要傳時是「要傳 0 個」。
        if (lastLine is not null && !lastLine.Contains("要傳 0 個 "))
        {
            _notifier.Send($"☁️ 備援完成：{lastLine}（{secs / 60:F0} 分鐘）");
        }
        return true;
    }
}

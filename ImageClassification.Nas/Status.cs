namespace ImageClassification.Nas;

/// <summary>給進度網頁看的即時狀態。所有存取都經過 lock，網頁每秒拿一次快照。</summary>
public sealed class Status
{
    public static readonly Status Instance = new();
    private readonly object _gate = new();

    // 目前這張卡
    private string _state = "idle";          // idle | scanning | copying | ejecting
    private string? _card;
    private DateTime? _cardStarted;
    private int _total, _done, _copied, _skipped, _failed;
    private long _totalBytes, _doneBytes, _copiedBytes;
    private string? _currentFile;
    private List<string> _cameras = new();

    // 背景備援
    private bool _backupBusy;
    private DateTime? _backupStarted;
    private string? _backupPlan, _backupLast;
    private DateTime? _backupFinished;
    private bool _backupFailing;
    private DateTime? _backupNextRetry;

    private string[] _paused = Array.Empty<string>();
    private readonly LinkedList<System.Text.Json.Nodes.JsonNode> _history = new();
    private string? _historyFile;

    /// <summary>從檔案載入歷史紀錄，之後每張卡完成都寫回，服務重啟不會清空。</summary>
    public void LoadHistory(string file)
    {
        lock (_gate)
        {
            _historyFile = file;
            try
            {
                if (!File.Exists(file)) return;
                var arr = System.Text.Json.Nodes.JsonNode.Parse(File.ReadAllText(file))?.AsArray();
                if (arr is null) return;
                foreach (var n in arr) if (n is not null) _history.AddLast(n.DeepClone());
            }
            catch (Exception ex)
            {
                Log.Warn($"讀取歷史紀錄失敗: {ex.Message}");
            }
        }
    }

    public void CardScanning(string card)
    {
        lock (_gate)
        {
            _state = "scanning"; _card = card; _cardStarted = DateTime.Now;
            _total = _done = _copied = _skipped = _failed = 0;
            _totalBytes = _doneBytes = _copiedBytes = 0;
            _currentFile = null; _cameras = new();
        }
    }

    public void CardPlanned(int files, long bytes, IEnumerable<string> cameras)
    {
        lock (_gate) { _state = "copying"; _total = files; _totalBytes = bytes; _cameras = cameras.ToList(); }
    }

    public void FileStarted(string name) { lock (_gate) _currentFile = name; }

    public void FileDone(long size, bool copied, bool failed)
    {
        lock (_gate)
        {
            _done++; _doneBytes += size;
            if (failed) _failed++;
            else if (copied) { _copied++; _copiedBytes += size; }
            else _skipped++;
        }
    }

    public void CardEjecting() { lock (_gate) { _state = "ejecting"; _currentFile = null; } }

    public void CardFinished(string result, bool ejected)
    {
        lock (_gate)
        {
            _history.AddFirst(System.Text.Json.JsonSerializer.SerializeToNode(new
            {
                at = DateTime.Now,
                card = _card,
                cameras = _cameras,
                result,
                ejected,
                copied = _copied,
                skipped = _skipped,
                failed = _failed,
                mb = Math.Round(_copiedBytes / 1048576.0),
                minutes = _cardStarted is null ? 0 : Math.Round((DateTime.Now - _cardStarted.Value).TotalMinutes, 1),
            })!);
            while (_history.Count > 30) _history.RemoveLast();
            if (_historyFile is not null)
            {
                try
                {
                    var arr = new System.Text.Json.Nodes.JsonArray(_history.Select(n => (System.Text.Json.Nodes.JsonNode?)n.DeepClone()).ToArray());
                    File.WriteAllText(_historyFile + ".tmp", arr.ToJsonString());
                    File.Move(_historyFile + ".tmp", _historyFile, true);
                }
                catch (Exception ex)
                {
                    Log.Warn($"寫入歷史紀錄失敗: {ex.Message}");
                }
            }
            _state = "idle"; _card = null; _currentFile = null;
        }
    }

    public void BackupStarted() { lock (_gate) { _backupBusy = true; _backupStarted = DateTime.Now; _backupPlan = null; } }
    public void BackupLine(string line) { lock (_gate) { _backupPlan ??= line; _backupLast = line; } }

    public void BackupFinished(bool ok, DateTime? nextRetry)
    {
        lock (_gate)
        {
            _backupBusy = false; _backupFinished = DateTime.Now;
            _backupFailing = !ok; _backupNextRetry = ok ? null : nextRetry;
        }
    }

    public void SetPaused(IEnumerable<string> packages) { lock (_gate) _paused = packages.ToArray(); }

    public object Snapshot()
    {
        lock (_gate)
        {
            double? speed = null, eta = null;
            if (_cardStarted is { } st && _state == "copying")
            {
                var secs = (DateTime.Now - st).TotalSeconds;
                if (secs > 3 && _copiedBytes > 0)
                {
                    speed = _copiedBytes / 1048576.0 / secs;
                    eta = (_totalBytes - _doneBytes) / 1048576.0 / speed;
                }
            }
            return new
            {
                now = DateTime.Now,
                host = Environment.MachineName,
                card = new
                {
                    state = _state, label = _card, started = _cardStarted, cameras = _cameras,
                    total = _total, done = _done, copied = _copied, skipped = _skipped, failed = _failed,
                    totalMb = Math.Round(_totalBytes / 1048576.0), doneMb = Math.Round(_doneBytes / 1048576.0),
                    current = _currentFile,
                    speedMbps = speed is null ? (double?)null : Math.Round(speed.Value, 1),
                    etaSeconds = eta is null ? (double?)null : Math.Round(eta.Value),
                },
                backup = new
                {
                    busy = _backupBusy, started = _backupStarted, plan = _backupPlan, last = _backupLast,
                    finished = _backupFinished, failing = _backupFailing, nextRetry = _backupNextRetry,
                },
                paused = _paused,
                history = new System.Text.Json.Nodes.JsonArray(_history.Select(n => (System.Text.Json.Nodes.JsonNode?)n.DeepClone()).ToArray()),
                log = Log.Recent(),
            };
        }
    }
}

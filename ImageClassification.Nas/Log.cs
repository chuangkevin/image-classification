namespace ImageClassification.Nas;

/// <summary>同時寫 stdout（journald）與可選的檔案。</summary>
public static class Log
{
    private static readonly object Gate = new();
    private static string? _file;
    private static readonly LinkedList<string> RecentLines = new();

    /// <summary>最近 60 行，給進度網頁。</summary>
    public static string[] Recent()
    {
        lock (Gate) return RecentLines.ToArray();
    }

    public static void Init(string? file)
    {
        if (string.IsNullOrWhiteSpace(file)) return;
        Directory.CreateDirectory(Path.GetDirectoryName(file)!);
        _file = file;
    }

    public static void Info(string message) => Write("INFO", message);
    public static void Warn(string message) => Write("WARN", message);
    public static void Error(string message) => Write("ERROR", message);

    private static void Write(string level, string message)
    {
        var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {message}";
        lock (Gate)
        {
            Console.WriteLine(line);
            RecentLines.AddLast(line);
            while (RecentLines.Count > 60) RecentLines.RemoveFirst();
            if (_file is null) return;
            try { File.AppendAllText(_file, line + Environment.NewLine); }
            catch { /* 日誌失敗不影響匯入 */ }
        }
    }
}

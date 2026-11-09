using System.Security.Cryptography;

namespace ImageClassification.Nas;

public sealed record ImportResult(int Copied, int Skipped, int Failed, int Deleted, long BytesCopied,
    IReadOnlyCollection<string> Folders, bool Aborted);

/// <summary>把一張卡的 DCIM 分類複製到 {OutputPath}/{品牌}/{型號}/{yyyy-MM-dd}/。</summary>
public sealed class Importer
{
    private const string Unknown = "Unknown";
    private readonly Settings _settings;
    private readonly bool _dryRun;
    private readonly Status? _status;

    public Importer(Settings settings, bool dryRun = false, Status? status = null)
    {
        _settings = settings;
        _dryRun = dryRun;
        _status = status;
    }

    public static string? FindDcim(string cardRoot)
    {
        if (!System.IO.Directory.Exists(cardRoot)) return null;
        return System.IO.Directory.EnumerateDirectories(cardRoot)
            .FirstOrDefault(d => string.Equals(Path.GetFileName(d), "DCIM", StringComparison.OrdinalIgnoreCase));
    }

    public static List<string> ListMedia(string dcim)
    {
        var files = new List<string>();
        var stack = new Stack<string>();
        stack.Push(dcim);
        while (stack.Count > 0)
        {
            var dir = stack.Pop();
            foreach (var sub in System.IO.Directory.EnumerateDirectories(dir))
            {
                if (!IsHidden(Path.GetFileName(sub))) stack.Push(sub);
            }
            files.AddRange(System.IO.Directory.EnumerateFiles(dir).Where(f => !IsHidden(Path.GetFileName(f))));
        }
        // 依資料夾、檔名排序：同一序號的 RAW/JPG/影片會相鄰，影片可沿用前一張照片的相機。
        files.Sort(StringComparer.OrdinalIgnoreCase);
        return files;
    }

    // DSM 會在卡上建 @eaDir（縮圖索引），macOS 會留 ._ 檔，都不是相機檔。
    private static bool IsHidden(string name) => name.StartsWith('.') || name.StartsWith('@') || name == "Thumbs.db";

    public ImportResult Import(string cardRoot, CancellationToken ct)
    {
        var dcim = FindDcim(cardRoot) ?? throw new DirectoryNotFoundException($"{cardRoot} 沒有 DCIM 資料夾");
        var files = ListMedia(dcim);
        Log.Info($"{cardRoot}: 共 {files.Count} 個檔案，目的地 {_settings.OutputPath}{(_dryRun ? "（演練，不複製）" : "")}");

        // 第一輪：讀 metadata，建立「同檔名」的相機對照（KCEM4123.ORF 與 KCEM4123.MOV 共用）。
        var infos = new Dictionary<string, MediaInfo>(StringComparer.OrdinalIgnoreCase);
        var cameraByStem = new Dictionary<string, MediaInfo>(StringComparer.OrdinalIgnoreCase);
        foreach (var f in files)
        {
            ct.ThrowIfCancellationRequested();
            var info = MediaInfo.Read(f);
            infos[f] = info;
            if (info.HasCamera) cameraByStem.TryAdd(StemKey(f), info);
        }

        var cameras = infos.Values.Where(i => i.HasCamera)
            .Select(i => $"{i.Make} {i.Model}".Trim()).Distinct().OrderBy(c => c).ToList();
        _status?.CardPlanned(files.Count, files.Sum(f => new FileInfo(f).Length), cameras);

        int copied = 0, skipped = 0, failed = 0, deleted = 0;
        long bytes = 0;
        var folders = new SortedSet<string>(StringComparer.Ordinal);
        MediaInfo? lastCamera = null;

        foreach (var src in files)
        {
            if (ct.IsCancellationRequested || !System.IO.Directory.Exists(dcim))
            {
                Log.Warn($"{cardRoot}: 卡片被移除或服務停止，中止。");
                return new ImportResult(copied, skipped, failed, deleted, bytes, folders, true);
            }

            var info = infos[src];
            var camera = info.HasCamera ? info : cameraByStem.GetValueOrDefault(StemKey(src)) ?? lastCamera;
            if (info.HasCamera) lastCamera = info;

            var make = camera?.Make ?? Unknown;
            var model = camera?.Model ?? Unknown;
            var taken = info.Taken ?? File.GetLastWriteTime(src);
            var destDir = Path.Combine(_settings.OutputPath, make, model, taken.ToString("yyyy-MM-dd"));
            folders.Add(Path.GetRelativePath(_settings.OutputPath, destDir));

            _status?.FileStarted(Path.GetFileName(src));
            long fileSize = 0;
            try
            {
                var outcome = CopyOne(src, destDir, out var dest, out var size);
                fileSize = size;
                if (outcome is Outcome.Copied or Outcome.DryRun) { copied++; bytes += size; }
                else skipped++;
                _status?.FileDone(size, outcome is Outcome.Copied, false);
                if (_settings.DeleteAfterCopy && outcome != Outcome.DryRun && DeleteVerified(src, dest!))
                {
                    deleted++;
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                failed++;
                _status?.FileDone(fileSize, false, true);
                Log.Error($"{src}: {ex.Message}");
                if (!System.IO.Directory.Exists(dcim))
                {
                    Log.Warn($"{cardRoot}: 卡片被移除，中止。");
                    return new ImportResult(copied, skipped, failed, deleted, bytes, folders, true);
                }
            }
        }

        if (!_dryRun) SyncDisks();
        return new ImportResult(copied, skipped, failed, deleted, bytes, folders, false);
    }

    [System.Runtime.InteropServices.DllImport("libc", EntryPoint = "sync")]
    private static extern void LibcSync();

    /// <summary>把所有尚未寫入的資料寫進硬碟（退卡與通知「可以拔卡」之前）。</summary>
    private static void SyncDisks()
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        try { LibcSync(); }
        catch (Exception ex) { Log.Warn($"sync 失敗: {ex.Message}"); return; }
        Log.Info($"資料已寫入硬碟（sync {sw.Elapsed.TotalSeconds:F1} 秒）");
    }

    private static string StemKey(string path) =>
        Path.Combine(Path.GetDirectoryName(path) ?? "", Path.GetFileNameWithoutExtension(path));

    private enum Outcome { Copied, Skipped, DryRun }

    private Outcome CopyOne(string src, string destDir, out string? dest, out long size)
    {
        var srcInfo = new FileInfo(src);
        size = srcInfo.Length;
        var name = Path.GetFileName(src);

        // 找第一個可用的目的檔名，絕不覆蓋：
        // - 同名、同大小、同修改時間 → 視為已匯入（複製時保留原時間，重插同一張卡很快）
        // - 同名同大小但時間不同 → 可能是同型號另一台機身的同序號檔，比 SHA-256 才算數
        // - 其他 → 改名 _1、_2…
        dest = null;
        for (var i = 0; ; i++)
        {
            var candidate = i == 0
                ? Path.Combine(destDir, name)
                : Path.Combine(destDir, $"{Path.GetFileNameWithoutExtension(name)}_{i}{Path.GetExtension(name)}");
            if (!File.Exists(candidate)) { dest = candidate; break; }
            if (IsSameFile(srcInfo, new FileInfo(candidate))) { dest = candidate; return Outcome.Skipped; }
        }

        if (_dryRun)
        {
            Log.Info($"[演練] {Path.GetFileName(src)} -> {Path.GetRelativePath(_settings.OutputPath, dest)}");
            return Outcome.DryRun;
        }

        System.IO.Directory.CreateDirectory(destDir);
        var part = dest + ".part";
        try
        {
            using (var input = new FileStream(src, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 20))
            using (var output = new FileStream(part, FileMode.Create, FileAccess.Write, FileShare.None, 1 << 20))
            {
                input.CopyTo(output, 1 << 20);
                // 不逐檔 fsync：在 btrfs 上慢 2.4 倍（實測 152MB：15.4s vs 6.5s）。
                // 整張卡複製完才 sync 一次（見 Import 結尾），退卡前資料一定落盤。
                // 例外：DeleteAfterCopy 會馬上刪卡上原檔，必須先確定這個檔已寫進硬碟。
                output.Flush(_settings.DeleteAfterCopy);
            }
            File.SetLastWriteTimeUtc(part, srcInfo.LastWriteTimeUtc);
            File.Move(part, dest);
        }
        catch
        {
            try { File.Delete(part); } catch { /* 盡力清掉半成品 */ }
            throw;
        }

        Log.Info($"{name} -> {Path.GetRelativePath(_settings.OutputPath, dest)}");
        return Outcome.Copied;
    }

    private static bool IsSameFile(FileInfo src, FileInfo existing)
    {
        if (src.Length != existing.Length) return false;
        // FAT 的時間解析度是 2 秒
        if (Math.Abs((src.LastWriteTimeUtc - existing.LastWriteTimeUtc).TotalSeconds) <= 2) return true;
        return HashEquals(src.FullName, existing.FullName);
    }

    /// <summary>兩邊 SHA-256 相同才刪卡上原檔。</summary>
    private static bool DeleteVerified(string src, string dest)
    {
        if (!HashEquals(src, dest))
        {
            Log.Error($"{src}: 與 {dest} 內容不同，保留卡上原檔。");
            return false;
        }
        File.Delete(src);
        return true;
    }

    private static bool HashEquals(string a, string b)
    {
        using var sa = File.OpenRead(a);
        using var sb = File.OpenRead(b);
        if (sa.Length != sb.Length) return false;
        return SHA256.HashData(sa).AsSpan().SequenceEqual(SHA256.HashData(sb));
    }
}

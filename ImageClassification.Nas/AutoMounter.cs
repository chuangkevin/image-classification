using System.Security.Cryptography;

namespace ImageClassification.Nas;

/// <summary>
/// DSM 只在「讀卡機插上 USB」時掛載卡片；同一台讀卡機換卡（媒體變更）不會自動掛載。
/// 這裡每輪檢查 USB 可移除磁碟：有媒體、沒被掛載、而且是沒處理過的卡，就呼叫
/// synousbdisk -mount 掛起來（與 DSM 介面掛的位置一致）。
/// 卡片以「磁碟名＋容量＋檔案系統的磁碟區序號」識別：退出後仍留在讀卡機裡的同一張卡不會被重掛，
/// 換卡後序號不同才會掛。不能用開機磁區雜湊：exFAT 的 VolumeFlags（掛載中旗標）與使用率
/// 欄位在掛載／退出時會變，同一張卡退出後會被誤認成新卡（2026-11-09 實際發生）。
/// </summary>
public sealed class AutoMounter
{
    private const string UsbDiskBin = "/usr/syno/bin/synousbdisk";
    private readonly HashSet<string> _seen = new();
    private readonly Dictionary<string, (DateTime At, int Count)> _attempts = new();
    private readonly Action<string> _alert;

    public AutoMounter(Action<string> alert) => _alert = alert;

    public void Tick(IReadOnlySet<string> mountedDevices)
    {
        if (!File.Exists(UsbDiskBin) || !System.IO.Directory.Exists("/sys/block")) return;
        var current = new HashSet<string>();

        foreach (var sys in System.IO.Directory.EnumerateDirectories("/sys/block", "sd*"))
        {
            var disk = Path.GetFileName(sys);
            try
            {
                if (ReadTrim(Path.Combine(sys, "removable")) != "1") continue;
                if (!ResolveLink(sys).Contains("/usb")) continue;
                if (!long.TryParse(ReadTrim(Path.Combine(sys, "size")), out var size) || size == 0) continue;

                var parts = System.IO.Directory.EnumerateDirectories(sys, disk + "*").Select(Path.GetFileName).ToList();
                var first = parts.OrderBy(p => p, StringComparer.Ordinal).FirstOrDefault() ?? disk;
                var id = $"{disk}|{size}|{VolumeId("/dev/" + first)}";
                current.Add(id);

                var anyMounted = parts.Append(disk).Any(p => mountedDevices.Contains("/dev/" + p));
                if (anyMounted) { _seen.Add(id); continue; }   // DSM 或我們已經掛好
                if (_seen.Contains(id)) continue;              // 處理過（例如匯入後已退出）

                // synousbdisk -mount 掛不上也可能回傳成功（例如沒裝 exFAT 套件），
                // 所以以「下一輪看到掛載」為準；掛不上就每 30 秒再試，只在第一次記錄。
                _attempts.TryGetValue(id, out var a);
                if (a.Count > 0 && DateTime.Now - a.At < TimeSpan.FromSeconds(30)) continue;
                if (a.Count == 0) Log.Info($"讀卡機 {disk} 有新卡（{size / 2048} MB）但沒被掛載，自動掛載");
                if (a.Count == 2)
                {
                    Log.Warn($"{disk} 掛載不上，可能是不支援的檔案系統（exFAT 要裝 exFAT Access 套件）");
                    _alert($"⚠️ 讀卡機 {disk} 的卡掛載不上，可能是不支援的檔案系統（exFAT 需安裝 exFAT Access 套件）。");
                }
                _attempts[id] = (DateTime.Now, a.Count + 1);
                Dsm.Mount(disk);
            }
            catch (Exception ex)
            {
                Log.Warn($"檢查 {disk} 失敗: {ex.Message}");
            }
        }

        // 拔掉或換掉的卡：忘掉，之後再插同一張卡會重新處理（已匯入的檔案會被略過）
        _seen.IntersectWith(current);
        foreach (var k in _attempts.Keys.Where(k => !current.Contains(k) || _seen.Contains(k)).ToList()) _attempts.Remove(k);
    }

    private static string ReadTrim(string path) => File.Exists(path) ? File.ReadAllText(path).Trim() : "";

    private static string ResolveLink(string path) =>
        new DirectoryInfo(path).ResolveLinkTarget(true)?.FullName ?? path;

    /// <summary>檔案系統的磁碟區序號（格式化時產生，掛載／退出不會變）。</summary>
    private static string VolumeId(string dev)
    {
        try
        {
            using var fs = new FileStream(dev, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var b = new byte[512];
            if (fs.Read(b, 0, b.Length) < 512) return "short";
            string Ascii(int off, int len) => System.Text.Encoding.ASCII.GetString(b, off, len);
            string Hex(int off, int len) => Convert.ToHexString(b, off, len);
            if (Ascii(3, 8) == "EXFAT   ") return "exfat-" + Hex(100, 4);
            if (Ascii(3, 8) == "NTFS    ") return "ntfs-" + Hex(72, 8);
            if (Ascii(82, 8) == "FAT32   ") return "fat32-" + Hex(67, 4);
            if (Ascii(54, 3) == "FAT") return "fat-" + Hex(39, 4);
            return "raw-" + Convert.ToHexString(SHA256.HashData(b))[..16];
        }
        catch
        {
            return "unreadable";
        }
    }
}

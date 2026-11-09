using System.Diagnostics;

namespace ImageClassification.Nas;

/// <summary>呼叫 DSM 內建工具。非 DSM 環境（找不到工具）時安靜略過。</summary>
public static class Dsm
{
    private const string UsbDiskBin = "/usr/syno/bin/synousbdisk";

    /// <summary>掛載 USB 磁碟（等同 DSM 偵測到裝置時的動作），例如 sdr。</summary>
    public static bool Mount(string disk)
    {
        if (!File.Exists(UsbDiskBin)) return false;
        return Run(UsbDiskBin, "-mount", disk) == 0;
    }

    /// <summary>退出 USB 磁碟（等同 DSM 介面的「退出」），例如 sdr。</summary>
    public static bool Eject(string disk)
    {
        if (!File.Exists(UsbDiskBin)) return false;
        return Run(UsbDiskBin, "-umount", disk) == 0;
    }

    private static int Run(string file, params string[] args)
    {
        try
        {
            var psi = new ProcessStartInfo(file) { RedirectStandardOutput = true, RedirectStandardError = true };
            foreach (var a in args) psi.ArgumentList.Add(a);
            using var p = Process.Start(psi)!;
            var err = p.StandardError.ReadToEnd();
            p.StandardOutput.ReadToEnd();
            if (!p.WaitForExit(60_000)) { p.Kill(); return -1; }
            if (p.ExitCode != 0) Log.Warn($"{Path.GetFileName(file)} {string.Join(' ', args)} 回傳 {p.ExitCode} {err.Trim()}");
            return p.ExitCode;
        }
        catch (Exception ex)
        {
            Log.Warn($"{Path.GetFileName(file)} 執行失敗: {ex.Message}");
            return -1;
        }
    }
}

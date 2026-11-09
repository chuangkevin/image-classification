using System.Runtime.InteropServices;
using ImageClassification.Nas;

// 用法：
//   sd-import run [--config 路徑]                     常駐監看（systemd 用）
//   sd-import import <卡片根目錄> [--dry-run] [--config 路徑]   手動匯入一次
//   sd-import notify-test [--config 路徑]              發一則 Telegram 測試訊息

var argv = args.ToList();
string? TakeOption(string name)
{
    var i = argv.IndexOf(name);
    if (i < 0 || i + 1 >= argv.Count) return null;
    var v = argv[i + 1];
    argv.RemoveRange(i, 2);
    return v;
}

var configPath = TakeOption("--config") ?? Path.Combine(AppContext.BaseDirectory, "sd-import.ini");
var dryRun = argv.Remove("--dry-run");
var command = argv.FirstOrDefault() ?? "run";

var settings = Settings.Load(configPath);
Log.Init(settings.LogPath);

using var cts = new CancellationTokenSource();
using var sigterm = PosixSignalRegistration.Create(PosixSignal.SIGTERM, ctx => { ctx.Cancel = true; cts.Cancel(); });
using var sigint = PosixSignalRegistration.Create(PosixSignal.SIGINT, ctx => { ctx.Cancel = true; cts.Cancel(); });

switch (command)
{
    case "run":
        new CardWatcher(settings, Path.GetDirectoryName(Path.GetFullPath(configPath))!).Run(cts.Token);
        return 0;

    case "import" when argv.Count >= 2:
        var r = new Importer(settings, dryRun).Import(argv[1], cts.Token);
        Log.Info($"{(dryRun ? "演練" : "完成")}：{(dryRun ? "將複製" : "新增")} {r.Copied} 個（{r.BytesCopied / 1048576.0:F0} MB），已存在 {r.Skipped}，失敗 {r.Failed}，刪除 {r.Deleted}");
        foreach (var f in r.Folders) Log.Info($"  資料夾 {f}");
        return r.Failed == 0 && !r.Aborted ? 0 : 1;

    case "notify-test":
        new Notifier(settings).Send("🔔 sd-import 通知測試：插卡匯入時會用這個頻道通知你。");
        return 0;

    default:
        Console.Error.WriteLine("用法: sd-import run | import <卡片根目錄> [--dry-run] | notify-test   [--config 路徑]");
        return 2;
}

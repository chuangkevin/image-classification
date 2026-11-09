using System.Net.Http.Json;

namespace ImageClassification.Nas;

/// <summary>
/// 透過 Telegram bot 發通知。token 與 chat id 每次發送時才從 env 檔讀
/// （TELEGRAM_BOT_TOKEN、TELEGRAM_HOME_CHANNEL），不另存副本。
/// DSM 7 的 synodsmnotify 只接受內建 i18n 字串，無法送自訂文字，所以不用它。
/// </summary>
public sealed class Notifier
{
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(15) };
    private readonly Settings _settings;

    public Notifier(Settings settings) => _settings = settings;

    public void Send(string text)
    {
        if (!_settings.Notify || string.IsNullOrWhiteSpace(_settings.TelegramEnvFile)) return;
        var env = ReadEnv(_settings.TelegramEnvFile);
        if (env is null
            || !env.TryGetValue("TELEGRAM_BOT_TOKEN", out var token)
            || !env.TryGetValue("TELEGRAM_HOME_CHANNEL", out var chat))
        {
            Log.Warn($"無法讀取 Telegram 設定 {_settings.TelegramEnvFile}，略過通知。");
            return;
        }

        try
        {
            var body = new { chat_id = chat, text = $"[{Environment.MachineName}] {text}", disable_web_page_preview = true };
            using var resp = Http.PostAsJsonAsync($"https://api.telegram.org/bot{token}/sendMessage", body)
                .GetAwaiter().GetResult();
            if (!resp.IsSuccessStatusCode)
            {
                Log.Warn($"Telegram 通知失敗：HTTP {(int)resp.StatusCode}");
            }
        }
        catch (Exception ex)
        {
            // 不把例外訊息全文寫進日誌：可能含帶 token 的網址。
            Log.Warn($"Telegram 通知失敗：{ex.GetType().Name}");
        }
    }

    private static Dictionary<string, string>? ReadEnv(string path)
    {
        try
        {
            return File.ReadAllLines(path)
                .Select(l => l.Trim())
                .Where(l => l.Length > 0 && !l.StartsWith('#') && l.Contains('='))
                .Select(l => l.Split('=', 2))
                .ToDictionary(p => p[0].Trim(), p => p[1].Trim().Trim('"', '\''));
        }
        catch
        {
            return null;
        }
    }
}

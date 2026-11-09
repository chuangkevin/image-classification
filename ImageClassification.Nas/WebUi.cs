using System.Net;
using System.Text;
using System.Text.Json;

namespace ImageClassification.Nas;

/// <summary>唯讀進度網頁：GET / 回頁面，GET /api/status 回 JSON。只在區網／tailnet 用。</summary>
public static class WebUi
{
    private static readonly JsonSerializerOptions Json = new()
    {
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public static void Start(string prefix, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(prefix)) return;
        var listener = new HttpListener();
        try
        {
            listener.Prefixes.Add(prefix);
            listener.Start();
        }
        catch (Exception ex)
        {
            Log.Warn($"進度網頁無法啟動（{prefix}）：{ex.Message}");
            return;
        }
        Log.Info($"進度網頁 {prefix}");
        ct.Register(() => { try { listener.Stop(); } catch { } });

        new Thread(() =>
        {
            while (!ct.IsCancellationRequested)
            {
                HttpListenerContext ctx;
                try { ctx = listener.GetContext(); }
                catch { break; }
                ThreadPool.QueueUserWorkItem(_ => Handle(ctx));
            }
        }) { IsBackground = true, Name = "web" }.Start();
    }

    private static void Handle(HttpListenerContext ctx)
    {
        try
        {
            var path = ctx.Request.Url?.AbsolutePath ?? "/";
            byte[] body;
            if (path == "/api/status")
            {
                body = JsonSerializer.SerializeToUtf8Bytes(Status.Instance.Snapshot(), Json);
                ctx.Response.ContentType = "application/json; charset=utf-8";
            }
            else if (path == "/")
            {
                body = Encoding.UTF8.GetBytes(Page);
                ctx.Response.ContentType = "text/html; charset=utf-8";
            }
            else if (path == "/manifest.webmanifest")
            {
                body = Encoding.UTF8.GetBytes(Manifest);
                ctx.Response.ContentType = "application/manifest+json; charset=utf-8";
            }
            else if (Icon(path) is { } png)
            {
                body = png;
                ctx.Response.ContentType = "image/png";
            }
            else
            {
                ctx.Response.StatusCode = 404;
                body = Array.Empty<byte>();
            }
            ctx.Response.Headers["Cache-Control"] = "no-store";
            ctx.Response.OutputStream.Write(body);
        }
        catch
        {
            // 瀏覽器中途關閉之類，忽略
        }
        finally
        {
            try { ctx.Response.Close(); } catch { }
        }
    }

    private static readonly Dictionary<string, byte[]> Icons = new();

    /// <summary>內嵌的 App 圖示：/icon-192.png、/icon-512.png、/apple-touch-icon.png。</summary>
    private static byte[]? Icon(string path)
    {
        var name = path switch
        {
            "/icon-192.png" => "icon-192.png",
            "/icon-512.png" => "icon-512.png",
            "/apple-touch-icon.png" or "/favicon.ico" => "icon-180.png",
            _ => null,
        };
        if (name is null) return null;
        lock (Icons)
        {
            if (Icons.TryGetValue(name, out var cached)) return cached;
            using var s = typeof(WebUi).Assembly.GetManifestResourceStream(name);
            if (s is null) return null;
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            return Icons[name] = ms.ToArray();
        }
    }

    private const string Manifest = """
{
  "name": "記憶卡匯入",
  "short_name": "記憶卡匯入",
  "start_url": "./",
  "scope": "./",
  "display": "standalone",
  "background_color": "#f6f7f9",
  "theme_color": "#2563eb",
  "icons": [
    { "src": "icon-192.png", "sizes": "192x192", "type": "image/png" },
    { "src": "icon-512.png", "sizes": "512x512", "type": "image/png" }
  ]
}
""";

    private const string Page = """
<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>記憶卡匯入</title>
<link rel="manifest" href="manifest.webmanifest">
<link rel="icon" href="icon-192.png">
<link rel="apple-touch-icon" href="apple-touch-icon.png">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-title" content="記憶卡匯入">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="theme-color" content="#2563eb">
<style>
:root{--bg:#f6f7f9;--card:#fff;--fg:#1d2330;--muted:#6b7280;--line:#e5e7eb;--accent:#2563eb;--ok:#15803d;--warn:#b45309;--bad:#b91c1c;--bar:#e8edf5}
@media (prefers-color-scheme:dark){:root{--bg:#0f1115;--card:#171a21;--fg:#e6e8ec;--muted:#9aa1ad;--line:#262b35;--accent:#60a5fa;--ok:#4ade80;--warn:#fbbf24;--bad:#f87171;--bar:#232833}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font:15px/1.5 -apple-system,BlinkMacSystemFont,"PingFang TC","Noto Sans TC","Microsoft JhengHei",sans-serif}
main{max-width:860px;margin:0 auto;padding:calc(20px + env(safe-area-inset-top)) 16px calc(40px + env(safe-area-inset-bottom))}
header{display:flex;justify-content:space-between;align-items:baseline;gap:12px;margin-bottom:16px}
h1{font-size:20px;margin:0;font-weight:600}
.sub{color:var(--muted);font-size:13px}
section{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:16px;margin-bottom:14px}
h2{font-size:13px;font-weight:600;color:var(--muted);margin:0 0 10px;letter-spacing:.02em}
.big{font-size:18px;font-weight:600;margin-bottom:6px}
.bar{height:10px;background:var(--bar);border-radius:5px;overflow:hidden;margin:10px 0}
.bar>i{display:block;height:100%;background:var(--accent);width:0;transition:width .4s}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:10px;margin-top:8px}
.k{color:var(--muted);font-size:12px}.v{font-size:16px;font-variant-numeric:tabular-nums}
.pill{display:inline-block;padding:1px 8px;border-radius:999px;font-size:12px;border:1px solid var(--line);color:var(--muted)}
.ok{color:var(--ok)}.warn{color:var(--warn)}.bad{color:var(--bad)}
table{width:100%;border-collapse:collapse;font-size:13px}
th,td{text-align:left;padding:6px 4px;border-bottom:1px solid var(--line);vertical-align:top}
th{color:var(--muted);font-weight:500}
th.n,td.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
.scroll{overflow-x:auto}
pre{margin:0;font:12px/1.5 ui-monospace,SFMono-Regular,Menlo,monospace;white-space:pre-wrap;word-break:break-all;color:var(--muted);max-height:260px;overflow:auto}
.cur{color:var(--muted);font-size:13px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
</style>
</head>
<body>
<main>
<header><h1>記憶卡匯入</h1><span class="sub" id="host"></span></header>

<section>
  <h2>目前的卡</h2>
  <div class="big" id="cardTitle">—</div>
  <div class="cur" id="cardSub"></div>
  <div class="bar"><i id="cardBar"></i></div>
  <div class="grid">
    <div><div class="k">進度</div><div class="v" id="cProg">—</div></div>
    <div><div class="k">新增</div><div class="v" id="cCopied">—</div></div>
    <div><div class="k">已存在略過</div><div class="v" id="cSkipped">—</div></div>
    <div><div class="k">失敗</div><div class="v" id="cFailed">—</div></div>
    <div><div class="k">速度</div><div class="v" id="cSpeed">—</div></div>
    <div><div class="k">剩餘時間</div><div class="v" id="cEta">—</div></div>
  </div>
</section>

<section>
  <h2>背景備援</h2>
  <div class="big" id="bTitle">—</div>
  <div class="cur" id="bSub"></div>
</section>

<section>
  <h2>最近處理的卡</h2>
  <div class="scroll"><table>
    <thead><tr><th>時間</th><th>相機</th><th>結果</th><th class="n">新增</th><th class="n">略過</th><th class="n">失敗</th><th class="n">MB</th><th class="n">分鐘</th></tr></thead>
    <tbody id="hist"><tr><td colspan="8" class="sub">還沒有</td></tr></tbody>
  </table></div>
</section>

<section>
  <h2>日誌</h2>
  <pre id="log"></pre>
</section>
</main>
<script>
const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const t = d => d ? new Date(d).toLocaleTimeString("zh-TW", {hour12:false}) : "";
const dur = s => { if (s == null) return "—"; s = Math.max(0, Math.round(s)); const h = Math.floor(s/3600), m = Math.floor(s%3600/60), x = s%60; return h ? `${h} 時 ${m} 分` : m ? `${m} 分 ${x} 秒` : `${x} 秒`; };
const since = d => d ? dur((Date.now() - new Date(d)) / 1000) : "";
async function tick() {
  let s;
  try { s = await (await fetch("api/status", {cache:"no-store"})).json(); }
  catch { $("host").textContent = "連不上服務"; return; }
  $("host").textContent = s.host + (s.paused.length ? ` · 已暫停 ${s.paused.join("、")}` : "");
  const c = s.card;
  const names = {idle:"等待插卡", scanning:"讀取相機資訊中", copying:"複製中", ejecting:"退出卡片中"};
  $("cardTitle").textContent = c.state === "idle" ? "等待插卡" : `${names[c.state]}${c.cameras.length ? "：" + c.cameras.join("、") : ""}`;
  $("cardSub").textContent = c.state === "idle" ? "插卡後會自動開始" : (c.current ? `正在處理 ${c.current}` : (c.label || ""));
  const pct = c.totalMb ? Math.min(100, c.doneMb / c.totalMb * 100) : (c.total ? c.done / c.total * 100 : 0);
  $("cardBar").style.width = (c.state === "idle" ? 0 : pct) + "%";
  $("cProg").textContent = c.state === "idle" ? "—" : `${c.done} / ${c.total}（${Math.floor(pct)}%）`;
  $("cCopied").textContent = c.state === "idle" ? "—" : c.copied;
  $("cSkipped").textContent = c.state === "idle" ? "—" : c.skipped;
  $("cFailed").innerHTML = c.state === "idle" ? "—" : (c.failed ? `<span class="bad">${c.failed}</span>` : "0");
  $("cSpeed").textContent = c.speedMbps != null ? `${c.speedMbps} MB/s` : "—";
  $("cEta").textContent = c.etaSeconds != null ? dur(c.etaSeconds) : "—";

  const b = s.backup;
  if (b.busy) { $("bTitle").innerHTML = `傳送中 <span class="pill">已 ${since(b.started)}</span>`; $("bSub").textContent = b.plan ? (b.last && b.last !== b.plan ? `${b.plan}　${b.last}` : b.plan) : "比對檔案清單中"; }
  else if (b.failing) { $("bTitle").innerHTML = `<span class="warn">上次失敗</span>，${t(b.nextRetry)} 重試`; $("bSub").textContent = b.last || ""; }
  else if (b.finished) { $("bTitle").innerHTML = `<span class="ok">已同步</span> <span class="pill">${t(b.finished)}</span>`; $("bSub").textContent = b.last || ""; }
  else { $("bTitle").textContent = "尚未執行"; $("bSub").textContent = ""; }

  $("hist").innerHTML = s.history.length ? s.history.map(h => {
    const cls = h.failed ? "bad" : h.result.startsWith("完成") ? "ok" : "warn";
    return `<tr><td>${t(h.at)}</td><td>${esc(h.cameras.join("、") || "—")}</td><td class="${cls}">${esc(h.result)}${h.ejected ? "，已退出" : ""}</td><td class="n">${h.copied}</td><td class="n">${h.skipped}</td><td class="n">${h.failed}</td><td class="n">${h.mb}</td><td class="n">${h.minutes}</td></tr>`;
  }).join("") : `<tr><td colspan="8" class="sub">還沒有</td></tr>`;

  const pre = $("log"), atBottom = pre.scrollTop + pre.clientHeight >= pre.scrollHeight - 8;
  pre.textContent = s.log.join("\n");
  if (atBottom) pre.scrollTop = pre.scrollHeight;
}
tick(); setInterval(tick, 1000);
</script>
</body>
</html>
""";
}

// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using System.Runtime.InteropServices;

namespace OfficeCli.Core;

/// <summary>
/// Headless HTML→PNG screenshot via shell-out to whichever browser is available.
/// Tries playwright CLI → Chromium-family (Chrome/Edge/Chromium) → Firefox.
/// No embedded browser engine; binary stays small.
/// </summary>
internal static class HtmlScreenshot
{
    public sealed record Result(bool Ok, string Backend, string? Error);

    public sealed record PaginationResult(int TotalPages, Dictionary<string, int> AnchorPageMap);

    /// Run a chromium-family browser in dump-dom mode against the given HTML
    /// and parse the document title for "PAGES:N|MAP:anchor=p,anchor=p,...".
    /// The HTML must set the title from JS after layout settles.
    /// <summary>
    /// Render <paramref name="htmlPath"/> in a chrome-family browser with a
    /// virtual-time budget (so async JS such as mermaid.js finishes) and return
    /// the final serialized DOM, or null if no chrome-family browser is found or
    /// the run fails. Callers extract whatever they need from the DOM (title,
    /// a rendered &lt;svg&gt;, …).
    /// </summary>
    public static string? DumpDom(string htmlPath, int timeoutMs = 60000)
    {
        var url = new Uri(Path.GetFullPath(htmlPath)).AbsoluteUri + "#screenshot";
        var bin = FindChrome();
        if (bin == null) return null;
        var args = new[]
        {
            "--headless=new",
            "--disable-gpu",
            "--no-sandbox",
            "--virtual-time-budget=15000",
            "--dump-dom",
            url,
        };
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = bin,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            foreach (var a in args) psi.ArgumentList.Add(a);
            using var p = Process.Start(psi);
            if (p == null) return null;
            var stdout = p.StandardOutput.ReadToEnd();
            if (!p.WaitForExit(timeoutMs)) { try { p.Kill(true); } catch { } return null; }
            return stdout;
        }
        catch { return null; }
    }

    /// <summary>True when a chrome-family browser (Chrome/Chromium/Edge) is available.</summary>
    public static bool HasChromeFamily() => FindChrome() != null;

    /// <summary>
    /// Screenshot <paramref name="htmlPath"/> in a chrome-family browser at an exact
    /// pixel window size, with a virtual-time budget (so async JS such as mermaid.js
    /// finishes before capture) and a HiDPI scale factor for crisp raster output.
    /// Returns true on a non-empty PNG at <paramref name="outPath"/>.
    /// </summary>
    public static bool CaptureChromeSized(string htmlPath, string outPath, int w, int h,
                                          int scale = 2, int timeoutMs = 60000)
    {
        var bin = FindChrome();
        if (bin == null) return false;
        outPath = Path.GetFullPath(outPath);
        var outDir = Path.GetDirectoryName(outPath);
        if (!string.IsNullOrEmpty(outDir)) Directory.CreateDirectory(outDir);
        var url = new Uri(Path.GetFullPath(htmlPath)).AbsoluteUri + "#screenshot";
        var args = new[]
        {
            "--headless=new",
            "--disable-gpu",
            "--no-sandbox",
            "--hide-scrollbars",
            $"--force-device-scale-factor={scale}",
            $"--window-size={w},{h}",
            "--virtual-time-budget=15000",
            "--default-background-color=00000000",
            $"--screenshot={outPath}",
            url,
        };
        var (ok, _) = RunBinary(bin, args);
        return ok && File.Exists(outPath) && new FileInfo(outPath).Length > 0;
    }

    public static PaginationResult? GetPaginationFromDom(string htmlPath, int timeoutMs = 60000)
    {
        var stdout = DumpDom(htmlPath, timeoutMs);
        if (stdout == null) return null;
        try
        {
            var m = System.Text.RegularExpressions.Regex.Match(stdout, @"<title>PAGES:(\d+)(?:\|MAP:([^<]*))?</title>");
            if (!m.Success || !int.TryParse(m.Groups[1].Value, out var n)) return null;
            var map = new Dictionary<string, int>();
            if (m.Groups[2].Success && m.Groups[2].Value.Length > 0)
            {
                foreach (var pair in m.Groups[2].Value.Split(','))
                {
                    var eq = pair.IndexOf('=');
                    if (eq > 0 && int.TryParse(pair[(eq + 1)..], out var pgNum))
                        map[pair[..eq]] = pgNum;
                }
            }
            return new PaginationResult(n, map);
        }
        catch { return null; }
    }

    public static int? GetPageCountFromDom(string htmlPath, int timeoutMs = 60000)
        => GetPaginationFromDom(htmlPath, timeoutMs)?.TotalPages;

    /// <summary>
    /// Pick a column count for a <paramref name="count"/>-item thumbnail contact
    /// sheet so the composed image is roughly square — used by `--grid auto`.
    /// Each cell is <paramref name="cellW"/>×<paramref name="cellH"/>; the grid
    /// image has aspect (rows·cellH)/(cols·cellW) = (count/cols²)·(cellH/cellW),
    /// so cols ≈ √(count·aspect) makes it ≈ 1. Portrait pages (aspect &gt; 1) get
    /// more columns, landscape slides (aspect &lt; 1) fewer. Clamped to [1, count].
    /// </summary>
    public static int AutoGridColumns(int count, double cellW, double cellH)
    {
        if (count <= 1) return 1;
        double aspect = cellH / Math.Max(1.0, cellW);
        int cols = (int)Math.Round(Math.Sqrt(count * aspect));
        return Math.Clamp(cols, 1, count);
    }

    public static Result Capture(string htmlPath, string outPath, int width = 1600, int height = 1200)
    {
        var url = new Uri(Path.GetFullPath(htmlPath)).AbsoluteUri + "#screenshot";
        outPath = Path.GetFullPath(outPath);
        var outDir = Path.GetDirectoryName(outPath);
        if (!string.IsNullOrEmpty(outDir)) Directory.CreateDirectory(outDir);

        // Cap to <= 1920px to stay within multi-image LLM limits.
        var (w, h) = CapDim(width, height, 1920);

        string? lastError = null;
        foreach (var (name, runner) in Backends())
        {
            var (ok, err) = runner(url, outPath, w, h);
            if (ok && File.Exists(outPath) && new FileInfo(outPath).Length > 0)
                return new Result(true, name, null);
            if (err != null) lastError = $"{name}: {err}";
        }
        return new Result(false, "", lastError ?? "no headless backend available");
    }

    private static IEnumerable<(string, Func<string, string, int, int, (bool, string?)>)> Backends()
    {
        yield return ("playwright", TryPlaywright);
        yield return ("chrome", TryChrome);
        yield return ("firefox", TryFirefox);
    }

    private static (int, int) CapDim(int w, int h, int limit)
    {
        var m = Math.Max(w, h);
        if (m <= limit) return (w, h);
        var s = (double)limit / m;
        return (Math.Max(1, (int)(w * s)), Math.Max(1, (int)(h * s)));
    }

    // ----- Playwright CLI -----------------------------------------------------------------

    private static (bool, string?) TryPlaywright(string url, string outPath, int w, int h)
    {
        var pw = WhichFirst("playwright");
        if (pw == null) return (false, null);
        var args = new[] { "screenshot", $"--viewport-size={w},{h}", "--full-page", url, outPath };
        return RunBinary(pw, args);
    }

    // ----- Chromium family ---------------------------------------------------------------

    private static (bool, string?) TryChrome(string url, string outPath, int w, int h)
    {
        var bin = FindChrome();
        if (bin == null) return (false, null);
        var args = new[]
        {
            "--headless=new",
            "--disable-gpu",
            "--no-sandbox",
            "--hide-scrollbars",
            $"--window-size={w},{h}",
            $"--screenshot={outPath}",
            url,
        };
        return RunBinary(bin, args);
    }

    private static string? FindChrome()
    {
        string[] names = ["google-chrome", "google-chrome-stable", "chromium", "chromium-browser",
                          "chrome", "microsoft-edge", "microsoft-edge-stable", "msedge"];
        var pathHit = WhichFirst(names);
        if (pathHit != null) return pathHit;

        var abs = new List<string>();
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            abs.AddRange(new[]
            {
                "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                "/Applications/Chromium.app/Contents/MacOS/Chromium",
                "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
            });
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            abs.AddRange(new[]
            {
                "/usr/bin/google-chrome", "/usr/bin/chromium", "/usr/bin/chromium-browser",
                "/snap/bin/chromium", "/snap/bin/google-chrome",
                "/usr/bin/microsoft-edge", "/usr/bin/microsoft-edge-stable",
            });
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            string[] roots = [
                Environment.GetEnvironmentVariable("PROGRAMFILES") ?? @"C:\Program Files",
                Environment.GetEnvironmentVariable("PROGRAMFILES(X86)") ?? @"C:\Program Files (x86)",
                Environment.GetEnvironmentVariable("LOCALAPPDATA") ?? "",
            ];
            string[] suffixes = [
                @"Google\Chrome\Application\chrome.exe",
                @"Chromium\Application\chrome.exe",
                @"Microsoft\Edge\Application\msedge.exe",
            ];
            foreach (var r in roots)
                if (!string.IsNullOrEmpty(r))
                    foreach (var s in suffixes) abs.Add(Path.Combine(r, s));
        }
        return abs.FirstOrDefault(File.Exists);
    }

    // ----- Firefox -----------------------------------------------------------------------

    private static (bool, string?) TryFirefox(string url, string outPath, int w, int h)
    {
        var bin = FindFirefox();
        if (bin == null) return (false, null);
        // Firefox: `--headless --screenshot=<out> --window-size=W,H <URL>`.
        // Note: no `=new` headless variant; --force-device-scale-factor not supported.
        var args = new[] { "--headless", $"--screenshot={outPath}", $"--window-size={w},{h}", url };
        return RunBinary(bin, args);
    }

    private static string? FindFirefox()
    {
        var pathHit = WhichFirst("firefox", "firefox-esr");
        if (pathHit != null) return pathHit;

        var abs = new List<string>();
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            abs.AddRange(new[]
            {
                "/Applications/Firefox.app/Contents/MacOS/firefox",
                "/Applications/Firefox Developer Edition.app/Contents/MacOS/firefox",
                "/Applications/Firefox Nightly.app/Contents/MacOS/firefox",
            });
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            abs.AddRange(new[] { "/usr/bin/firefox", "/usr/bin/firefox-esr", "/snap/bin/firefox" });
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            foreach (var r in new[]
            {
                Environment.GetEnvironmentVariable("PROGRAMFILES") ?? @"C:\Program Files",
                Environment.GetEnvironmentVariable("PROGRAMFILES(X86)") ?? @"C:\Program Files (x86)",
            })
                if (!string.IsNullOrEmpty(r)) abs.Add(Path.Combine(r, @"Mozilla Firefox\firefox.exe"));
        }
        return abs.FirstOrDefault(File.Exists);
    }

    // ----- Helpers -----------------------------------------------------------------------

    /// <summary>Find the first of <paramref name="names"/> on PATH (honouring
    /// Windows PATHEXT). Shared executable lookup — same mechanism used to detect
    /// playwright/chrome/firefox, so callers like the mermaid renderer detect mmdc
    /// identically.</summary>
    public static string? Which(params string[] names) => WhichFirst(names);

    private static string? WhichFirst(params string[] names)
    {
        var pathSep = RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? ';' : ':';
        var pathExt = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? (Environment.GetEnvironmentVariable("PATHEXT") ?? ".COM;.EXE;.BAT;.CMD").Split(';')
            : new[] { "" };
        var paths = (Environment.GetEnvironmentVariable("PATH") ?? "").Split(pathSep);

        foreach (var name in names)
        {
            foreach (var dir in paths)
            {
                if (string.IsNullOrEmpty(dir)) continue;
                foreach (var ext in pathExt)
                {
                    var candidate = Path.Combine(dir, name + ext);
                    if (File.Exists(candidate)) return candidate;
                }
            }
        }
        return null;
    }

    private static (bool, string?) RunBinary(string bin, string[] args)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = bin,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            foreach (var a in args) psi.ArgumentList.Add(a);
            using var p = Process.Start(psi);
            if (p == null) return (false, "process did not start");
            if (!p.WaitForExit(120_000))
            {
                try { p.Kill(true); } catch { /* ignore */ }
                return (false, "timeout after 120s");
            }
            if (p.ExitCode != 0)
            {
                var stderr = p.StandardError.ReadToEnd();
                var lastLine = stderr.Trim().Split('\n').LastOrDefault() ?? $"exit {p.ExitCode}";
                return (false, lastLine);
            }
            return (true, null);
        }
        catch (Exception e)
        {
            return (false, e.Message);
        }
    }
}

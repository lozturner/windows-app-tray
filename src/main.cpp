// TraySys v0.1.0
// One Shell_NotifyIcon per active top-level window — that's the entire product.
// + a tiny embedded HTTP server on http://127.0.0.1:8731 for settings & live status.
// Native Win32 / C++17. Built with MinGW-w64 g++. No runtime, single .exe.
//
// Build:
//   build.bat                           (or)
//   g++ -std=c++17 -O2 -municode -mwindows src/main.cpp -o build/TraySys.exe \
//       -luser32 -lgdi32 -lshell32 -lshlwapi -lpsapi -ldwmapi -lole32 -lws2_32

#define WIN32_LEAN_AND_MEAN
#ifndef WINVER
#define WINVER 0x0A00
#endif
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0A00
#endif

#include <winsock2.h>
#include <ws2tcpip.h>
#include <windows.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <psapi.h>
#include <dwmapi.h>
#include <shobjidl.h>
#include <objbase.h>
#include <cstdio>
#include <cstdarg>
#include <cstring>
#include <chrono>
#include <string>
#include <vector>
#include <unordered_map>
#include <unordered_set>
#include <thread>
#include <mutex>
#include <atomic>
#include <algorithm>
#include <random>

// ---------- config ----------
#define TRAYSYS_VERSION       L"0.4.0"
#define TRAYSYS_VERSION_A      "0.4.0"
#define TIMER_DEBOUNCE_REFRESH 7
#define REMINDER_INTERVAL_HOURS 168                 // weekly nudge
#define PROFILE_SLOTS          5
#define TRAY_CALLBACK_MSG     (WM_APP + 1)
#define HTTP_PORT             8731
#define REFRESH_INTERVAL_MS   1000
#define MASTER_TRAY_UID       0x7000             // dedicated UID for the animated master icon
#define PINNED_UID_BASE       0x6000             // 0x6000..0x6FFF reserved for pinned-app launcher icons
#define STAR_FRAME_COUNT      8
#define STAR_FRAME_BASE_RES   100                // resource IDs 100..107
#define STAR_ANIM_INTERVAL_MS 140                // ~7 fps colour cycle

// ---------- log + bench + ini ----------
// Paths resolved at startup to be next to the .exe so we never lose them to a weird CWD.
static char g_logPath[MAX_PATH]   = "traysys.log";
static char g_benchPath[MAX_PATH] = "traysys.bench.csv";
static char g_iniPath[MAX_PATH]   = "traysys.ini";

static void resolveDataPaths() {
    char exe[MAX_PATH] = {};
    GetModuleFileNameA(nullptr, exe, MAX_PATH);
    char* slash = strrchr(exe, '\\');
    if (slash) {
        *slash = 0;
        snprintf(g_logPath,   MAX_PATH, "%s\\LaurenceTrayhost_v%s.log",       exe, TRAYSYS_VERSION_A);
        snprintf(g_benchPath, MAX_PATH, "%s\\LaurenceTrayhost_v%s.bench.csv", exe, TRAYSYS_VERSION_A);
        snprintf(g_iniPath,   MAX_PATH, "%s\\LaurenceTrayhost_v%s.ini",       exe, TRAYSYS_VERSION_A);
    }
}

// ---------- INI live config ----------
static int  iniGetInt(const char* section, const char* key, int def) {
    return GetPrivateProfileIntA(section, key, def, g_iniPath);
}
static bool iniGetBool(const char* section, const char* key, bool def) {
    return iniGetInt(section, key, def ? 1 : 0) != 0;
}
static void iniSetBool(const char* section, const char* key, bool val) {
    WritePrivateProfileStringA(section, key, val ? "1" : "0", g_iniPath);
}
static long long iniGetLL(const char* section, const char* key, long long def) {
    char buf[64]; char defStr[64];
    snprintf(defStr, sizeof(defStr), "%lld", def);
    GetPrivateProfileStringA(section, key, defStr, buf, sizeof(buf), g_iniPath);
    return _atoi64(buf);
}
static void iniSetLL(const char* section, const char* key, long long val) {
    char buf[64]; snprintf(buf, sizeof(buf), "%lld", val);
    WritePrivateProfileStringA(section, key, buf, g_iniPath);
}

// ---------- autostart (HKCU\Software\Microsoft\Windows\CurrentVersion\Run) ----------
static const wchar_t* AUTOSTART_KEY  = L"Software\\Microsoft\\Windows\\CurrentVersion\\Run";
static const wchar_t* AUTOSTART_NAME = L"LaurenceTrayhost";

static bool autostartIsEnabled() {
    HKEY h{};
    if (RegOpenKeyExW(HKEY_CURRENT_USER, AUTOSTART_KEY, 0, KEY_READ, &h) != ERROR_SUCCESS) return false;
    bool found = (RegQueryValueExW(h, AUTOSTART_NAME, nullptr, nullptr, nullptr, nullptr) == ERROR_SUCCESS);
    RegCloseKey(h);
    return found;
}

static int  clampInt(int v, int lo, int hi) { return v < lo ? lo : (v > hi ? hi : v); }

// Forward decls — actual atomic globals defined further down with the rest of
// the runtime state. This lets the load/save helpers live in the INI block.
extern std::atomic<int> g_refreshIntervalMs;
extern std::atomic<int> g_debounceMs;
extern std::atomic<int> g_starAnimMs;
extern std::atomic<int> g_engine;

static void tuningLoad() {
    g_refreshIntervalMs.store(clampInt(iniGetInt("tuning", "refresh_interval_ms", 5000), 500, 60000));
    g_debounceMs.store       (clampInt(iniGetInt("tuning", "debounce_ms",          150),  20,  2000));
    g_starAnimMs.store       (clampInt(iniGetInt("tuning", "star_anim_ms",         140),   0,  5000));
    g_engine.store           (clampInt(iniGetInt("engine", "mode",                  0),    0,     2));
}
static void engineSave() {
    char b[8]; snprintf(b, 8, "%d", g_engine.load());
    WritePrivateProfileStringA("engine", "mode", b, g_iniPath);
}
static void tuningSave() {
    char b[32];
    snprintf(b, 32, "%d", g_refreshIntervalMs.load());
    WritePrivateProfileStringA("tuning", "refresh_interval_ms", b, g_iniPath);
    snprintf(b, 32, "%d", g_debounceMs.load());
    WritePrivateProfileStringA("tuning", "debounce_ms", b, g_iniPath);
    snprintf(b, 32, "%d", g_starAnimMs.load());
    WritePrivateProfileStringA("tuning", "star_anim_ms", b, g_iniPath);
}

static void autostartSet(bool on) {
    HKEY h{};
    if (RegOpenKeyExW(HKEY_CURRENT_USER, AUTOSTART_KEY, 0, KEY_WRITE, &h) != ERROR_SUCCESS) return;
    if (on) {
        wchar_t exe[MAX_PATH]; GetModuleFileNameW(nullptr, exe, MAX_PATH);
        std::wstring quoted = L"\"" + std::wstring(exe) + L"\"";
        RegSetValueExW(h, AUTOSTART_NAME, 0, REG_SZ,
            (const BYTE*)quoted.c_str(), (DWORD)((quoted.size() + 1) * sizeof(wchar_t)));
    } else {
        RegDeleteValueW(h, AUTOSTART_NAME);
    }
    RegCloseKey(h);
}

static std::mutex g_logMx;
static void logf(const char* fmt, ...) {
    std::lock_guard<std::mutex> lk(g_logMx);
    char line[1024];
    SYSTEMTIME st; GetLocalTime(&st);
    int n = snprintf(line, sizeof(line), "[%02d:%02d:%02d.%03d] ",
                     st.wHour, st.wMinute, st.wSecond, st.wMilliseconds);
    va_list ap; va_start(ap, fmt);
    n += vsnprintf(line + n, sizeof(line) - n - 2, fmt, ap);
    va_end(ap);
    if (n < (int)sizeof(line) - 1) line[n++] = '\n';
    line[n] = 0;
    OutputDebugStringA(line);
    FILE* f = fopen(g_logPath, "a");
    if (f) { fputs(line, f); fclose(f); }
}

struct BenchSample {
    long long ts_ms;        // since epoch
    double   refresh_ms;    // last refresh tick took this long
    int      icons;         // icons currently registered
    size_t   rss_kb;        // working set
    int      added, removed;
};
static std::mutex g_benchMx;
static void benchHeader() {
    FILE* f = fopen(g_benchPath, "r");
    if (f) { fclose(f); return; }
    f = fopen(g_benchPath, "w");
    if (!f) return;
    fputs("ts_ms,refresh_ms,icons,rss_kb,added,removed\n", f);
    fclose(f);
}
static void benchAppend(const BenchSample& s) {
    std::lock_guard<std::mutex> lk(g_benchMx);
    FILE* f = fopen(g_benchPath, "a");
    if (!f) return;
    fprintf(f, "%lld,%.3f,%d,%zu,%d,%d\n",
            s.ts_ms, s.refresh_ms, s.icons, s.rss_kb, s.added, s.removed);
    fclose(f);
}
static long long nowMs() {
    return std::chrono::duration_cast<std::chrono::milliseconds>(
        std::chrono::system_clock::now().time_since_epoch()).count();
}
static size_t rssKB() {
    PROCESS_MEMORY_COUNTERS pmc{};
    if (GetProcessMemoryInfo(GetCurrentProcess(), &pmc, sizeof(pmc)))
        return pmc.WorkingSetSize / 1024;
    return 0;
}

// ---------- model ----------
struct ManagedWindow {
    HWND     hwnd;
    UINT     uid;          // tray icon ID (stable within process lifetime)
    HICON    rawIcon;      // unbadged source icon (we own — DestroyIcon on remove)
    HICON    badgedIcon;   // composited icon currently shown (we own — DestroyIcon on replace)
    int      badgeNum = 0; // 1..N if N windows share this exe; 0 means no badge needed
    std::wstring title;
    std::wstring exeName;       // lower-case basename
    std::wstring exeNameLower;
    bool     hiddenFromTaskbar = false;  // tracked so we know whether to call DeleteTab
};

static std::unordered_map<HWND, ManagedWindow> g_windows;
static std::mutex                              g_windowsMx;
static UINT                                    g_nextUid = 1;
static HWND                                    g_msgWnd = nullptr;
static std::atomic<bool>                       g_quit{false};
// uid -> hwnd for fast click routing
static std::unordered_map<UINT, HWND>          g_uidToHwnd;

// last refresh stats for /status endpoint
static std::atomic<long long>  g_lastRefreshMs{0};
static std::atomic<double>     g_lastRefreshDuration{0.0};

// runtime flags (settable via /api/config in a future patch)
static std::atomic<bool>       g_hideBrowsersFromTaskbar{true};
static std::atomic<bool>       g_hideAllFromTaskbar{true};      // default: full takeover

// TaskbarCreated registered message — explorer.exe broadcasts this when the
// shell starts/restarts. We re-register every tray icon when it arrives,
// which is THE fix for the "icons go blank until you hover" Windows quirk.
static UINT                    g_taskbarCreatedMsg = 0;

// Live-tunable timings (default values; overridden from INI [tuning] at startup;
// settings UI sliders update these live via /api/tuning POST).
static std::atomic<int>        g_refreshIntervalMs{5000};   // safety-net poll
static std::atomic<int>        g_debounceMs{150};           // event-burst coalesce
static std::atomic<int>        g_starAnimMs{140};           // 0 = stop animation

// Per-window event hook (WinEvents) — replaces polling-as-primary with
// "react when Windows tells us a window changed".
static HWINEVENTHOOK           g_winEventHook = nullptr;

// Shell hook message — explorer.exe pushes HSHELL_* notifications directly.
// This is the cleanest push channel for window lifecycle.
static UINT                    g_shellHookMsg = 0;

// Engine selector — three independent strategies:
//   PUSH   (default): RegisterShellHookWindow + WinEventHook, NO polling.
//   HYBRID:           push events + the safety-net poll thread.
//   POLL:              legacy — disable both push hooks, poll-only.
enum class Engine : int { PUSH = 0, HYBRID = 1, POLL = 2 };
static std::atomic<int>        g_engine{(int)Engine::PUSH};

// COM ITaskbarList for AddTab/DeleteTab — hide browser windows from the real taskbar.
// MUST only be touched from the main message thread (the COM apartment that created it).
static ITaskbarList*           g_taskbarList = nullptr;
#define WM_TBL_SETVISIBLE      (WM_APP + 3)        // wParam = hwnd, lParam = visible

static void taskbarListInit() {
    if (g_taskbarList) return;
    HRESULT hr = CoCreateInstance(CLSID_TaskbarList, nullptr, CLSCTX_INPROC_SERVER,
                                  IID_ITaskbarList, (void**)&g_taskbarList);
    if (SUCCEEDED(hr) && g_taskbarList) g_taskbarList->HrInit();
}

// Direct call (main thread only).
static void taskbarListSetVisibleDirect(HWND h, bool visible) {
    if (!g_taskbarList) return;
    if (visible) g_taskbarList->AddTab(h);
    else         g_taskbarList->DeleteTab(h);
}

// Thread-safe wrapper: posts a message to the main thread so the call happens
// in the COM apartment that owns g_taskbarList.
static void taskbarListSetVisible(HWND h, bool visible) {
    if (!g_msgWnd) { taskbarListSetVisibleDirect(h, visible); return; }
    PostMessageW(g_msgWnd, WM_TBL_SETVISIBLE, (WPARAM)h, visible ? 1 : 0);
}

// ---------- util ----------
static std::wstring exeBasenameForHwnd(HWND h) {
    DWORD pid = 0;
    GetWindowThreadProcessId(h, &pid);
    if (!pid) return L"";
    HANDLE hp = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, pid);
    if (!hp) return L"";
    wchar_t path[MAX_PATH] = {};
    DWORD sz = MAX_PATH;
    std::wstring out;
    if (QueryFullProcessImageNameW(hp, 0, path, &sz)) {
        const wchar_t* base = PathFindFileNameW(path);
        out = base ? base : L"";
    }
    CloseHandle(hp);
    return out;
}

static std::wstring titleOf(HWND h) {
    int len = GetWindowTextLengthW(h);
    if (len <= 0) return L"";
    std::wstring s(len + 1, L'\0');
    int got = GetWindowTextW(h, s.data(), len + 1);
    s.resize(got);
    return s;
}

static HICON copyIconForHwnd(HWND h) {
    HICON src = (HICON)SendMessageW(h, WM_GETICON, ICON_SMALL2, 0);
    if (!src) src = (HICON)SendMessageW(h, WM_GETICON, ICON_SMALL, 0);
    if (!src) src = (HICON)SendMessageW(h, WM_GETICON, ICON_BIG, 0);
    if (!src) src = (HICON)GetClassLongPtrW(h, GCLP_HICONSM);
    if (!src) src = (HICON)GetClassLongPtrW(h, GCLP_HICON);
    if (!src) src = LoadIcon(nullptr, IDI_APPLICATION);
    return CopyIcon(src);
}

// Draw a small filled circle with a number on top of the source icon.
// Returns a brand-new HICON (caller owns and must DestroyIcon).
// badgeColor encoded as 0x00BBGGRR (Win32 COLORREF).
static HICON makeBadgedIcon(HICON src, int number, COLORREF badgeColor) {
    if (!src || number <= 0) return CopyIcon(src);

    const int sz = 32;
    HDC screen = GetDC(nullptr);
    HDC mem    = CreateCompatibleDC(screen);

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = sz;
    bmi.bmiHeader.biHeight      = -sz;          // top-down DIB
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    HBITMAP color = CreateDIBSection(screen, &bmi, DIB_RGB_COLORS, &bits, nullptr, 0);
    HBITMAP mask  = CreateBitmap(sz, sz, 1, 1, nullptr);

    HBITMAP oldBmp = (HBITMAP)SelectObject(mem, color);

    // 1) draw the source icon (preserves alpha)
    DrawIconEx(mem, 0, 0, src, sz, sz, 0, nullptr, DI_NORMAL);

    // 2) badge: filled circle bottom-right with white outline,
    //    sized large so it stays visible when Windows downsamples the icon
    //    from 32x32 to 16x16 in the notification area.
    int br = 11;                            // radius
    int cx = sz - br, cy = sz - br;

    // white outer ring (3px thick) for contrast
    HBRUSH ringBrush = CreateSolidBrush(RGB(255, 255, 255));
    HPEN   ringPen   = CreatePen(PS_NULL, 0, 0);
    HBRUSH oldBrush = (HBRUSH)SelectObject(mem, ringBrush);
    HPEN   oldPen   = (HPEN)SelectObject(mem, ringPen);
    Ellipse(mem, cx - br, cy - br, cx + br, cy + br);
    SelectObject(mem, oldBrush);
    SelectObject(mem, oldPen);
    DeleteObject(ringBrush);
    DeleteObject(ringPen);

    // coloured fill, slightly smaller
    int fr = br - 2;
    HBRUSH bg = CreateSolidBrush(badgeColor);
    HPEN   bgPen = CreatePen(PS_NULL, 0, 0);
    oldBrush = (HBRUSH)SelectObject(mem, bg);
    oldPen   = (HPEN)SelectObject(mem, bgPen);
    Ellipse(mem, cx - fr, cy - fr, cx + fr, cy + fr);
    SelectObject(mem, oldBrush);
    SelectObject(mem, oldPen);
    DeleteObject(bg);
    DeleteObject(bgPen);

    // 3) number on top — bold sans, scaled with badge
    HFONT font = CreateFontW(18, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        ANTIALIASED_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
    HFONT oldFont = (HFONT)SelectObject(mem, font);
    SetBkMode(mem, TRANSPARENT);
    SetTextColor(mem, RGB(255, 255, 255));
    wchar_t buf[8]; swprintf(buf, 8, L"%d", number > 9 ? 9 : number);  // single digit fits cleanly
    RECT r = { cx - br, cy - br, cx + br, cy + br };
    DrawTextW(mem, buf, -1, &r, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    SelectObject(mem, oldFont);
    DeleteObject(font);

    // 4) restore alpha on the badge area — Ellipse on a 32-bit DIB clears alpha to 0,
    //    which makes the badge transparent. Force alpha=255 inside the badge bbox.
    auto* pixels = static_cast<unsigned char*>(bits);
    int x0 = cx - br - 1, y0 = cy - br - 1;
    int x1 = cx + br + 1, y1 = cy + br + 1;
    if (x0 < 0) x0 = 0; if (y0 < 0) y0 = 0;
    if (x1 > sz) x1 = sz; if (y1 > sz) y1 = sz;
    for (int y = y0; y < y1; ++y)
        for (int x = x0; x < x1; ++x)
            pixels[(y * sz + x) * 4 + 3] = 255;

    SelectObject(mem, oldBmp);

    ICONINFO ii{};
    ii.fIcon    = TRUE;
    ii.hbmColor = color;
    ii.hbmMask  = mask;
    HICON out = CreateIconIndirect(&ii);

    DeleteObject(color);
    DeleteObject(mask);
    DeleteDC(mem);
    ReleaseDC(nullptr, screen);
    return out;
}

// Convert any HICON into a desaturated, dimmed version — used for static
// pinned-launcher icons so they read as "shortcuts" rather than active windows.
static HICON makeDimIcon(HICON src) {
    if (!src) return nullptr;
    const int sz = 32;
    HDC screen = GetDC(nullptr);
    HDC mem    = CreateCompatibleDC(screen);

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = sz;
    bmi.bmiHeader.biHeight      = -sz;
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    HBITMAP color = CreateDIBSection(screen, &bmi, DIB_RGB_COLORS, &bits, nullptr, 0);
    HBITMAP mask  = CreateBitmap(sz, sz, 1, 1, nullptr);
    HBITMAP oldBmp = (HBITMAP)SelectObject(mem, color);

    DrawIconEx(mem, 0, 0, src, sz, sz, 0, nullptr, DI_NORMAL);

    // Desaturate + dim: every pixel -> luminance-blend(60% gray, 40% colour),
    // then alpha multiplied by 0.55.
    auto* px = static_cast<unsigned char*>(bits);
    for (int i = 0; i < sz * sz; ++i) {
        unsigned char b = px[i*4 + 0];
        unsigned char g = px[i*4 + 1];
        unsigned char r = px[i*4 + 2];
        unsigned char a = px[i*4 + 3];
        // Y' = 0.299R + 0.587G + 0.114B  (Rec.601 luma)
        unsigned char y = (unsigned char)((r * 77 + g * 150 + b * 29) >> 8);
        px[i*4 + 0] = (unsigned char)((b + y * 3) >> 2);   // 25% colour, 75% gray
        px[i*4 + 1] = (unsigned char)((g + y * 3) >> 2);
        px[i*4 + 2] = (unsigned char)((r + y * 3) >> 2);
        px[i*4 + 3] = (unsigned char)((a * 140) >> 8);     // ~55% opacity
    }

    SelectObject(mem, oldBmp);

    ICONINFO ii{}; ii.fIcon = TRUE; ii.hbmColor = color; ii.hbmMask = mask;
    HICON out = CreateIconIndirect(&ii);

    DeleteObject(color); DeleteObject(mask);
    DeleteDC(mem); ReleaseDC(nullptr, screen);
    return out;
}

// Classify exe name: returns badge color for distinguishing categories.
//   browsers -> amber/orange (so duplicates of Chrome look like browser-tabs)
//   default  -> blue
static COLORREF badgeColorForExe(const std::wstring& exeLower) {
    static const wchar_t* browsers[] = {
        L"chrome.exe", L"msedge.exe", L"comet.exe", L"firefox.exe",
        L"brave.exe", L"opera.exe", L"vivaldi.exe", L"arc.exe"
    };
    for (const wchar_t* b : browsers)
        if (lstrcmpiW(exeLower.c_str(), b) == 0) return RGB(230, 130, 30);  // amber
    return RGB(40, 130, 220);  // blue
}

static bool isBrowserExe(const std::wstring& exe) {
    return badgeColorForExe(exe) == RGB(230, 130, 30);
}

// Coarse exe categorization for the Group / Stack action.
// Lower number = grouped earlier in the tray.
enum class IconCategory : int {
    BROWSER       = 0,
    SEARCH_AI     = 1,
    SYSTEM        = 2,
    PRODUCTIVITY  = 3,
    OTHER_ACTIVE  = 4,
    PINNED        = 5,   // greyed static launchers — always last
};

static IconCategory categorizeExe(const std::wstring& exeLower) {
    if (isBrowserExe(exeLower)) return IconCategory::BROWSER;

    static const wchar_t* search_ai[] = {
        L"chatgpt.exe", L"claude.exe", L"copilot.exe", L"perplexity.exe",
        L"gemini.exe",  L"groqchat.exe", L"better gpt.exe", L"chatgpt lt.exe",
        L"chatgpt 4.exe", L"chatgpt3.exe", L"jan.exe", L"lm studio.exe",
        L"gpt4all.exe", L"notebooklm.exe", L"google ai studio.exe",
        L"everything.exe", L"command palette.exe"
    };
    for (auto* s : search_ai) if (lstrcmpiW(exeLower.c_str(), s) == 0) return IconCategory::SEARCH_AI;

    static const wchar_t* system[] = {
        L"explorer.exe", L"cmd.exe", L"powershell.exe", L"taskmgr.exe",
        L"notepad.exe", L"calc.exe", L"calculator.exe", L"systemsettings.exe",
        L"control.exe", L"regedit.exe", L"mmc.exe", L"snippingtool.exe",
        L"voicerecorder.exe", L"camera.exe", L"clock.exe", L"copyq.exe",
        L"applicationframehost.exe"
    };
    for (auto* s : system) if (lstrcmpiW(exeLower.c_str(), s) == 0) return IconCategory::SYSTEM;

    static const wchar_t* productivity[] = {
        L"excel.exe", L"winword.exe", L"powerpnt.exe", L"outlook.exe",
        L"onenote.exe", L"obsidian.exe", L"notion.exe", L"anytype.exe",
        L"affine.exe", L"cursor.exe", L"code.exe", L"vscode.exe",
        L"telegram.exe", L"discord.exe", L"slack.exe", L"streamdeck.exe",
        L"obs64.exe", L"audacity.exe"
    };
    for (auto* s : productivity) if (lstrcmpiW(exeLower.c_str(), s) == 0) return IconCategory::PRODUCTIVITY;

    return IconCategory::OTHER_ACTIVE;
}

static bool isManageable(HWND h) {
    if (!IsWindowVisible(h)) return false;
    if (GetWindow(h, GW_OWNER) != nullptr) return false;
    LONG ex = GetWindowLongW(h, GWL_EXSTYLE);
    if (ex & WS_EX_TOOLWINDOW) return false;
    int cloaked = 0;
    DwmGetWindowAttribute(h, DWMWA_CLOAKED, &cloaked, sizeof(cloaked));
    if (cloaked) return false;
    if (GetWindowTextLengthW(h) == 0) return false;
    if (h == g_msgWnd) return false;
    // skip the shell windows themselves
    wchar_t cls[64] = {};
    GetClassNameW(h, cls, 64);
    if (lstrcmpW(cls, L"Shell_TrayWnd") == 0) return false;
    if (lstrcmpW(cls, L"Shell_SecondaryTrayWnd") == 0) return false;
    if (lstrcmpW(cls, L"Progman") == 0) return false;
    if (lstrcmpW(cls, L"WorkerW") == 0) return false;
    return true;
}

// ---------- tray-icon CRUD ----------
// Apply the current m.badgeNum to m.badgedIcon (regenerating from m.rawIcon).
static void rebakeBadge(ManagedWindow& m) {
    if (m.badgedIcon) { DestroyIcon(m.badgedIcon); m.badgedIcon = nullptr; }
    if (m.badgeNum > 0) {
        m.badgedIcon = makeBadgedIcon(m.rawIcon, m.badgeNum,
                                      badgeColorForExe(m.exeNameLower));
    } else {
        m.badgedIcon = CopyIcon(m.rawIcon);
    }
}

static bool trayAdd(ManagedWindow& m) {
    NOTIFYICONDATAW nid{};
    nid.cbSize           = sizeof(nid);
    nid.hWnd             = g_msgWnd;
    nid.uID              = m.uid;
    nid.uFlags           = NIF_ICON | NIF_MESSAGE | NIF_TIP | NIF_SHOWTIP;
    nid.uCallbackMessage = TRAY_CALLBACK_MSG;
    nid.hIcon            = m.badgedIcon;
    std::wstring tip = m.title;
    if (tip.size() > 127) tip.resize(127);
    lstrcpynW(nid.szTip, tip.c_str(), 128);
    if (Shell_NotifyIconW(NIM_ADD, &nid) != TRUE) return false;
    // Opt into NOTIFYICON_VERSION_4 — required for reliable WM_LBUTTONDBLCLK delivery.
    nid.uVersion = NOTIFYICON_VERSION_4;
    Shell_NotifyIconW(NIM_SETVERSION, &nid);
    return true;
}

static void trayModify(ManagedWindow& m) {
    NOTIFYICONDATAW nid{};
    nid.cbSize = sizeof(nid);
    nid.hWnd   = g_msgWnd;
    nid.uID    = m.uid;
    nid.uFlags = NIF_ICON | NIF_TIP | NIF_SHOWTIP;
    nid.hIcon  = m.badgedIcon;
    std::wstring tip = m.title;
    if (tip.size() > 127) tip.resize(127);
    lstrcpynW(nid.szTip, tip.c_str(), 128);
    Shell_NotifyIconW(NIM_MODIFY, &nid);
}

// ---------- pinned-app static launcher icons ----------
// We scan the user's TaskBar pinned-shortcut folder, resolve each .lnk to its
// target exe, extract the icon, and register a tray icon.  Click → ShellExecute
// the .lnk so it launches with whatever args the original pin was set up with.

struct PinnedApp {
    UINT          uid;
    HICON         icon;
    std::wstring  lnkPath;
    std::wstring  displayName;   // shortcut filename minus .lnk
    std::wstring  targetExe;
    bool          launchAtBoot = false;
    std::string   iniKey;        // ASCII-safe key for [boot_apps]
};
static std::vector<PinnedApp> g_pinned;
static std::mutex             g_pinnedMx;

static std::wstring shellLinkResolve(const std::wstring& lnk) {
    IShellLinkW*  sl = nullptr;
    IPersistFile* pf = nullptr;
    std::wstring out;
    HRESULT hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER,
                                  IID_IShellLinkW, (void**)&sl);
    if (FAILED(hr) || !sl) return out;
    if (SUCCEEDED(sl->QueryInterface(IID_IPersistFile, (void**)&pf)) && pf) {
        if (SUCCEEDED(pf->Load(lnk.c_str(), STGM_READ))) {
            wchar_t target[MAX_PATH] = {};
            WIN32_FIND_DATAW fd{};
            sl->Resolve(nullptr, SLR_NO_UI | SLR_NOSEARCH | SLR_NOTRACK | SLR_NOLINKINFO);
            sl->GetPath(target, MAX_PATH, &fd, SLGP_RAWPATH);
            out = target;
        }
        pf->Release();
    }
    sl->Release();
    return out;
}

static HICON extractIconFromExe(const std::wstring& exePath) {
    HICON ico = nullptr;
    ExtractIconExW(exePath.c_str(), 0, nullptr, &ico, 1);
    if (!ico) {
        SHFILEINFOW sfi{};
        SHGetFileInfoW(exePath.c_str(), 0, &sfi, sizeof(sfi),
                       SHGFI_ICON | SHGFI_SMALLICON);
        if (sfi.hIcon) ico = sfi.hIcon;
    }
    return ico;
}

static bool pinnedTrayAdd(PinnedApp& p) {
    NOTIFYICONDATAW nid{};
    nid.cbSize           = sizeof(nid);
    nid.hWnd             = g_msgWnd;
    nid.uID              = p.uid;
    nid.uFlags           = NIF_ICON | NIF_MESSAGE | NIF_TIP | NIF_SHOWTIP;
    nid.uCallbackMessage = TRAY_CALLBACK_MSG;
    nid.hIcon            = p.icon ? p.icon : LoadIcon(nullptr, IDI_APPLICATION);
    std::wstring tip = L"[Pinned] " + p.displayName;
    if (tip.size() > 127) tip.resize(127);
    lstrcpynW(nid.szTip, tip.c_str(), 128);
    if (Shell_NotifyIconW(NIM_ADD, &nid) != TRUE) return false;
    nid.uVersion = NOTIFYICON_VERSION_4;
    Shell_NotifyIconW(NIM_SETVERSION, &nid);
    return true;
}

static void loadPinnedApps() {
    char* appdata = nullptr; size_t sz = 0;
    _dupenv_s(&appdata, &sz, "APPDATA");
    if (!appdata) return;
    std::wstring base;
    {
        wchar_t wbuf[MAX_PATH];
        MultiByteToWideChar(CP_UTF8, 0, appdata, -1, wbuf, MAX_PATH);
        base = wbuf;
    }
    free(appdata);
    base += L"\\Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\TaskBar\\*.lnk";

    WIN32_FIND_DATAW fd{};
    HANDLE h = FindFirstFileW(base.c_str(), &fd);
    if (h == INVALID_HANDLE_VALUE) {
        logf("loadPinnedApps: no pinned dir or empty");
        return;
    }
    UINT nextUid = PINNED_UID_BASE;
    int i = 0;
    do {
        if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) continue;
        // strip filename out of base: replace last '\\*.lnk' segment
        std::wstring dir = base.substr(0, base.find_last_of(L'\\') + 1);
        std::wstring lnk = dir + fd.cFileName;
        std::wstring target = shellLinkResolve(lnk);
        if (target.empty()) continue;

        PinnedApp p{};
        p.uid          = nextUid++;
        p.lnkPath      = lnk;
        p.targetExe    = target;
        p.displayName  = fd.cFileName;
        size_t dot = p.displayName.rfind(L".lnk");
        if (dot != std::wstring::npos) p.displayName.resize(dot);
        HICON raw = extractIconFromExe(target);
        p.icon = makeDimIcon(raw);
        if (raw) DestroyIcon(raw);

        if (pinnedTrayAdd(p)) {
            std::lock_guard<std::mutex> lk(g_pinnedMx);
            g_pinned.push_back(std::move(p));
            ++i;
        } else if (p.icon) {
            DestroyIcon(p.icon);
        }
    } while (FindNextFileW(h, &fd));
    FindClose(h);
    logf("loadPinnedApps: registered %d pinned launchers", i);
}

static const PinnedApp* pinnedFromUid(UINT uid) {
    std::lock_guard<std::mutex> lk(g_pinnedMx);
    for (auto& p : g_pinned) if (p.uid == uid) return &p;
    return nullptr;
}

static void pinnedLaunch(const PinnedApp& p) {
    ShellExecuteW(nullptr, L"open", p.lnkPath.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
}

static void unregisterAllPinned() {
    std::lock_guard<std::mutex> lk(g_pinnedMx);
    for (auto& p : g_pinned) {
        NOTIFYICONDATAW nid{};
        nid.cbSize = sizeof(nid);
        nid.hWnd   = g_msgWnd;
        nid.uID    = p.uid;
        Shell_NotifyIconW(NIM_DELETE, &nid);
        if (p.icon) DestroyIcon(p.icon);
    }
    g_pinned.clear();
}

// ---------- master tray icon (the animated star) ----------
static HICON g_starFrames[STAR_FRAME_COUNT] = {};
static std::atomic<int> g_starFrameIdx{0};
static std::atomic<bool> g_masterRegistered{false};

static void loadStarFrames() {
    HMODULE mod = GetModuleHandleW(nullptr);
    int wantedSize = GetSystemMetrics(SM_CXSMICON);   // 16 on stock DPI
    if (wantedSize <= 0) wantedSize = 16;
    for (int i = 0; i < STAR_FRAME_COUNT; ++i) {
        g_starFrames[i] = (HICON)LoadImageW(
            mod, MAKEINTRESOURCEW(STAR_FRAME_BASE_RES + i),
            IMAGE_ICON, wantedSize, wantedSize, LR_DEFAULTCOLOR);
        if (!g_starFrames[i])
            g_starFrames[i] = (HICON)LoadImageW(
                mod, MAKEINTRESOURCEW(STAR_FRAME_BASE_RES + i),
                IMAGE_ICON, 0, 0, LR_DEFAULTSIZE);
    }
}

static void masterTrayRegister() {
    NOTIFYICONDATAW nid{};
    nid.cbSize           = sizeof(nid);
    nid.hWnd             = g_msgWnd;
    nid.uID              = MASTER_TRAY_UID;
    nid.uFlags           = NIF_ICON | NIF_MESSAGE | NIF_TIP | NIF_SHOWTIP;
    nid.uCallbackMessage = TRAY_CALLBACK_MSG;
    nid.hIcon            = g_starFrames[0] ? g_starFrames[0] : LoadIcon(nullptr, IDI_APPLICATION);
    lstrcpynW(nid.szTip, L"Laurence Trayhost " TRAYSYS_VERSION L" — right-click for menu", 128);
    if (Shell_NotifyIconW(NIM_ADD, &nid) == TRUE) {
        nid.uVersion = NOTIFYICON_VERSION_4;
        Shell_NotifyIconW(NIM_SETVERSION, &nid);
        g_masterRegistered.store(true);
    }
}

static void masterTrayAnimate() {
    if (!g_masterRegistered.load()) return;
    int idx = (g_starFrameIdx.fetch_add(1) + 1) % STAR_FRAME_COUNT;
    HICON h = g_starFrames[idx];
    if (!h) return;
    NOTIFYICONDATAW nid{};
    nid.cbSize = sizeof(nid);
    nid.hWnd   = g_msgWnd;
    nid.uID    = MASTER_TRAY_UID;
    nid.uFlags = NIF_ICON;
    nid.hIcon  = h;
    Shell_NotifyIconW(NIM_MODIFY, &nid);
}

static void masterTrayUnregister() {
    if (!g_masterRegistered.load()) return;
    NOTIFYICONDATAW nid{};
    nid.cbSize = sizeof(nid);
    nid.hWnd   = g_msgWnd;
    nid.uID    = MASTER_TRAY_UID;
    Shell_NotifyIconW(NIM_DELETE, &nid);
    g_masterRegistered.store(false);
}

// ---------- balloon helper ----------
// Show a one-shot Windows toast/balloon, attached to one of our existing tray icons.
static void trayBalloon(UINT uid, const wchar_t* title, const wchar_t* body) {
    NOTIFYICONDATAW nid{};
    nid.cbSize = sizeof(nid);
    nid.hWnd   = g_msgWnd;
    nid.uID    = uid;
    nid.uFlags = NIF_INFO;
    nid.dwInfoFlags = NIIF_INFO | NIIF_NOSOUND;
    lstrcpynW(nid.szInfoTitle, title, 64);
    lstrcpynW(nid.szInfo,      body,  256);
    Shell_NotifyIconW(NIM_MODIFY, &nid);
}

static void trayDelete(ManagedWindow& m) {
    NOTIFYICONDATAW nid{};
    nid.cbSize = sizeof(nid);
    nid.hWnd   = g_msgWnd;
    nid.uID    = m.uid;
    Shell_NotifyIconW(NIM_DELETE, &nid);
}

// ---------- window-list refresh ----------
static std::atomic<int> g_lastEnumTotal{0};
static std::atomic<int> g_lastEnumKept{0};

static BOOL CALLBACK enumProc(HWND h, LPARAM lp) {
    auto* seen = reinterpret_cast<std::unordered_set<HWND>*>(lp);
    g_lastEnumTotal.fetch_add(1);
    if (isManageable(h)) {
        seen->insert(h);
        g_lastEnumKept.fetch_add(1);
    }
    return TRUE;
}

static void refreshWindows() {
    auto t0 = std::chrono::steady_clock::now();
    g_lastEnumTotal.store(0);
    g_lastEnumKept.store(0);
    std::unordered_set<HWND> seen;
    EnumWindows(enumProc, reinterpret_cast<LPARAM>(&seen));

    int added = 0, removed = 0;
    {
        std::lock_guard<std::mutex> lk(g_windowsMx);

        // remove gone
        for (auto it = g_windows.begin(); it != g_windows.end(); ) {
            if (seen.find(it->first) == seen.end()) {
                if (it->second.hiddenFromTaskbar)
                    taskbarListSetVisible(it->second.hwnd, true);  // restore taskbar slot
                trayDelete(it->second);
                if (it->second.rawIcon)    DestroyIcon(it->second.rawIcon);
                if (it->second.badgedIcon) DestroyIcon(it->second.badgedIcon);
                g_uidToHwnd.erase(it->second.uid);
                it = g_windows.erase(it);
                ++removed;
            } else ++it;
        }

        // add new + update titles
        for (HWND h : seen) {
            auto it = g_windows.find(h);
            if (it == g_windows.end()) {
                ManagedWindow m;
                m.hwnd         = h;
                m.uid          = g_nextUid++;
                m.rawIcon      = copyIconForHwnd(h);
                m.badgedIcon   = CopyIcon(m.rawIcon);  // initial: unbadged
                m.title        = titleOf(h);
                std::wstring exe = exeBasenameForHwnd(h);
                m.exeName      = exe;
                m.exeNameLower = exe;
                std::transform(m.exeNameLower.begin(), m.exeNameLower.end(),
                               m.exeNameLower.begin(), ::towlower);
                if (trayAdd(m)) {
                    g_uidToHwnd[m.uid] = h;
                    auto [ins, _] = g_windows.emplace(h, std::move(m));
                    // hide window from real taskbar per the global flags
                    bool wantHide =
                        g_hideAllFromTaskbar.load() ||
                        (g_hideBrowsersFromTaskbar.load() && isBrowserExe(ins->second.exeNameLower));
                    if (wantHide) {
                        taskbarListSetVisible(h, false);
                        ins->second.hiddenFromTaskbar = true;
                    }
                    ++added;
                } else {
                    if (m.rawIcon)    DestroyIcon(m.rawIcon);
                    if (m.badgedIcon) DestroyIcon(m.badgedIcon);
                    logf("trayAdd failed for hwnd=%p uid=%u", h, (unsigned)m.uid);
                }
            } else {
                // Title may have changed — update internal cache for /api/state,
                // but do NOT push to Shell_NotifyIcon. NIM_MODIFY on every browser
                // tab retitle causes the visible tray to flicker / reflow. The
                // tooltip is set once at NIM_ADD; live titles are visible in the
                // window's own title bar.
                std::wstring t = titleOf(h);
                if (t != it->second.title) it->second.title = std::move(t);
            }
        }

        // ----- recompute per-exe badge indices and re-bake icons where they changed -----
        std::unordered_map<std::wstring, std::vector<ManagedWindow*>> byExe;
        for (auto& kv : g_windows) byExe[kv.second.exeNameLower].push_back(&kv.second);

        for (auto& kv : byExe) {
            auto& vec = kv.second;
            // stable order by uid so badge numbers don't reshuffle on every tick
            std::sort(vec.begin(), vec.end(),
                      [](ManagedWindow* a, ManagedWindow* b){ return a->uid < b->uid; });
            int n = (int)vec.size();
            for (int i = 0; i < n; ++i) {
                int wantBadge = (n > 1) ? (i + 1) : 0;
                if (vec[i]->badgeNum != wantBadge) {
                    vec[i]->badgeNum = wantBadge;
                    rebakeBadge(*vec[i]);
                    trayModify(*vec[i]);
                }
            }
        }
    }

    auto t1 = std::chrono::steady_clock::now();
    double ms = std::chrono::duration<double, std::milli>(t1 - t0).count();
    g_lastRefreshMs.store(nowMs());
    g_lastRefreshDuration.store(ms);

    BenchSample s;
    s.ts_ms      = nowMs();
    s.refresh_ms = ms;
    s.icons      = (int)g_windows.size();
    s.rss_kb     = rssKB();
    s.added      = added;
    s.removed    = removed;
    benchAppend(s);
    logf("refresh: enum_total=%d enum_kept=%d seen=%zu icons=%d added=%d removed=%d in %.1fms",
         g_lastEnumTotal.load(), g_lastEnumKept.load(),
         seen.size(), (int)g_windows.size(), added, removed, ms);
}

// ---------- focus / close from tray ----------
static void focusHwnd(HWND h) {
    if (!IsWindow(h)) return;
    if (IsIconic(h)) ShowWindow(h, SW_RESTORE);
    DWORD fg = GetWindowThreadProcessId(GetForegroundWindow(), nullptr);
    DWORD me = GetCurrentThreadId();
    if (fg != me) AttachThreadInput(fg, me, TRUE);
    SetForegroundWindow(h);
    BringWindowToTop(h);
    if (fg != me) AttachThreadInput(fg, me, FALSE);
}

static HWND hwndFromUid(UINT uid) {
    std::lock_guard<std::mutex> lk(g_windowsMx);
    auto it = g_uidToHwnd.find(uid);
    return it == g_uidToHwnd.end() ? nullptr : it->second;
}

// ---- forward decls for menu actions ----
static void actionMasterReset();
static void actionCloseAll();
static void actionExportJson();
static void actionOpenConsole();
static void actionOpenSettings();

// ---------- profiles (5 named slots, persisted in INI) ----------
// A profile snapshots the set of currently-running tracked apps (by exe).
// Apply = launch any matching pinned-app .lnk that isn't already running.
// Stored in [profile_N] sections: name=..., apps=chrome.exe,slack.exe,...

struct Profile {
    int          slot;
    std::string  name;
    std::string  apps;   // comma-separated exe basenames (ascii-lowered)
};

static std::string defaultProfileName(int slot) {
    char b[32]; snprintf(b, sizeof(b), "Profile %d", slot); return b;
}

static Profile profileLoad(int slot) {
    Profile p; p.slot = slot;
    char section[24]; snprintf(section, sizeof(section), "profile_%d", slot);
    char buf[1024];
    GetPrivateProfileStringA(section, "name",
        defaultProfileName(slot).c_str(), buf, sizeof(buf), g_iniPath);
    p.name = buf;
    GetPrivateProfileStringA(section, "apps", "", buf, sizeof(buf), g_iniPath);
    p.apps = buf;
    return p;
}

static void profileSave(const Profile& p) {
    char section[24]; snprintf(section, sizeof(section), "profile_%d", p.slot);
    WritePrivateProfileStringA(section, "name", p.name.c_str(), g_iniPath);
    WritePrivateProfileStringA(section, "apps", p.apps.c_str(), g_iniPath);
}

static std::string captureRunningExes() {
    std::string out;
    std::lock_guard<std::mutex> lk(g_windowsMx);
    std::unordered_set<std::string> seen;
    for (auto& kv : g_windows) {
        std::string s; s.reserve(kv.second.exeNameLower.size());
        for (wchar_t c : kv.second.exeNameLower) s.push_back((char)(c & 0x7F));
        if (seen.insert(s).second) {
            if (!out.empty()) out += ",";
            out += s;
        }
    }
    return out;
}

static void profileSnapshot(int slot) {
    Profile p = profileLoad(slot);
    p.apps = captureRunningExes();
    profileSave(p);
    wchar_t wname[256] = {};
    MultiByteToWideChar(CP_UTF8, 0, p.name.c_str(), -1, wname, 256);
    wchar_t body[400];
    swprintf(body, 400, L"Snapshotted to \"%ls\". Right-click star → Profiles → Apply to restore.", wname);
    trayBalloon(MASTER_TRAY_UID, L"Profile saved", body);
}

static void profileApply(int slot) {
    Profile p = profileLoad(slot);
    if (p.apps.empty()) {
        trayBalloon(MASTER_TRAY_UID, L"Profile empty", L"Nothing to launch — snapshot one first.");
        return;
    }
    int launched = 0;
    std::string token;
    auto tryLaunch = [&](const std::string& exe) {
        std::lock_guard<std::mutex> lk(g_pinnedMx);
        for (auto& pa : g_pinned) {
            std::wstring tw = pa.targetExe;
            size_t bs = tw.find_last_of(L"\\/");
            std::wstring base = (bs == std::wstring::npos) ? tw : tw.substr(bs + 1);
            std::transform(base.begin(), base.end(), base.begin(), ::towlower);
            std::string baseAscii; baseAscii.reserve(base.size());
            for (wchar_t c : base) baseAscii.push_back((char)(c & 0x7F));
            if (baseAscii == exe) {
                ShellExecuteW(nullptr, L"open", pa.lnkPath.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
                ++launched;
                return;
            }
        }
    };
    for (size_t i = 0; i <= p.apps.size(); ++i) {
        if (i == p.apps.size() || p.apps[i] == ',') {
            if (!token.empty()) tryLaunch(token);
            token.clear();
        } else token.push_back(p.apps[i]);
    }
    wchar_t body[200];
    swprintf(body, 200, L"Launched %d of %zu apps.", launched,
             std::count(p.apps.begin(), p.apps.end(), ',') + (p.apps.empty() ? 0 : 1));
    trayBalloon(MASTER_TRAY_UID, L"Profile applied", body);
}

static void profileRename(int slot) {
    wchar_t url[80];
    swprintf(url, 80, L"http://127.0.0.1:%d/#profile-%d", HTTP_PORT, slot);
    ShellExecuteW(nullptr, L"open", url, nullptr, nullptr, SW_SHOWNORMAL);
}

// ---------- Group / Stack & Lucky Dip ----------
// Tear down every per-window + pinned tray icon, then re-add them in a new
// order. Windows shows tray icons in the order we register them (until the
// user manually drags one).
//
//   mode = "group"  -> sort by category (browsers first, pinned last)
//   mode = "random" -> Fisher-Yates shuffle
//   mode = "uid"    -> stable original (uid ascending), used for "Reset order"

enum class RearrangeMode { GROUP, RANDOM, UID };

static void rearrangeIcons(RearrangeMode mode) {
    struct Item { UINT uid; bool pinned; std::wstring exeLower; };
    std::vector<Item> items;

    {
        std::lock_guard<std::mutex> lk(g_pinnedMx);
        for (auto& p : g_pinned) items.push_back({p.uid, true, L""});
    }
    {
        std::lock_guard<std::mutex> lk(g_windowsMx);
        for (auto& kv : g_windows) items.push_back({kv.second.uid, false, kv.second.exeNameLower});
    }
    if (items.empty()) return;

    if (mode == RearrangeMode::RANDOM) {
        std::random_device rd; std::mt19937 g(rd());
        std::shuffle(items.begin(), items.end(), g);
    } else if (mode == RearrangeMode::GROUP) {
        std::sort(items.begin(), items.end(), [](const Item& a, const Item& b){
            int ac = a.pinned ? (int)IconCategory::PINNED : (int)categorizeExe(a.exeLower);
            int bc = b.pinned ? (int)IconCategory::PINNED : (int)categorizeExe(b.exeLower);
            if (ac != bc) return ac < bc;
            return a.uid < b.uid;
        });
    } else {
        std::sort(items.begin(), items.end(),
                  [](const Item& a, const Item& b){ return a.uid < b.uid; });
    }

    // Tear down (NIM_DELETE for every icon in our managed set).
    for (auto& it : items) {
        NOTIFYICONDATAW nid{};
        nid.cbSize = sizeof(nid);
        nid.hWnd   = g_msgWnd;
        nid.uID    = it.uid;
        Shell_NotifyIconW(NIM_DELETE, &nid);
    }

    // Re-add in the new order.
    int reAddedPinned = 0, reAddedWin = 0;
    for (auto& it : items) {
        if (it.pinned) {
            std::lock_guard<std::mutex> lk(g_pinnedMx);
            for (auto& p : g_pinned) if (p.uid == it.uid) { pinnedTrayAdd(p); ++reAddedPinned; break; }
        } else {
            std::lock_guard<std::mutex> lk(g_windowsMx);
            for (auto& kv : g_windows) if (kv.second.uid == it.uid) { trayAdd(kv.second); ++reAddedWin; break; }
        }
    }

    const wchar_t* label = mode == RearrangeMode::GROUP  ? L"Grouped"
                        :  mode == RearrangeMode::RANDOM ? L"Lucky dip"
                        :                                  L"Reset order";
    wchar_t body[200];
    swprintf(body, 200, L"%d windows + %d pinned icons re-registered.", reAddedWin, reAddedPinned);
    trayBalloon(MASTER_TRAY_UID, label, body);
    logf("rearrangeIcons mode=%d: %d windows + %d pinned re-registered",
         (int)mode, reAddedWin, reAddedPinned);
}

static void showContextMenuFor(UINT uid) {
    POINT pt; GetCursorPos(&pt);
    SetForegroundWindow(g_msgWnd);

    HMENU m = CreatePopupMenu();
    if (uid == MASTER_TRAY_UID) {
        // ----- profiles submenu -----
        HMENU snapMenu  = CreatePopupMenu();
        HMENU applyMenu = CreatePopupMenu();
        HMENU nameMenu  = CreatePopupMenu();
        for (int i = 1; i <= PROFILE_SLOTS; ++i) {
            Profile p = profileLoad(i);
            wchar_t label[256]; size_t apps = 0;
            for (char c : p.apps) if (c == ',') ++apps;
            if (!p.apps.empty()) ++apps;
            int n = MultiByteToWideChar(CP_UTF8, 0, p.name.c_str(), -1, label, 256);
            (void)n;
            // Snapshot label: "Slot N — current name"
            wchar_t snapLbl[300]; swprintf(snapLbl, 300, L"Slot %d — %ls", i, label);
            AppendMenuW(snapMenu,  MF_STRING, 1100 + i, snapLbl);
            // Apply label: "Profile name (X apps)"
            wchar_t appLbl[300]; swprintf(appLbl, 300, L"%ls  (%zu apps)", label, apps);
            AppendMenuW(applyMenu, apps == 0 ? (MF_STRING | MF_GRAYED) : MF_STRING,
                        1200 + i, appLbl);
            wchar_t renLbl[300]; swprintf(renLbl, 300, L"Rename slot %d (%ls)…", i, label);
            AppendMenuW(nameMenu,  MF_STRING, 1300 + i, renLbl);
        }
        HMENU profMenu = CreatePopupMenu();
        AppendMenuW(profMenu, MF_POPUP, (UINT_PTR)snapMenu,  L"Snapshot current →");
        AppendMenuW(profMenu, MF_POPUP, (UINT_PTR)applyMenu, L"Apply profile →");
        AppendMenuW(profMenu, MF_POPUP, (UINT_PTR)nameMenu,  L"Rename →");
        AppendMenuW(profMenu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(profMenu, MF_STRING, 1400, L"Group / stack icons (browsers → search → system → other → pinned)");
        AppendMenuW(profMenu, MF_STRING, 1401, L"Lucky dip (random shuffle)");
        AppendMenuW(profMenu, MF_STRING, 1402, L"Reset order (original)");

        // ----- master star icon menu -----
        bool autostart = autostartIsEnabled();
        AppendMenuW(m, MF_STRING | (autostart ? MF_CHECKED : MF_UNCHECKED),
                    1010, L"Auto-start with Windows");
        AppendMenuW(m, MF_POPUP, (UINT_PTR)profMenu, L"Profiles");
        AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(m, MF_STRING, 1000, L"Master reset");
        AppendMenuW(m, MF_STRING, 1001, L"Close all windows");
        AppendMenuW(m, MF_STRING, 1002, L"Export to JSON…");
        AppendMenuW(m, MF_STRING, 1003, L"Open console (log)");
        AppendMenuW(m, MF_STRING, 1004, L"Settings…");
        AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(m, MF_STRING, 1099, L"Quit Laurence Trayhost");

        int cmd = TrackPopupMenu(m, TPM_RETURNCMD | TPM_RIGHTBUTTON, pt.x, pt.y, 0, g_msgWnd, nullptr);
        DestroyMenu(m);
        PostMessageW(g_msgWnd, WM_NULL, 0, 0);

        if (cmd == 1010) {
            autostartSet(!autostart);
            trayBalloon(MASTER_TRAY_UID,
                !autostart ? L"Auto-start enabled" : L"Auto-start disabled",
                !autostart ? L"Laurence Trayhost will launch with Windows."
                           : L"Laurence Trayhost will not launch with Windows.");
        }
        else if (cmd >= 1100 && cmd < 1100 + PROFILE_SLOTS + 1) profileSnapshot(cmd - 1100);
        else if (cmd >= 1200 && cmd < 1200 + PROFILE_SLOTS + 1) profileApply(cmd - 1200);
        else if (cmd >= 1300 && cmd < 1300 + PROFILE_SLOTS + 1) profileRename(cmd - 1300);
        else if (cmd == 1400) rearrangeIcons(RearrangeMode::GROUP);
        else if (cmd == 1401) rearrangeIcons(RearrangeMode::RANDOM);
        else if (cmd == 1402) rearrangeIcons(RearrangeMode::UID);
        else switch (cmd) {
            case 1000: actionMasterReset();   break;
            case 1001: actionCloseAll();      break;
            case 1002: actionExportJson();    break;
            case 1003: actionOpenConsole();   break;
            case 1004: actionOpenSettings();  break;
            case 1099: PostMessageW(g_msgWnd, WM_CLOSE, 0, 0); break;
        }
        return;
    }

    // ----- per-window icon menu -----
    HWND target = hwndFromUid(uid);
    if (!target) { DestroyMenu(m); return; }
    AppendMenuW(m, MF_STRING, 100, L"Focus");
    AppendMenuW(m, MF_STRING, 102, IsIconic(target) ? L"Restore" : L"Minimise");
    AppendMenuW(m, MF_STRING, 101, L"Close window");
    AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(m, MF_STRING, 200, L"Open settings (browser)");
    AppendMenuW(m, MF_STRING, 999, L"Quit Laurence Trayhost");

    int cmd = TrackPopupMenu(m, TPM_RETURNCMD | TPM_RIGHTBUTTON, pt.x, pt.y, 0, g_msgWnd, nullptr);
    DestroyMenu(m);
    PostMessageW(g_msgWnd, WM_NULL, 0, 0);

    if (cmd == 100) focusHwnd(target);
    else if (cmd == 101) PostMessageW(target, WM_CLOSE, 0, 0);
    else if (cmd == 102) {
        if (IsIconic(target)) { ShowWindow(target, SW_RESTORE); focusHwnd(target); }
        else                  { ShowWindow(target, SW_MINIMIZE); }
    }
    else if (cmd == 200) actionOpenSettings();
    else if (cmd == 999) PostMessageW(g_msgWnd, WM_CLOSE, 0, 0);
}

// ---------- master-menu actions ----------
static std::string buildStateJson();   // forward decl

static void actionMasterReset() {
    // Tear down all per-window tray icons + their hidden-from-taskbar state,
    // then re-enumerate from scratch. Master star icon is left untouched.
    {
        std::lock_guard<std::mutex> lk(g_windowsMx);
        for (auto& kv : g_windows) {
            if (kv.second.hiddenFromTaskbar)
                taskbarListSetVisibleDirect(kv.second.hwnd, true);
            trayDelete(kv.second);
            if (kv.second.rawIcon)    DestroyIcon(kv.second.rawIcon);
            if (kv.second.badgedIcon) DestroyIcon(kv.second.badgedIcon);
        }
        g_windows.clear();
        g_uidToHwnd.clear();
    }
    refreshWindows();
    trayBalloon(MASTER_TRAY_UID, L"Master reset",
                L"All tray icons cleared and re-enumerated.");
}

static void actionCloseAll() {
    if (MessageBoxW(nullptr,
        L"Send WM_CLOSE to every tracked window?\n\nUnsaved work in any of them may be lost.",
        L"Laurence Trayhost — close all",
        MB_YESNO | MB_ICONWARNING | MB_TOPMOST | MB_SETFOREGROUND) != IDYES) return;
    std::vector<HWND> targets;
    {
        std::lock_guard<std::mutex> lk(g_windowsMx);
        for (auto& kv : g_windows) targets.push_back(kv.second.hwnd);
    }
    for (HWND h : targets) PostMessageW(h, WM_CLOSE, 0, 0);
    trayBalloon(MASTER_TRAY_UID, L"Close all",
                L"Sent WM_CLOSE to all tracked windows.");
}

static void actionExportJson() {
    // Write a snapshot of the current state to %USERPROFILE%\Desktop\trayhost-export-<ts>.json
    char ts[32]; SYSTEMTIME st; GetLocalTime(&st);
    snprintf(ts, sizeof(ts), "%04d%02d%02d-%02d%02d%02d",
             st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
    char desktop[MAX_PATH] = {};
    char* up = nullptr; size_t sz = 0;
    _dupenv_s(&up, &sz, "USERPROFILE");
    if (up) { snprintf(desktop, sizeof(desktop), "%s\\Desktop\\trayhost-export-%s.json", up, ts); free(up); }
    else    { snprintf(desktop, sizeof(desktop), "trayhost-export-%s.json", ts); }
    std::string body = buildStateJson();
    FILE* f = fopen(desktop, "wb");
    if (!f) {
        trayBalloon(MASTER_TRAY_UID, L"Export failed", L"Couldn't write to Desktop.");
        return;
    }
    fwrite(body.data(), 1, body.size(), f);
    fclose(f);
    wchar_t wpath[MAX_PATH]; MultiByteToWideChar(CP_UTF8, 0, desktop, -1, wpath, MAX_PATH);
    wchar_t msg[MAX_PATH + 64];
    swprintf(msg, MAX_PATH + 64, L"Saved snapshot to %ls", wpath);
    trayBalloon(MASTER_TRAY_UID, L"Exported state", msg);
}

static void actionOpenConsole() {
    // Open the live log in the default text viewer (typically Notepad).
    wchar_t wlog[MAX_PATH];
    MultiByteToWideChar(CP_UTF8, 0, g_logPath, -1, wlog, MAX_PATH);
    ShellExecuteW(nullptr, L"open", L"notepad.exe", wlog, nullptr, SW_SHOWNORMAL);
}

static void actionOpenSettings() {
    wchar_t url[64];
    swprintf(url, 64, L"http://127.0.0.1:%d/", HTTP_PORT);
    ShellExecuteW(nullptr, L"open", url, nullptr, nullptr, SW_SHOWNORMAL);
}

// ---------- HTTP server (tiny, single-threaded accept-loop, blocking) ----------
static const char* HTTP_INDEX_HTML = R"HTML(<!doctype html>
<meta charset="utf-8">
<title>TraySys</title>
<style>
:root{color-scheme:dark;}
body{font-family:Segoe UI,system-ui,sans-serif;background:#0a0a0a;color:#e6e6e6;margin:0;padding:24px;max-width:920px;}
h1{font-weight:600;font-size:20px;letter-spacing:.5px;margin:0 0 4px}
.sub{color:#888;font-size:12px;margin-bottom:24px}
table{width:100%;border-collapse:collapse;font-size:13px}
th{text-align:left;color:#888;font-weight:500;padding:6px 10px;border-bottom:1px solid #222}
td{padding:8px 10px;border-bottom:1px solid #161616;vertical-align:top}
tr:hover td{background:#141414}
.uid{color:#666;font-variant-numeric:tabular-nums;width:48px}
.exe{color:#9ad}
.title{color:#e6e6e6}
.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:16px 0 24px}
.stat{background:#111;border:1px solid #222;padding:10px 12px;border-radius:6px}
.stat .v{font-size:18px;font-weight:600}
.stat .k{color:#888;font-size:11px;text-transform:uppercase;letter-spacing:.5px}
button{background:#1a1a1a;color:#e6e6e6;border:1px solid #2a2a2a;padding:6px 12px;border-radius:4px;cursor:pointer;margin-right:6px}
button:hover{background:#222}
</style>
<h1>TraySys <span style="color:#666;font-weight:400">v)HTML" "0.1.0" R"HTML(</span></h1>
<div class="sub">Every active window, represented in your real system tray.</div>
<div class="stats" id="stats"></div>
<div>
  <button onclick="r()">Refresh now</button>
  <button onclick="q()">Quit Laurence Trayhost</button>
  <label style="margin-left:18px;color:#bbb">
    <input type="checkbox" id="autostart" onchange="toggleAutostart(this)"> Auto-start with Windows
  </label>
</div>

<h2 style="margin-top:32px;font-size:14px;color:#888;text-transform:uppercase;letter-spacing:.7px">Engine</h2>
<p style="color:#888;font-size:12px;margin:0 0 12px;line-height:1.5">
How Trayhost learns about window changes.  Push uses two independent OS push channels (RegisterShellHookWindow + WinEventHook) and never polls. Hybrid adds a slow safety-net poll. Polling is the legacy fallback.
</p>
<div id="engine" style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:24px"></div>

<h2 style="margin-top:8px;font-size:14px;color:#888;text-transform:uppercase;letter-spacing:.7px">Timings (live — saved instantly)</h2>
<p id="tuning-blurb" style="color:#888;font-size:12px;margin:0 0 14px;line-height:1.5"></p>
<div id="tuning" style="display:grid;grid-template-columns:200px 1fr 80px;gap:10px 16px;align-items:center;margin-bottom:24px"></div>

<h2 style="margin-top:32px;font-size:14px;color:#888;text-transform:uppercase;letter-spacing:.7px">Profiles</h2>
<div id="profiles"></div>

<h2 style="margin-top:32px;font-size:14px;color:#888;text-transform:uppercase;letter-spacing:.7px">Active windows</h2>
<table id="t"><thead><tr><th class="uid">UID</th><th>Exe</th><th>Title</th></tr></thead><tbody></tbody></table>
<script>
async function load(){
  const r = await fetch('/api/state').then(x=>x.json());
  const s = document.getElementById('stats');
  s.innerHTML = `
    <div class="stat"><div class="k">Tray icons</div><div class="v">${r.icons}</div></div>
    <div class="stat"><div class="k">Last refresh</div><div class="v">${r.refresh_ms.toFixed(1)} ms</div></div>
    <div class="stat"><div class="k">RSS</div><div class="v">${(r.rss_kb/1024).toFixed(1)} MB</div></div>
    <div class="stat"><div class="k">Uptime</div><div class="v">${(r.uptime_s).toFixed(0)} s</div></div>`;
  document.querySelector('#t tbody').innerHTML =
    r.windows.map(w=>`<tr><td class="uid">${w.uid}</td><td class="exe">${w.exe}</td><td class="title">${w.title}</td></tr>`).join('');
}
async function loadAutostart(){
  const r = await fetch('/api/autostart').then(x=>x.json());
  document.getElementById('autostart').checked = !!r.enabled;
}
async function toggleAutostart(cb){
  await fetch('/api/autostart',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:'enabled='+(cb.checked?1:0)});
}
async function loadProfiles(){
  const r = await fetch('/api/profiles').then(x=>x.json());
  const wrap = document.getElementById('profiles');
  wrap.innerHTML = r.profiles.map(p=>`
    <div id="profile-${p.slot}" style="background:#0f0f0f;border:1px solid #222;border-radius:6px;padding:10px 12px;margin:8px 0;display:flex;align-items:center;gap:12px;flex-wrap:wrap">
      <div style="color:#666;font-variant-numeric:tabular-nums;width:28px">#${p.slot}</div>
      <input type="text" value="${p.name.replace(/"/g,'&quot;')}" data-slot="${p.slot}"
             onblur="renameProfile(this)" onkeydown="if(event.key==='Enter')this.blur()"
             style="background:#1a1a1a;color:#e6e6e6;border:1px solid #2a2a2a;padding:6px 10px;border-radius:4px;flex:1;min-width:160px"/>
      <div style="color:#888;font-size:12px;flex:2;min-width:240px">${p.apps || '<em style=color:#444>empty</em>'}</div>
      <button onclick="snapProfile(${p.slot})">Snapshot</button>
      <button onclick="applyProfile(${p.slot})" ${p.apps?'':'disabled style=opacity:.4;cursor:default'}>Apply</button>
    </div>`).join('');
}
async function snapProfile(slot){
  await fetch('/api/profile/'+slot+'/snapshot',{method:'POST'});
  loadProfiles();
}
async function applyProfile(slot){
  await fetch('/api/profile/'+slot+'/apply',{method:'POST'});
}
async function renameProfile(input){
  const slot = input.dataset.slot;
  const name = encodeURIComponent(input.value);
  await fetch('/api/profile/'+slot+'/rename',{
    method:'POST',
    headers:{'Content-Type':'application/x-www-form-urlencoded'},
    body:'name='+name
  });
  loadProfiles();
}
async function r(){ await fetch('/api/refresh',{method:'POST'}); load(); }
async function q(){ if(confirm('Quit Laurence Trayhost and remove all tray icons?')) await fetch('/api/quit',{method:'POST'}); }
async function loadEngine(){
  const r = await fetch('/api/engine').then(x=>x.json());
  const modes = [
    {v:0, k:'PUSH',   d:'Pure push (recommended)'},
    {v:1, k:'HYBRID', d:'Push + safety-net poll'},
    {v:2, k:'POLL',   d:'Legacy polling only'},
  ];
  document.getElementById('engine').innerHTML = modes.map(m=>`
    <button onclick="setEngine(${m.v})"
            style="background:${m.v===r.mode?'#1f3a5f':'#1a1a1a'};color:#e6e6e6;border:1px solid ${m.v===r.mode?'#3b6ea3':'#2a2a2a'};padding:8px 14px;border-radius:4px;cursor:pointer;text-align:left">
      <div style="font-weight:600">${m.k}</div>
      <div style="font-size:11px;color:#888">${m.d}</div>
    </button>`).join('');
  document.body.dataset.engine = r.mode;
  // Hide the safety-net slider when engine is PUSH (it's irrelevant).
  const blurb = document.getElementById('tuning-blurb');
  blurb.textContent = r.mode === 0
    ? 'Push engine — no polling. Tweak the debounce only if you see icon-add bursts feel jittery.'
    : 'Polling enabled. Lower the poll for snappier updates; raise it to reduce flicker.';
}
async function setEngine(v){
  await fetch('/api/engine',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:'mode='+v});
  loadEngine(); loadTuning();
}
async function loadTuning(){
  const eng = parseInt(document.body.dataset.engine || '0');
  const r = await fetch('/api/tuning').then(x=>x.json());
  const sliders = [
    eng !== 0 ? {key:'refresh_interval_ms', label:'Safety-net poll',  min:500, max:30000, step:250, val:r.refresh_interval_ms, suf:'ms'} : null,
    {key:'debounce_ms',         label:'Event debounce',   min:20,  max:1000,  step:10,  val:r.debounce_ms,         suf:'ms'},
    {key:'star_anim_ms',        label:'Star animation',   min:0,   max:1500,  step:20,  val:r.star_anim_ms,        suf:'ms (0 = off)'},
  ].filter(Boolean);
  const wrap = document.getElementById('tuning');
  wrap.innerHTML = sliders.map(s=>`
    <label style="color:#bbb">${s.label}</label>
    <input type="range" min="${s.min}" max="${s.max}" step="${s.step}" value="${s.val}"
           data-key="${s.key}" oninput="onSlide(this)" onchange="saveTuning()"
           style="background:#1a1a1a">
    <span data-suffix="${s.key}" style="color:#9ad;font-variant-numeric:tabular-nums">${s.val}${s.suf}</span>
  `).join('');
}
function onSlide(input){
  const key = input.dataset.key;
  const span = document.querySelector('[data-suffix="'+key+'"]');
  span.textContent = input.value + (key==='star_anim_ms' ? 'ms (0 = off)' : 'ms');
}
async function saveTuning(){
  const body = Array.from(document.querySelectorAll('#tuning input')).
    map(i=>i.dataset.key+'='+i.value).join('&');
  await fetch('/api/tuning',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body});
}
load(); loadAutostart(); loadProfiles(); loadEngine().then(loadTuning);
setInterval(load, 1500);
</script>
)HTML";

static long long g_startMs = 0;

static std::string utf8(const std::wstring& w) {
    if (w.empty()) return {};
    int n = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), nullptr, 0, nullptr, nullptr);
    std::string s(n, 0);
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), s.data(), n, nullptr, nullptr);
    return s;
}
static std::string jsonEscape(const std::string& s) {
    std::string o; o.reserve(s.size() + 8);
    for (char c : s) {
        switch (c) {
            case '"':  o += "\\\""; break;
            case '\\': o += "\\\\"; break;
            case '\n': o += "\\n";  break;
            case '\r': o += "\\r";  break;
            case '\t': o += "\\t";  break;
            default:
                if ((unsigned char)c < 0x20) { char b[8]; sprintf_s(b, "\\u%04x", c); o += b; }
                else o += c;
        }
    }
    return o;
}
static std::string buildStateJson() {
    std::string j = "{";
    {
        std::lock_guard<std::mutex> lk(g_windowsMx);
        char buf[256];
        sprintf_s(buf, "\"icons\":%zu,", g_windows.size());                        j += buf;
        sprintf_s(buf, "\"refresh_ms\":%.3f,", g_lastRefreshDuration.load());      j += buf;
        sprintf_s(buf, "\"rss_kb\":%zu,",      rssKB());                           j += buf;
        sprintf_s(buf, "\"uptime_s\":%.3f,",   (nowMs() - g_startMs) / 1000.0);    j += buf;
        j += "\"windows\":[";
        bool first = true;
        for (auto& kv : g_windows) {
            if (!first) j += ",";
            first = false;
            sprintf_s(buf, "{\"uid\":%u,\"exe\":\"", (unsigned)kv.second.uid);    j += buf;
            j += jsonEscape(utf8(kv.second.exeName));
            j += "\",\"title\":\"";
            j += jsonEscape(utf8(kv.second.title));
            j += "\"}";
        }
        j += "]";
    }
    j += "}";
    return j;
}

static void httpSendAll(SOCKET s, const char* data, size_t len) {
    while (len) {
        int n = send(s, data, (int)len, 0);
        if (n <= 0) break;
        data += n; len -= n;
    }
}
static void httpRespond(SOCKET s, const char* status, const char* contentType, const std::string& body) {
    char hdr[512];
    int n = sprintf_s(hdr,
        "HTTP/1.1 %s\r\nContent-Type: %s\r\nContent-Length: %zu\r\nCache-Control: no-store\r\nConnection: close\r\n\r\n",
        status, contentType, body.size());
    httpSendAll(s, hdr, n);
    httpSendAll(s, body.data(), body.size());
}
static void httpHandle(SOCKET s) {
    char buf[4096] = {};
    int n = recv(s, buf, sizeof(buf) - 1, 0);
    if (n <= 0) { closesocket(s); return; }

    char method[8] = {}, path[256] = {};
    sscanf_s(buf, "%7s %255s", method, (unsigned)sizeof(method), path, (unsigned)sizeof(path));

    if (strcmp(path, "/") == 0) {
        httpRespond(s, "200 OK", "text/html; charset=utf-8", HTTP_INDEX_HTML);
    } else if (strcmp(path, "/api/state") == 0) {
        httpRespond(s, "200 OK", "application/json", buildStateJson());
    } else if (strcmp(path, "/api/refresh") == 0 && strcmp(method, "POST") == 0) {
        PostMessageW(g_msgWnd, WM_APP + 2, 0, 0);
        httpRespond(s, "200 OK", "application/json", "{\"ok\":true}");
    } else if (strcmp(path, "/api/quit") == 0 && strcmp(method, "POST") == 0) {
        httpRespond(s, "200 OK", "application/json", "{\"ok\":true}");
        PostMessageW(g_msgWnd, WM_CLOSE, 0, 0);
    } else if (strcmp(path, "/api/profiles") == 0 && strcmp(method, "GET") == 0) {
        // dump all 5 profile slots as JSON
        std::string j = "{\"profiles\":[";
        for (int i = 1; i <= PROFILE_SLOTS; ++i) {
            if (i > 1) j += ",";
            Profile p = profileLoad(i);
            char b[1024];
            snprintf(b, sizeof(b),
                "{\"slot\":%d,\"name\":\"%s\",\"apps\":\"%s\"}",
                p.slot,
                jsonEscape(p.name).c_str(),
                jsonEscape(p.apps).c_str());
            j += b;
        }
        j += "]}";
        httpRespond(s, "200 OK", "application/json", j);
    } else if (strcmp(method, "POST") == 0 && strncmp(path, "/api/profile/", 13) == 0) {
        // /api/profile/<slot>/<rename|snapshot|apply>?body for rename
        int slot = 0; char op[24] = {};
        sscanf_s(path + 13, "%d/%23s", &slot, op, (unsigned)sizeof(op));
        if (slot < 1 || slot > PROFILE_SLOTS) {
            httpRespond(s, "400 Bad Request", "text/plain", std::string("bad slot"));
        } else {
            if (strcmp(op, "snapshot") == 0)      profileSnapshot(slot);
            else if (strcmp(op, "apply") == 0)    profileApply(slot);
            else if (strcmp(op, "rename") == 0) {
                // Body: name=...   (form-encoded)
                const char* body = strstr(buf, "\r\n\r\n");
                if (body) {
                    body += 4;
                    const char* p = strstr(body, "name=");
                    if (p) {
                        Profile pf = profileLoad(slot);
                        pf.name = p + 5;
                        // simple URL-decode: replace + with space, %xx with byte
                        std::string clean;
                        for (size_t i = 0; i < pf.name.size(); ++i) {
                            char c = pf.name[i];
                            if (c == '+') clean += ' ';
                            else if (c == '%' && i + 2 < pf.name.size()) {
                                char hex[3] = { pf.name[i+1], pf.name[i+2], 0 };
                                clean += (char)strtol(hex, nullptr, 16);
                                i += 2;
                            } else clean += c;
                        }
                        pf.name = clean.substr(0, clean.find('&'));
                        profileSave(pf);
                    }
                }
            }
            httpRespond(s, "200 OK", "application/json", "{\"ok\":true}");
        }
    } else if (strcmp(path, "/api/autostart") == 0 && strcmp(method, "GET") == 0) {
        std::string j = std::string("{\"enabled\":") + (autostartIsEnabled() ? "true" : "false") + "}";
        httpRespond(s, "200 OK", "application/json", j);
    } else if (strcmp(path, "/api/autostart") == 0 && strcmp(method, "POST") == 0) {
        const char* body = strstr(buf, "\r\n\r\n");
        bool on = body && strstr(body, "enabled=1");
        autostartSet(on);
        httpRespond(s, "200 OK", "application/json", "{\"ok\":true}");
    } else if (strcmp(path, "/api/tuning") == 0 && strcmp(method, "GET") == 0) {
        char j[256];
        snprintf(j, sizeof(j),
            "{\"refresh_interval_ms\":%d,\"debounce_ms\":%d,\"star_anim_ms\":%d}",
            g_refreshIntervalMs.load(), g_debounceMs.load(), g_starAnimMs.load());
        httpRespond(s, "200 OK", "application/json", j);
    } else if (strcmp(path, "/api/engine") == 0 && strcmp(method, "GET") == 0) {
        char j[80]; snprintf(j, sizeof(j), "{\"mode\":%d}", g_engine.load());
        httpRespond(s, "200 OK", "application/json", j);
    } else if (strcmp(path, "/api/engine") == 0 && strcmp(method, "POST") == 0) {
        const char* body = strstr(buf, "\r\n\r\n");
        if (body) {
            const char* p = strstr(body, "mode=");
            if (p) {
                int v = atoi(p + 5);
                g_engine.store(clampInt(v, 0, 2));
                engineSave();
                PostMessageW(g_msgWnd, WM_APP + 6, 0, 0);   // apply on main thread
            }
        }
        httpRespond(s, "200 OK", "application/json", "{\"ok\":true}");
    } else if (strcmp(path, "/api/tuning") == 0 && strcmp(method, "POST") == 0) {
        const char* body = strstr(buf, "\r\n\r\n");
        if (body) {
            body += 4;
            auto getInt = [&](const char* key, int& out) {
                const char* p = strstr(body, key);
                if (!p) return;
                p += strlen(key);
                if (*p != '=') return;
                out = atoi(p + 1);
            };
            int r = g_refreshIntervalMs.load();
            int d = g_debounceMs.load();
            int s = g_starAnimMs.load();
            getInt("refresh_interval_ms", r);
            getInt("debounce_ms",          d);
            getInt("star_anim_ms",         s);
            g_refreshIntervalMs.store(clampInt(r,  500, 60000));
            g_debounceMs.store       (clampInt(d,   20,  2000));
            g_starAnimMs.store       (clampInt(s,    0,  5000));
            tuningSave();
        }
        httpRespond(s, "200 OK", "application/json", "{\"ok\":true}");
    } else {
        httpRespond(s, "404 Not Found", "text/plain", std::string("not found"));
    }
    closesocket(s);
}

static void httpServerThread() {
    WSADATA wsa; WSAStartup(MAKEWORD(2, 2), &wsa);
    SOCKET ls = socket(AF_INET, SOCK_STREAM, 0);
    if (ls == INVALID_SOCKET) { logf("http: socket failed"); return; }
    // Deliberately NOT setting SO_REUSEADDR — on Windows it allows the same port
    // to be bound by multiple processes, which silently breaks the singleton model.
    sockaddr_in addr{};
    addr.sin_family = AF_INET;
    addr.sin_port = htons(HTTP_PORT);
    inet_pton(AF_INET, "127.0.0.1", &addr.sin_addr);
    if (bind(ls, (sockaddr*)&addr, sizeof(addr)) < 0) {
        logf("http: bind failed (port %d in use?)", HTTP_PORT);
        closesocket(ls); return;
    }
    if (listen(ls, 16) < 0) { logf("http: listen failed"); closesocket(ls); return; }
    logf("http: listening on 127.0.0.1:%d", HTTP_PORT);

    while (!g_quit.load()) {
        sockaddr_in ca{}; int cal = sizeof(ca);
        SOCKET cs = accept(ls, (sockaddr*)&ca, &cal);
        if (cs == INVALID_SOCKET) continue;
        // each request is short — handle inline; spawn thread if it ever isn't
        std::thread(httpHandle, cs).detach();
    }
    closesocket(ls);
    WSACleanup();
}

// ---------- WinEvent hook: event-driven refresh ----------
// Windows fires these for every window create/destroy/show/hide/name-change.
// Hooking them lets us react in real time instead of polling, and we
// debounce so a burst of events (e.g. an app launching) coalesces into one
// refresh.

static std::atomic<long long> g_lastEventMs{0};

static void CALLBACK winEventProc(HWINEVENTHOOK, DWORD event, HWND hwnd,
                                  LONG idObject, LONG idChild,
                                  DWORD, DWORD) {
    if (idObject != OBJID_WINDOW || idChild != CHILDID_SELF) return;
    if (!hwnd || GetWindow(hwnd, GW_OWNER) != nullptr) return;
    // Only react to *meaningful* changes. NAMECHANGE (e.g. browser tab
    // retitling) doesn't change anything we display and would otherwise
    // trigger constant refresh churn.
    if (event != EVENT_OBJECT_CREATE  && event != EVENT_OBJECT_DESTROY &&
        event != EVENT_OBJECT_SHOW    && event != EVENT_OBJECT_HIDE) return;
    g_lastEventMs.store(nowMs());
    PostMessageW(g_msgWnd, WM_APP + 5, 0, 0);
}

static void winEventInstall() {
    if (g_winEventHook) return;
    g_winEventHook = SetWinEventHook(
        EVENT_OBJECT_CREATE, EVENT_OBJECT_HIDE,
        nullptr, winEventProc, 0, 0,
        WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS);
    logf("WinEvent hook installed: %p (CREATE..HIDE only)", g_winEventHook);
}
static void winEventUninstall() {
    if (g_winEventHook) UnhookWinEvent(g_winEventHook);
    g_winEventHook = nullptr;
}

// ---------- ShellHook: explorer.exe pushes HSHELL_* messages to us ----------
static bool                    g_shellHookRegistered = false;
static void shellHookInstall() {
    if (g_shellHookRegistered || !g_msgWnd) return;
    g_shellHookMsg = RegisterWindowMessageW(L"SHELLHOOK");
    ChangeWindowMessageFilterEx(g_msgWnd, g_shellHookMsg, MSGFLT_ALLOW, nullptr);
    if (RegisterShellHookWindow(g_msgWnd)) {
        g_shellHookRegistered = true;
        logf("ShellHook registered (msg=%u)", g_shellHookMsg);
    } else {
        logf("RegisterShellHookWindow FAILED err=%lu", GetLastError());
        g_shellHookMsg = 0;   // don't keep a stale atom that wndProc would match against
    }
}
static void shellHookUninstall() {
    if (g_shellHookRegistered) {
        DeregisterShellHookWindow(g_msgWnd);
        g_shellHookRegistered = false;
    }
}

// Apply the engine selection — install/uninstall hooks + control polling.
static void engineApply() {
    Engine e = (Engine)g_engine.load();
    if (e == Engine::POLL) {
        winEventUninstall();
        shellHookUninstall();
    } else {
        winEventInstall();
        shellHookInstall();
    }
    logf("engine = %s", e == Engine::PUSH ? "PUSH" : e == Engine::HYBRID ? "HYBRID" : "POLL");
}

// ---------- re-register every tray icon (used after TaskbarCreated) ----------
static void reRegisterAllTrayIcons() {
    logf("re-registering all tray icons (master + %zu pinned + %zu windows)",
         g_pinned.size(), g_windows.size());
    masterTrayRegister();
    {
        std::lock_guard<std::mutex> lk(g_pinnedMx);
        for (auto& p : g_pinned) pinnedTrayAdd(p);
    }
    {
        std::lock_guard<std::mutex> lk(g_windowsMx);
        for (auto& kv : g_windows) trayAdd(kv.second);
    }
}

// ---------- message-only window proc ----------
static LRESULT CALLBACK wndProc(HWND h, UINT m, WPARAM w, LPARAM l) {
    if (g_taskbarCreatedMsg && m == g_taskbarCreatedMsg) {
        reRegisterAllTrayIcons();
        return 0;
    }
    if (g_shellHookMsg && m == g_shellHookMsg) {
        // wParam = HSHELL_* code, lParam = HWND that triggered it.
        int code = (int)(w & 0x7FFF);
        if (code == HSHELL_WINDOWCREATED   || code == HSHELL_WINDOWDESTROYED ||
            code == HSHELL_REDRAW          || code == HSHELL_FLASH           ||
            code == HSHELL_RUDEAPPACTIVATED) {
            SetTimer(g_msgWnd, TIMER_DEBOUNCE_REFRESH, g_debounceMs.load(), nullptr);
        }
        return 0;
    }
    switch (m) {
    // (WM_TIMER no longer used — refresh runs on a worker thread.)
    case TRAY_CALLBACK_MSG: {
        // NOTIFYICON_VERSION_4 layout:
        //   wParam = MAKEWPARAM(x, y)        -- cursor pos
        //   lParam = MAKELPARAM(message, uID)
        UINT evt = LOWORD(l);
        UINT uid = HIWORD(l);

        // ----- master star icon -----
        if (uid == MASTER_TRAY_UID) {
            if (evt == WM_LBUTTONUP)            actionOpenSettings();
            else if (evt == WM_LBUTTONDBLCLK)   actionOpenSettings();
            else if (evt == WM_RBUTTONUP || evt == WM_CONTEXTMENU)
                showContextMenuFor(MASTER_TRAY_UID);
            return 0;
        }

        // ----- pinned launcher icons (range PINNED_UID_BASE..) -----
        if (uid >= PINNED_UID_BASE && uid < MASTER_TRAY_UID) {
            const PinnedApp* p = pinnedFromUid(uid);
            if (!p) return 0;
            if (evt == WM_LBUTTONUP || evt == WM_LBUTTONDBLCLK) {
                pinnedLaunch(*p);
            } else if (evt == WM_RBUTTONUP || evt == WM_CONTEXTMENU) {
                // simple menu: launch, open file location
                POINT pt; GetCursorPos(&pt);
                SetForegroundWindow(g_msgWnd);
                HMENU m = CreatePopupMenu();
                AppendMenuW(m, MF_STRING, 1, L"Launch");
                AppendMenuW(m, MF_STRING, 2, L"Open file location");
                int cmd = TrackPopupMenu(m, TPM_RETURNCMD | TPM_RIGHTBUTTON, pt.x, pt.y, 0, g_msgWnd, nullptr);
                DestroyMenu(m);
                PostMessageW(g_msgWnd, WM_NULL, 0, 0);
                if (cmd == 1) pinnedLaunch(*p);
                else if (cmd == 2) {
                    std::wstring args = L"/select,\"" + p->lnkPath + L"\"";
                    ShellExecuteW(nullptr, L"open", L"explorer.exe", args.c_str(), nullptr, SW_SHOWNORMAL);
                }
            }
            return 0;
        }

        // ----- per-window icons -----
        HWND t = hwndFromUid(uid);
        if (!t) return 0;

        if (evt == WM_LBUTTONUP) {
            // Windows-taskbar-style toggle: clicking the foreground window's
            // tray icon minimises it; clicking any other tray icon focuses
            // (restoring if minimised). Eliminates the dblclk race where
            // the trailing WM_LBUTTONUP after WM_LBUTTONDBLCLK un-minimised.
            HWND fg = GetForegroundWindow();
            if (t == fg && !IsIconic(t)) {
                ShowWindow(t, SW_MINIMIZE);
            } else {
                if (IsIconic(t)) ShowWindow(t, SW_RESTORE);
                focusHwnd(t);
            }
        } else if (evt == WM_LBUTTONDBLCLK) {
            // Double click — always force foreground, never minimise.
            if (IsIconic(t)) ShowWindow(t, SW_RESTORE);
            focusHwnd(t);
        } else if (evt == WM_MBUTTONUP) {
            PostMessageW(t, WM_CLOSE, 0, 0);
        } else if (evt == WM_RBUTTONUP || evt == WM_CONTEXTMENU) {
            showContextMenuFor(uid);
        }
        return 0;
    }
    case WM_APP + 2:                  // refresh requested via HTTP
        refreshWindows();
        return 0;
    case WM_TBL_SETVISIBLE:           // (HWND, bool) - only main thread touches g_taskbarList
        taskbarListSetVisibleDirect((HWND)w, l != 0);
        return 0;
    case WM_APP + 4:                  // animation tick — advance master star frame
        masterTrayAnimate();
        return 0;
    case WM_APP + 5:                  // a WinEvent fired — arm/re-arm debounce
        SetTimer(h, TIMER_DEBOUNCE_REFRESH, g_debounceMs.load(), nullptr);
        return 0;
    case WM_APP + 6:                  // engine selection changed — re-apply
        engineApply();
        return 0;
    case WM_TIMER:
        if (w == TIMER_DEBOUNCE_REFRESH) {
            KillTimer(h, TIMER_DEBOUNCE_REFRESH);
            refreshWindows();          // event-driven coalesced refresh
        }
        return 0;
    case WM_DESTROY:
        g_quit.store(true);
        winEventUninstall();
        shellHookUninstall();
        masterTrayUnregister();
        unregisterAllPinned();
        // remove all our tray icons + restore any taskbar slots we hid
        {
            std::lock_guard<std::mutex> lk(g_windowsMx);
            for (auto& kv : g_windows) {
                if (kv.second.hiddenFromTaskbar)
                    taskbarListSetVisibleDirect(kv.second.hwnd, true);
                trayDelete(kv.second);
                if (kv.second.rawIcon)    DestroyIcon(kv.second.rawIcon);
                if (kv.second.badgedIcon) DestroyIcon(kv.second.badgedIcon);
            }
            g_windows.clear();
            g_uidToHwnd.clear();
        }
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(h, m, w, l);
}

// ---------- entry ----------
int APIENTRY wWinMain(HINSTANCE hi, HINSTANCE, LPWSTR, int) {
    SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    g_startMs = nowMs();
    resolveDataPaths();
    benchHeader();
    logf("--- TraySys %ls start (pid=%lu) ---", TRAYSYS_VERSION, GetCurrentProcessId());
    taskbarListInit();

    // Single-instance guard so we never double-register tray icons
    HANDLE mutex = CreateMutexW(nullptr, FALSE, L"Global\\TraySys.singleton.v1");
    if (!mutex) {
        logf("CreateMutex failed err=%lu", GetLastError());
    } else if (GetLastError() == ERROR_ALREADY_EXISTS) {
        logf("already running (pid=%lu), exiting", GetCurrentProcessId());
        CloseHandle(mutex);
        return 0;
    }
    logf("singleton mutex acquired");

    WNDCLASSW wc{};
    wc.lpfnWndProc   = wndProc;
    wc.hInstance     = hi;
    wc.lpszClassName = L"TraySysMsg";
    if (!RegisterClassW(&wc)) {
        logf("RegisterClass failed err=%lu", GetLastError());
    }

    g_msgWnd = CreateWindowExW(0, L"TraySysMsg", L"TraySys",
                               0, 0, 0, 0, 0,
                               HWND_MESSAGE, nullptr, hi, nullptr);
    if (!g_msgWnd) { logf("CreateWindow failed err=%lu", GetLastError()); return 1; }
    logf("msgwnd=%p created", g_msgWnd);

    // listen for explorer.exe restarts so we can re-register icons
    g_taskbarCreatedMsg = RegisterWindowMessageW(L"TaskbarCreated");
    ChangeWindowMessageFilterEx(g_msgWnd, g_taskbarCreatedMsg, MSGFLT_ALLOW, nullptr);
    logf("TaskbarCreated msg registered (id=%u)", g_taskbarCreatedMsg);

    tuningLoad();
    logf("tuning: refresh=%dms debounce=%dms star=%dms",
         g_refreshIntervalMs.load(), g_debounceMs.load(), g_starAnimMs.load());

    engineApply();   // installs WinEventHook + ShellHook unless engine is POLL

    // Load animated star frames + register the master tray icon BEFORE the
    // first refresh so the right-click menu is available immediately.
    loadStarFrames();
    masterTrayRegister();
    logf("master tray registered (uid=%u)", (unsigned)MASTER_TRAY_UID);

    loadPinnedApps();

    logf("calling initial refreshWindows...");
    refreshWindows();
    logf("initial refreshWindows done, icons=%zu", g_windows.size());

    // launch notification — balloon attached to the first registered icon
    {
        std::lock_guard<std::mutex> lk(g_windowsMx);
        if (!g_windows.empty()) {
            UINT firstUid = g_windows.begin()->second.uid;
            wchar_t body[200];
            swprintf(body, 200,
                L"Tracking %zu windows. Click an icon to focus, double-click to minimise/restore. Settings: http://127.0.0.1:%d/",
                g_windows.size(), HTTP_PORT);
            trayBalloon(firstUid, L"Laurence Trayhost " TRAYSYS_VERSION L" running", body);
        }
    }

    // Safety-net polling. PURE-PUSH engine skips this entirely; HYBRID runs
    // it as a fallback; POLL runs it as the only refresh source.
    std::thread refreshThread([]{
        while (!g_quit.load()) {
            int total = g_refreshIntervalMs.load();
            for (int i = 0; i < total / 50 && !g_quit.load(); ++i)
                std::this_thread::sleep_for(std::chrono::milliseconds(50));
            if (g_quit.load()) break;
            if (g_engine.load() == (int)Engine::PUSH) continue;  // never poll in pure push
            refreshWindows();
        }
    });
    logf("refresh thread spawned");

    // Animation thread — cycles the master star icon through colour frames.
    // g_starAnimMs == 0 pauses animation entirely (still polled at 250ms so
    // the slider can re-enable it instantly).
    std::thread starAnimThread([]{
        while (!g_quit.load()) {
            int sleep = g_starAnimMs.load();
            if (sleep <= 0) {
                std::this_thread::sleep_for(std::chrono::milliseconds(250));
                continue;
            }
            std::this_thread::sleep_for(std::chrono::milliseconds(sleep));
            if (g_quit.load()) break;
            PostMessageW(g_msgWnd, WM_APP + 4, 0, 0);
        }
    });

    // Heartbeat — proves the main thread is pumping messages.
    std::thread heartbeatThread([]{
        int n = 0;
        while (!g_quit.load()) {
            std::this_thread::sleep_for(std::chrono::milliseconds(2000));
            if (g_quit.load()) break;
            logf("heartbeat #%d  icons=%zu  rss_kb=%zu", ++n, g_windows.size(), rssKB());
        }
    });

    std::thread httpThread(httpServerThread);
    logf("http thread spawned");

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    g_quit.store(true);
    // wake the accept() so the http thread exits — open a connection to ourselves
    {
        SOCKET s = socket(AF_INET, SOCK_STREAM, 0);
        sockaddr_in a{}; a.sin_family = AF_INET; a.sin_port = htons(HTTP_PORT);
        inet_pton(AF_INET, "127.0.0.1", &a.sin_addr);
        connect(s, (sockaddr*)&a, sizeof(a));
        closesocket(s);
    }
    if (httpThread.joinable())      httpThread.join();
    if (refreshThread.joinable())   refreshThread.join();
    if (heartbeatThread.joinable()) heartbeatThread.join();
    if (starAnimThread.joinable())  starAnimThread.join();

    if (mutex) { ReleaseMutex(mutex); CloseHandle(mutex); }
    logf("--- TraySys end ---");
    return (int)msg.wParam;
}

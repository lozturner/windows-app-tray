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

// ---------- config ----------
#define TRAYSYS_VERSION       L"0.3.0"
#define TRAYSYS_VERSION_A      "0.3.0"
#define TRAY_CALLBACK_MSG     (WM_APP + 1)
#define HTTP_PORT             8731
#define REFRESH_INTERVAL_MS   1000
#define MASTER_TRAY_UID       0x7000             // dedicated UID for the animated master icon
#define STAR_FRAME_COUNT      8
#define STAR_FRAME_BASE_RES   100                // resource IDs 100..107
#define STAR_ANIM_INTERVAL_MS 140                // ~7 fps colour cycle

// ---------- log + bench ----------
// Paths resolved at startup to be next to the .exe so we never lose them to a weird CWD.
static char g_logPath[MAX_PATH]   = "traysys.log";
static char g_benchPath[MAX_PATH] = "traysys.bench.csv";

static void resolveDataPaths() {
    char exe[MAX_PATH] = {};
    GetModuleFileNameA(nullptr, exe, MAX_PATH);
    char* slash = strrchr(exe, '\\');
    if (slash) {
        *slash = 0;
        snprintf(g_logPath,   MAX_PATH, "%s\\LaurenceTrayhost_v%s.log",       exe, TRAYSYS_VERSION_A);
        snprintf(g_benchPath, MAX_PATH, "%s\\LaurenceTrayhost_v%s.bench.csv", exe, TRAYSYS_VERSION_A);
    }
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
                    // hide browser window from real taskbar if configured
                    if (g_hideBrowsersFromTaskbar.load() && isBrowserExe(ins->second.exeNameLower)) {
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
                std::wstring t = titleOf(h);
                if (t != it->second.title) {
                    it->second.title = std::move(t);
                    trayModify(it->second);
                }
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

static void showContextMenuFor(UINT uid) {
    POINT pt; GetCursorPos(&pt);
    SetForegroundWindow(g_msgWnd);

    HMENU m = CreatePopupMenu();
    if (uid == MASTER_TRAY_UID) {
        // ----- master star icon menu -----
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

        switch (cmd) {
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
  <button onclick="q()">Quit TraySys</button>
</div>
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
  const tb = document.querySelector('#t tbody');
  tb.innerHTML = r.windows.map(w=>`<tr><td class="uid">${w.uid}</td><td class="exe">${w.exe}</td><td class="title">${w.title}</td></tr>`).join('');
}
async function r(){ await fetch('/api/refresh',{method:'POST'}); load(); }
async function q(){ if(confirm('Quit TraySys and remove all tray icons?')) await fetch('/api/quit',{method:'POST'}); }
load(); setInterval(load, 1500);
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
        // post a custom message; the main thread does the work
        PostMessageW(g_msgWnd, WM_APP + 2, 0, 0);
        httpRespond(s, "200 OK", "application/json", "{\"ok\":true}");
    } else if (strcmp(path, "/api/quit") == 0 && strcmp(method, "POST") == 0) {
        httpRespond(s, "200 OK", "application/json", "{\"ok\":true}");
        PostMessageW(g_msgWnd, WM_CLOSE, 0, 0);
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

// ---------- message-only window proc ----------
static LRESULT CALLBACK wndProc(HWND h, UINT m, WPARAM w, LPARAM l) {
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

        // ----- per-window icons -----
        HWND t = hwndFromUid(uid);
        if (!t) return 0;

        if (evt == WM_LBUTTONUP) {
            // single click — focus, restoring if minimised
            focusHwnd(t);
        } else if (evt == WM_LBUTTONDBLCLK) {
            // double click — toggle visibility (minimise <-> restore+focus)
            if (IsIconic(t)) {
                ShowWindow(t, SW_RESTORE);
                focusHwnd(t);
            } else {
                ShowWindow(t, SW_MINIMIZE);
            }
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
    case WM_DESTROY:
        g_quit.store(true);
        masterTrayUnregister();
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

    // Load animated star frames + register the master tray icon BEFORE the
    // first refresh so the right-click menu is available immediately.
    loadStarFrames();
    masterTrayRegister();
    logf("master tray registered (uid=%u)", (unsigned)MASTER_TRAY_UID);

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

    // Refresh on a dedicated worker thread, NOT WM_TIMER —
    // Shell_NotifyIcon can block, and a wedged message thread kills tray callbacks.
    std::thread refreshThread([]{
        while (!g_quit.load()) {
            for (int i = 0; i < REFRESH_INTERVAL_MS / 50 && !g_quit.load(); ++i)
                std::this_thread::sleep_for(std::chrono::milliseconds(50));
            if (g_quit.load()) break;
            refreshWindows();
        }
    });
    logf("refresh thread spawned");

    // Animation thread — cycles the master star icon through colour frames.
    std::thread starAnimThread([]{
        while (!g_quit.load()) {
            std::this_thread::sleep_for(std::chrono::milliseconds(STAR_ANIM_INTERVAL_MS));
            if (g_quit.load()) break;
            // Drive the animation from the main thread to avoid Shell_NotifyIcon
            // races; PostMessage is cheap.
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

# Laurence Trayhost

Every active window — and (soon) every browser tab — represented as its own icon in the real Windows system tray. Click to focus, double-click to toggle minimise, middle-click to close, right-click for menu. Live state and settings on a localhost HTML page.

Single native Win32 binary, MSVC-built, no runtime, ~245 KB.

## Build

Requires **Visual Studio 2022 Build Tools** (cl.exe). No other dependencies.

```cmd
build.bat
```

The script:
1. Reads the version from `src/main.cpp`
2. Compiles with MSVC to `C:\LaurenceTrayhost\LaurenceTrayhost_v<version>.exe`
3. Kills any previously-running instance
4. Launches the new version

## Run

```cmd
C:\LaurenceTrayhost\LaurenceTrayhost_v0.1.1.exe
```

Settings UI: <http://127.0.0.1:8731/>

## Interaction grammar

| Action | Effect |
|---|---|
| **Left click** | Focus window (restoring if minimised) |
| **Double click** | Toggle minimise / restore — flick on/off views |
| **Middle click** | Close window |
| **Right click** | Context menu (focus, close, settings, quit) |

## Files (versioned)

- `LaurenceTrayhost_v<ver>.exe` — the binary
- `LaurenceTrayhost_v<ver>.log` — runtime log
- `LaurenceTrayhost_v<ver>.bench.csv` — per-refresh benchmark sample (ts, refresh_ms, icons, rss_kb, added, removed)

## Architecture

Single C++17 file. Three threads:

1. **Main / message thread** — pumps Win32 messages, handles tray callbacks (click → focus, dblclick → toggle minimise, etc.)
2. **Refresh thread** — every ~1s: `EnumWindows` → diff against tracked map → `Shell_NotifyIcon NIM_ADD/MODIFY/DELETE`. Decoupled from the message loop because `Shell_NotifyIcon` can block on a busy shell.
3. **HTTP thread** — winsock listener on `127.0.0.1:8731`, serves `/`, `/api/state`, `/api/refresh`, `/api/quit`.

Singleton enforced via named mutex.

## Roadmap

- v0.2: browser tabs as tray icons (Chrome/Edge/Comet via UI Automation)
- v0.2: settings UI persists ordering / hidden-window rules
- v0.3: per-monitor "send-to" actions, app-grouping toggle

## Benchmarks

Baseline (v0.1.1, 14 windows):
- Cold start → first icon visible: ~250 ms
- Steady-state refresh: 1–5 ms
- RSS: 8–9 MB
- CPU idle: 0%

## Legacy Python prototype

The previous `pystray`-based prototype (fixed-icon launcher) lives in [`legacy_python/`](legacy_python/) — preserved for the icon-extraction work. The C++ Trayhost replaces it: instead of a fixed set of launcher icons, every active window gets its own.

## Part of

[Lawrence: Move In](https://github.com/lozturner/lawrence-move-in) — Loz's Windows productivity suite.

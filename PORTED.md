# PORTED.md — tools-to-port audit

Track for each existing tray/taskbar tool: **wrap** (call its DLL), **port** (rewrite cleanly), **refract** (steal the idea, redesign), **skip** (not worth it).

## Decided

| Tool | Logo / vibe | What it does | Disposition |
|---|---|---|---|
| **ExplorerPatcher** | feather (the "eagle") | Hooks explorer.exe, remodels Win11 taskbar back to Win10 style | **Refract** — we read its source for AppBar registration patterns and Shell_TrayWnd hide tricks, but stay out-of-process (no DLL injection). |
| **RBTray** | tiny red box | Minimise any window to tray | **Port** — re-implement minimise-to-tray as a built-in TrayDeck verb (Shift+min btn). |
| **Traymond** | T mark | Send-to-tray via hotkey | **Refract** — same idea, lives as our `key→close` action. |
| **7+ Taskbar Tweaker** | wrench | Reorder/group/middle-click-close on real taskbar | **Refract** — interactions inspired (middle-click closes, scroll cycles), but on our bar, not the vanilla one. |
| **TrayStatus** | square | Caps/Num/Scroll lock indicators in tray | **Port** — small, self-contained, becomes a TrayDeck "system glyph" tile. |
| **AutoHideMouseCursor** | cursor | Hide cursor when idle | **Skip** for v0.1, revisit. |
| **Bins** (Mac) | folders | Group dock icons into stacks | **Refract** — our app-grouping with child-window grid is essentially this idea on Windows. |
| **Stardock Fences** | rectangle | Desktop region grouping | **Skip** — different concept, not tray. |
| **NotifyIconOverflowWindow hack** | n/a | Force all tray icons to "always show" | **Port** — TrayDeck enforces "show all" on first run via TB_GETBUTTON walk. |
| **TaskbarX** | bar | Centre/animate Win taskbar icons | **Skip** — works on the thing we're hiding. |

## Pending audit

- StartIsBack / StartAllBack — start menu plugins, may inform our start-button replacement.
- ObjectDock / RocketDock — Mac-style dock; old but icon-grouping logic worth a read.
- ueli / Wox / PowerToys Run — keystroke launchers; could fold into right-click "new instance" with search.
- AltDrag (already installed on this box) — alt+drag-to-move; complementary, leave alone.

## Notes on the tray-icon harvest

The technique used by RBTray / 7+TT to enumerate the real notification area:

```
FindWindow("Shell_TrayWnd")
  -> FindWindowEx (..., "TrayNotifyWnd")
    -> FindWindowEx (..., "SysPager")
      -> FindWindowEx (..., "ToolbarWindow32")  // the icon list
SendMessage TB_BUTTONCOUNT, then per index TB_GETBUTTON
ReadProcessMemory to pull TBBUTTON across process boundary (explorer.exe owns it)
```

This is cross-process and undocumented but stable since XP. We'll wrap it in `tray_harvest.cpp` once v0.1 lands.

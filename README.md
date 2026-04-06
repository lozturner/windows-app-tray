# Windows App Tray

Fixed system tray icons for your frequently used apps. Real icons extracted from the actual Windows executables — not generated placeholders.

## What it does

Adds permanent tray icons for your apps. Right-click any icon to open the app. They don't shuffle, they don't disappear.

| Icon | App | Source |
|------|-----|--------|
| ![Explorer](icons/explorer.png) | File Explorer | Extracted from `explorer.exe` |
| ![Edge](icons/edge.png) | Microsoft Edge | Extracted from `msedge.exe` |
| ![Chrome](icons/chrome.png) | Google Chrome | Extracted from `chrome.exe` |
| ![Perplexity](icons/perplexity.png) | Perplexity | Official favicon |
| ![Comet](icons/comet.png) | Comet Browser | Extracted from `comet.exe` |

## Quick Start

```bash
pip install pystray pillow psutil pywin32
python app_tray.py
```

## Configuration

Edit `app_tray_config.json` to add/remove apps. Each entry needs:

```json
{
  "name": "My App",
  "icon": "myapp.ico",
  "path": "C:\\path\\to\\app.exe",
  "args": ""
}
```

Drop the `.ico` file in the `icons/` folder.

## How the icons were sourced

See [`how-icons-were-sourced.html`](how-icons-were-sourced.html) for the full documented process with code samples.

## Requirements

- Python 3.11+
- `pystray`, `pillow`, `psutil`, `pywin32`

## Part of

[Lawrence: Move In](https://github.com/lozturner/lawrence-move-in) — 17 Python applets that fix what Windows gets wrong.

"""
Windows App Tray — Standalone v1.0.0
Fixed system tray icons for your frequently used apps.
Uses REAL application icons extracted from the actual executables.

Apps: File Explorer, Microsoft Edge, Google Chrome, Perplexity, Comet Browser
"""
__version__ = "1.0.0"

import json, os, subprocess, sys, threading
from pathlib import Path

import pystray
from PIL import Image

SCRIPT_DIR = Path(__file__).resolve().parent
ICONS_DIR  = SCRIPT_DIR / "icons"
CONFIG_PATH = SCRIPT_DIR / "app_tray_config.json"

DEFAULT_APPS = [
    {
        "name": "File Explorer",
        "icon": "explorer.ico",
        "path": "explorer.exe",
        "args": "",
    },
    {
        "name": "Microsoft Edge",
        "icon": "edge.ico",
        "path": r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "args": "",
    },
    {
        "name": "Google Chrome",
        "icon": "chrome.ico",
        "path": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        "args": "",
    },
    {
        "name": "Perplexity",
        "icon": "perplexity.ico",
        "path": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        "args": "--app=https://perplexity.ai",
    },
    {
        "name": "Comet Browser",
        "icon": "comet.ico",
        "path": r"C:\Users\123\AppData\Local\Perplexity\Comet\Application\comet.exe",
        "args": "",
    },
]

def load_config():
    if CONFIG_PATH.exists():
        try: return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except: pass
    cfg = {"apps": DEFAULT_APPS}
    save_config(cfg)
    return cfg

def save_config(cfg):
    CONFIG_PATH.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")

def load_icon(icon_name):
    """Load a .ico file as a PIL Image for pystray."""
    icon_path = ICONS_DIR / icon_name
    if icon_path.exists():
        try:
            return Image.open(str(icon_path))
        except Exception:
            pass
    # Fallback: generate a coloured square
    from PIL import ImageDraw
    img = Image.new("RGBA", (64, 64), (100, 100, 100, 255))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([4, 4, 59, 59], radius=12, fill=(180, 180, 180, 255))
    return img

def launch_app(app):
    path = app["path"]
    args = app.get("args", "")
    try:
        if args:
            subprocess.Popen(f'"{path}" {args}', shell=True, creationflags=0x8)
        elif path.lower() == "explorer.exe":
            subprocess.Popen(["explorer.exe"], creationflags=0x8)
        else:
            subprocess.Popen([path], creationflags=0x8)
    except Exception as e:
        print(f"Failed to launch {app['name']}: {e}")

def make_launcher(app):
    def _fn(icon, item):
        launch_app(app)
    return _fn

def make_quitter():
    def _fn(icon, item):
        icon.stop()
    return _fn

def create_tray(app):
    icon_img = load_icon(app.get("icon", ""))

    menu = pystray.Menu(
        pystray.MenuItem(f"Open {app['name']}", make_launcher(app)),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Quit", make_quitter()),
    )

    return pystray.Icon(
        f"apptray_{app['name'].replace(' ','_')}",
        icon_img,
        app["name"],
        menu)

def main():
    cfg = load_config()
    apps = cfg["apps"]

    # Verify paths, try alternatives
    for app in apps:
        p = app["path"]
        if p.lower() == "explorer.exe":
            continue
        if not os.path.exists(p):
            for alt in [
                p.replace("Program Files", "Program Files (x86)"),
                p.replace("Program Files (x86)", "Program Files"),
            ]:
                if os.path.exists(alt):
                    app["path"] = alt
                    break

    threads = []
    icons = []
    for app in apps:
        icon = create_tray(app)
        icons.append(icon)
        t = threading.Thread(target=icon.run, daemon=True)
        t.start()
        threads.append(t)

    try:
        for t in threads:
            t.join()
    except KeyboardInterrupt:
        for icon in icons:
            icon.stop()

if __name__ == "__main__":
    # Kill old instances
    try:
        import psutil
        my_pid = os.getpid()
        for p in psutil.process_iter(["pid", "name", "cmdline"]):
            if p.info["pid"] == my_pid: continue
            try:
                cmd = " ".join(p.info.get("cmdline") or [])
                if "app_tray" in cmd and "python" in (p.info.get("name") or "").lower():
                    p.kill()
            except: pass
    except: pass

    main()

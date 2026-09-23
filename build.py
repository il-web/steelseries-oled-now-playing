"""
Build a standalone NowPlaying.exe (no Python needed to run it).

    pip install -r requirements.txt pyinstaller
    python build.py

The result is dist/NowPlaying.exe.
"""

import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
BUILD_DIR = os.path.join(ROOT, "build")
DIST_DIR = os.path.join(ROOT, "dist")

sys.path.insert(0, ROOT)
from now_playing import create_tray_icon  # noqa: E402


def main():
    os.makedirs(BUILD_DIR, exist_ok=True)

    # Reuse the tray icon artwork as the .exe icon
    icon_path = os.path.join(BUILD_DIR, "icon.ico")
    create_tray_icon(size=256).save(
        icon_path, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    )

    subprocess.run([
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",  # no console window
        "--name", "NowPlaying",
        "--icon", icon_path,
        # winrt and pycaw load parts of themselves dynamically
        "--collect-submodules", "winrt",
        "--collect-submodules", "pycaw",
        "--hidden-import", "pystray._win32",
        "--distpath", DIST_DIR,
        "--workpath", os.path.join(BUILD_DIR, "pyinstaller"),
        "--specpath", BUILD_DIR,
        os.path.join(ROOT, "now_playing.py"),
    ], check=True, cwd=ROOT)

    print(f"\nBuilt {os.path.join(DIST_DIR, 'NowPlaying.exe')}")


if __name__ == "__main__":
    main()

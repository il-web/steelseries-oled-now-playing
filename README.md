# SteelSeries OLED Now Playing

Display currently playing media info on your SteelSeries keyboard's OLED screen.

Shows the **song title**, **artist**, **album art** and a **progress bar with elapsed/remaining time** while music is playing. When paused or idle, it shows a **clock**, **system stats** or the **weather**. Volume changes are shown as a temporary overlay with a bar and percentage.

## Compatible Keyboards

Any SteelSeries keyboard with a 128x40 OLED screen:
- Apex 5
- Apex 7 / 7 TKL
- Apex Pro / Pro TKL
- Apex Pro 2023 / Pro TKL 2023

## Features

- Pulls "now playing" info from **any media source** via Windows (Spotify, YouTube, VLC, etc.)
- Scrolling marquee effect for long song titles and artist names
- Progress bar with elapsed / total or remaining time
- Optional album art (dithered for the monochrome screen)
- Japanese, Chinese, Korean, Hebrew, Arabic and emoji titles display correctly (fallback fonts + right-to-left support, including scrolling direction)
- Idle screen when nothing is playing: clock (12/24h, optional date), CPU/RAM/network stats, current weather, or cycle through all of them
- Volume overlay when you change system volume (shows bar + percentage + mute state)
- Everything configurable from the system tray menu
- Start with Windows, pause the display, single instance
- Survives SteelSeries GG restarts and can be started before GG (reconnects automatically)

## Installation

Requirements: Windows 10/11, Python 3.9+, [SteelSeries GG](https://steelseries.com/gg).

```bash
git clone https://github.com/il-web/steelseries-oled-now-playing.git
cd steelseries-oled-now-playing
pip install -r requirements.txt
```

**Double-click `Start Now Playing.bat`** to run with no console window. To launch it automatically at login, use **Start with Windows** in the tray menu.

Or run from a terminal:

```bash
python now_playing.py
```

### Updating

Quit the app from the tray menu, then:

```bash
git pull
pip install -r requirements.txt
```

Re-running `pip install` matters: new versions may need new packages.

## Usage

A music-note icon appears in the system tray. Right-click it for options:

| Menu | What it does |
|------|--------------|
| **Now playing** | Show album art; time display (elapsed / total, elapsed / remaining, hidden) |
| **When idle** | Clock, system stats, weather, or cycle through all; 24-hour clock; show date; set weather city; °C / °F |
| **Volume overlay** | Show the volume bar when you change the volume |
| **Pause display** | Give the OLED back to SteelSeries GG until you unpause |
| **Start with Windows** | Launch automatically when you log in |
| **Open settings file** | Edit advanced settings (see below) |
| **Open log** | View the log, useful for bug reports |

Only one copy runs at a time; starting it again just shows a reminder that it's already in the tray.

### Settings file

Settings are stored in `%APPDATA%\SteelSeriesNowPlaying\settings.json` and picked up automatically when you save the file. Besides the tray options, you can change:

| Setting | Default | Description |
|---------|---------|-------------|
| `scroll_speed` | `2` | Marquee speed in pixels per frame |
| `title_font_size` | `11` | Title text size in pixels |
| `artist_font_size` | `10` | Artist text size in pixels |
| `cycle_seconds` | `10` | How long each idle screen shows in "cycle" mode |

Numbers must be between 1 and 60. Invalid values are ignored (and noted in the log), and the default is used instead.

### Privacy

Nothing leaves your PC unless you set a weather city. Then the city name is sent to [Open-Meteo](https://open-meteo.com/) (free, no account needed) to look up the weather, every 15 minutes while the weather screen is in use.

## Troubleshooting

- **Nothing shows on the keyboard:** make sure SteelSeries GG is running. The app waits for GG and connects on its own, including after GG restarts.
- **"Already running" message:** the app is already in the system tray (check the hidden-icons arrow).
- **Something else:** open the log from the tray menu (**Open log**), or find it at `%APPDATA%\SteelSeriesNowPlaying\now_playing.log`, and include it in a [bug report](https://github.com/il-web/steelseries-oled-now-playing/issues/new/choose).

## Uninstalling

1. In the tray menu, turn off **Start with Windows**, then **Quit**.
2. Delete the project folder and `%APPDATA%\SteelSeriesNowPlaying`.

## Building a standalone .exe (optional)

```bash
pip install -r requirements.txt pyinstaller
python build.py
```

The result is `dist/NowPlaying.exe`, which runs without Python installed. The .exe is not code-signed, so Windows **Smart App Control** blocks it (SmartScreen may also warn about it); on those PCs use `Start Now Playing.bat` instead.

## How It Works

1. Reads currently playing media (and album art) from Windows System Media Transport Controls (SMTC)
2. Monitors system volume via Windows Core Audio API
3. Renders frames as 128x40 monochrome bitmaps using Pillow
4. Sends bitmaps to the OLED via the SteelSeries GameSense SDK

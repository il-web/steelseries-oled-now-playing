# SteelSeries OLED Now Playing

Display currently playing media info on your SteelSeries keyboard's OLED screen.

Shows the **song title**, **artist**, and a **progress bar** while music is playing. When paused or idle, it displays the **current time**. Volume changes are shown as a temporary overlay with a bar and percentage.

## Compatible Keyboards

Any SteelSeries keyboard with a 128x40 OLED screen:
- Apex 5
- Apex 7 / 7 TKL
- Apex Pro / Pro TKL
- Apex Pro 2023 / Pro TKL 2023

## Features

- Pulls "now playing" info from **any media source** via Windows (Spotify, YouTube, VLC, etc.)
- Scrolling marquee effect for long song titles and artist names
- Progress bar for track position
- Clock display when media is paused/stopped
- Volume overlay when you change system volume (shows bar + percentage + mute state)

## Requirements

- Windows 10/11
- Python 3.9+
- [SteelSeries GG](https://steelseries.com/gg) installed and running

## Installation

```bash
git clone https://github.com/il-web/steelseries-oled-now-playing.git
cd steelseries-oled-now-playing
pip install -r requirements.txt
```

## Usage

**Double-click `Start Now Playing.bat`** to run with no console window. A tray icon will appear — right-click it to quit.

Or run from terminal:

```bash
python now_playing.py
```

## How It Works

1. Reads currently playing media from Windows System Media Transport Controls (SMTC)
2. Monitors system volume via Windows Core Audio API
3. Renders frames as 128x40 monochrome bitmaps using Pillow
4. Sends bitmaps to the OLED via the SteelSeries GameSense SDK

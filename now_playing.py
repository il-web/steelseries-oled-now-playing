"""
SteelSeries Apex 5 OLED - Now Playing Display
Shows current media info (title, artist, progress bar) on the keyboard OLED screen.
"""

import asyncio
import json
import logging
import os
import sys
import threading
import time

import pystray
import requests
from PIL import Image, ImageDraw, ImageFont

# Log to file for debugging (especially when running .pyw with no console)
log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "now_playing.log")
logging.basicConfig(
    filename=log_path,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s: %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("now_playing")

# --- SteelSeries GameSense SDK ---

GAME_NAME = "NOWPLAYING"
EVENT_NAME = "SCREEN"
OLED_WIDTH = 128
OLED_HEIGHT = 40


def get_steelseries_address():
    """Read the SteelSeries Engine address from coreProps.json."""
    props_path = os.path.join(
        os.environ.get("PROGRAMDATA", "C:\\ProgramData"),
        "SteelSeries", "SteelSeries Engine 3", "coreProps.json"
    )
    try:
        with open(props_path, "r") as f:
            props = json.load(f)
        return f"http://{props['address']}"
    except FileNotFoundError:
        print("ERROR: SteelSeries GG/Engine 3 not found.")
        print(f"Expected file: {props_path}")
        sys.exit(1)
    except (json.JSONDecodeError, KeyError):
        print("ERROR: Could not parse SteelSeries coreProps.json")
        sys.exit(1)


def register_game(base_url):
    """Register our app with SteelSeries Engine."""
    resp = requests.post(f"{base_url}/game_metadata", json={
        "game": GAME_NAME,
        "game_display_name": "Now Playing",
        "developer": "Custom",
    })
    resp.raise_for_status()


def bind_screen_event(base_url):
    """Bind a screen handler for bitmap mode on the 128x40 OLED."""
    resp = requests.post(f"{base_url}/bind_game_event", json={
        "game": GAME_NAME,
        "event": EVENT_NAME,
        "value_optional": True,
        "handlers": [{
            "device-type": "screened-128x40",
            "zone": "one",
            "mode": "screen",
            "datas": [{
                "has-text": False,
                "image-data": [0] * 640,
            }]
        }]
    })
    resp.raise_for_status()


def image_to_bitmap(image):
    """Convert a PIL Image to a 640-byte array for the 128x40 OLED."""
    img = image.convert("1").resize((OLED_WIDTH, OLED_HEIGHT))
    pixels = img.load()
    bitmap = []
    for y in range(OLED_HEIGHT):
        for x_byte in range(OLED_WIDTH // 8):
            byte_val = 0
            for bit in range(8):
                x = x_byte * 8 + bit
                if pixels[x, y] > 0:
                    byte_val |= (1 << (7 - bit))
            bitmap.append(byte_val)
    return bitmap


def send_frame(base_url, bitmap_data):
    """Send a bitmap frame to the OLED."""
    requests.post(f"{base_url}/game_event", json={
        "game": GAME_NAME,
        "event": EVENT_NAME,
        "data": {
            "value": 0,
            "frame": {
                "image-data-128x40": bitmap_data,
            }
        }
    })


def send_heartbeat(base_url):
    """Prevent the 15-second timeout."""
    requests.post(f"{base_url}/game_heartbeat", json={"game": GAME_NAME})


def cleanup(base_url):
    """Remove our game from SteelSeries Engine."""
    try:
        requests.post(f"{base_url}/remove_game", json={"game": GAME_NAME})
    except Exception:
        pass


# --- Windows Volume ---

def get_system_volume():
    """Get the current system volume (0.0 - 1.0) and mute state."""
    from pycaw.pycaw import AudioUtilities

    speakers = AudioUtilities.GetSpeakers()
    vol = speakers.EndpointVolume
    level = vol.GetMasterVolumeLevelScalar()  # 0.0 to 1.0
    muted = vol.GetMute()
    return level, bool(muted)


# --- Windows Media Info ---

async def get_media_info():
    """Get currently playing media info from Windows SMTC."""
    from winrt.windows.media.control import (
        GlobalSystemMediaTransportControlsSessionManager as MediaManager,
    )

    manager = await MediaManager.request_async()
    session = manager.get_current_session()
    if session is None:
        return None

    try:
        media_props = await session.try_get_media_properties_async()
        title = media_props.title or ""
        artist = media_props.artist or ""
    except Exception:
        return None

    try:
        playback_info = session.get_playback_info()
        status = playback_info.playback_status
    except Exception:
        status = None

    try:
        timeline = session.get_timeline_properties()
        position_secs = timeline.position.total_seconds()
        duration_secs = timeline.end_time.total_seconds()
    except Exception:
        position_secs = 0
        duration_secs = 0

    return {
        "title": title,
        "artist": artist,
        "status": status,
        "position": position_secs,
        "duration": duration_secs,
    }


# --- Rendering ---

def get_font(size):
    """Try to load a clean font, fall back to default."""
    font_paths = [
        "C:/Windows/Fonts/segoeui.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/tahoma.ttf",
    ]
    for path in font_paths:
        try:
            return ImageFont.truetype(path, size)
        except (IOError, OSError):
            continue
    return ImageFont.load_default()


class ScrollingText:
    """Manages horizontal scrolling for a single line of text."""

    def __init__(self, scroll_speed=2, pause_ticks=15):
        self.text = ""
        self.scroll_offset = 0
        self.pause_counter = 0
        self.scroll_speed = scroll_speed  # pixels per tick
        self.pause_ticks = pause_ticks    # ticks to pause at start and end
        self.scrolling_forward = True     # True = left, False = right (back)
        self.pausing = True
        self.text_width = 0

    def update_text(self, new_text):
        """Reset scroll state when the text changes."""
        if new_text != self.text:
            self.text = new_text
            self.scroll_offset = 0
            self.pause_counter = 0
            self.scrolling_forward = True
            self.pausing = True
            self.text_width = 0

    def tick(self, max_width):
        """Advance the scroll animation by one frame (bouncing back and forth)."""
        if self.text_width <= max_width:
            self.scroll_offset = 0
            return

        max_scroll = self.text_width - max_width

        # Pause at each end before reversing direction
        if self.pausing:
            self.pause_counter += 1
            if self.pause_counter >= self.pause_ticks:
                self.pausing = False
                self.pause_counter = 0
            return

        if self.scrolling_forward:
            self.scroll_offset = min(self.scroll_offset + self.scroll_speed, max_scroll)
            if self.scroll_offset >= max_scroll:
                self.scrolling_forward = False
                self.pausing = True
        else:
            self.scroll_offset = max(self.scroll_offset - self.scroll_speed, 0)
            if self.scroll_offset <= 0:
                self.scrolling_forward = True
                self.pausing = True


# Global scroll state for title and artist
title_scroller = ScrollingText(scroll_speed=2, pause_ticks=15)
artist_scroller = ScrollingText(scroll_speed=2, pause_ticks=15)


def draw_scrolling_text(draw, scroller, text, font, x, y, max_width, line_height):
    """Draw text with horizontal scrolling if it exceeds max_width."""
    scroller.update_text(text)
    bbox = draw.textbbox((0, 0), text, font=font)
    scroller.text_width = bbox[2] - bbox[0]
    scroller.tick(max_width)

    # Create a temporary image with fixed line height for clipping
    text_img = Image.new("1", (max_width, line_height), 0)
    text_draw = ImageDraw.Draw(text_img)
    # Draw at -bbox[1] to account for font ascender offset
    text_draw.text((-scroller.scroll_offset, -bbox[1]), text, fill=1, font=font)

    # Paste the clipped text onto the main image
    draw._image.paste(text_img, (x, y))


def render_now_playing(info):
    """Render the now-playing screen as a PIL Image."""
    img = Image.new("1", (OLED_WIDTH, OLED_HEIGHT), 0)
    draw = ImageDraw.Draw(img)

    title_font = get_font(11)
    artist_font = get_font(9)

    title = info["title"] if info["title"] else "Unknown"
    artist = info["artist"] if info["artist"] else "Unknown"

    max_text_width = OLED_WIDTH - 2

    # Draw title (line 1) - scrolls if too long
    draw_scrolling_text(draw, title_scroller, title, title_font, 1, 0, max_text_width, 13)

    # Draw artist (line 2) - scrolls if too long
    draw_scrolling_text(draw, artist_scroller, artist, artist_font, 1, 14, max_text_width, 18)

    # Progress bar (pushed to bottom)
    bar_y = 34
    bar_height = 5
    bar_left = 1
    bar_right = OLED_WIDTH - 2

    # Bar outline
    draw.rectangle([bar_left, bar_y, bar_right, bar_y + bar_height], outline=1, fill=0)

    # Bar fill
    duration = info["duration"]
    position = info["position"]
    if duration > 0:
        progress = min(position / duration, 1.0)
        fill_width = int((bar_right - bar_left - 2) * progress)
        if fill_width > 0:
            draw.rectangle(
                [bar_left + 1, bar_y + 1, bar_left + 1 + fill_width, bar_y + bar_height - 1],
                fill=1,
            )

    return img


def render_volume(volume_pct, muted):
    """Render a volume overlay on the OLED."""
    img = Image.new("1", (OLED_WIDTH, OLED_HEIGHT), 0)
    draw = ImageDraw.Draw(img)

    label_font = get_font(11)
    pct_font = get_font(14)

    # Label
    if muted:
        label = "Muted"
    else:
        label = "Volume"
    bbox = draw.textbbox((0, 0), label, font=label_font)
    label_w = bbox[2] - bbox[0]
    draw.text(((OLED_WIDTH - label_w) // 2, -bbox[1]), label, fill=1, font=label_font)

    # Volume bar
    bar_y = 16
    bar_height = 8
    bar_left = 4
    bar_right = OLED_WIDTH - 5

    # Bar outline
    draw.rectangle([bar_left, bar_y, bar_right, bar_y + bar_height], outline=1, fill=0)

    # Bar fill
    if not muted:
        fill_width = int((bar_right - bar_left - 2) * volume_pct)
        if fill_width > 0:
            draw.rectangle(
                [bar_left + 1, bar_y + 1, bar_left + 1 + fill_width, bar_y + bar_height - 1],
                fill=1,
            )

    # Percentage text
    pct_str = f"{int(volume_pct * 100)}%"
    bbox = draw.textbbox((0, 0), pct_str, font=pct_font)
    pct_w = bbox[2] - bbox[0]
    draw.text(((OLED_WIDTH - pct_w) // 2, 28 - bbox[1]), pct_str, fill=1, font=pct_font)

    return img


def render_idle():
    """Render the current time when nothing is playing."""
    from datetime import datetime

    img = Image.new("1", (OLED_WIDTH, OLED_HEIGHT), 0)
    draw = ImageDraw.Draw(img)

    time_font = get_font(18)
    now = datetime.now().strftime("%H:%M")
    bbox = draw.textbbox((0, 0), now, font=time_font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    x = (OLED_WIDTH - text_w) // 2
    y = (OLED_HEIGHT - text_h) // 2 - 2
    draw.text((x, y - bbox[1]), now, fill=1, font=time_font)

    return img


# --- System Tray ---

running = True


def create_tray_icon():
    """Create a simple tray icon image (music note)."""
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    # Background circle
    draw.ellipse([4, 4, 60, 60], fill=(30, 30, 30, 255))
    # Music note shape
    draw.ellipse([14, 36, 30, 50], fill=(255, 255, 255, 255))
    draw.ellipse([34, 30, 50, 44], fill=(255, 255, 255, 255))
    draw.rectangle([28, 12, 32, 40], fill=(255, 255, 255, 255))
    draw.rectangle([48, 6, 52, 34], fill=(255, 255, 255, 255))
    draw.rectangle([28, 8, 52, 14], fill=(255, 255, 255, 255))
    return img


def on_quit(icon, item):
    """Handle quit from tray menu."""
    global running
    running = False
    icon.stop()


# --- Main Loop ---

def oled_loop(base_url):
    """Main OLED update loop running in a background thread."""
    global running

    last_heartbeat = 0
    last_title = None
    last_artist = None
    last_media_fetch = 0
    cached_info = None
    frame_interval = 0.1  # 10 FPS for smooth scrolling
    media_fetch_interval = 2  # fetch media info every 2 seconds

    # Volume tracking
    last_volume = None
    last_muted = None
    volume_display_until = 0  # timestamp when volume overlay should disappear
    volume_overlay_duration = 2  # seconds to show volume overlay

    while running:
        try:
            now = time.time()

            # Check volume every frame (cheap call)
            try:
                current_volume, current_muted = get_system_volume()
                if last_volume is None:
                    last_volume = current_volume
                    last_muted = current_muted
                elif abs(current_volume - last_volume) > 0.005 or current_muted != last_muted:
                    last_volume = current_volume
                    last_muted = current_muted
                    volume_display_until = now + volume_overlay_duration
            except Exception:
                pass

            # Show volume overlay if active
            if now < volume_display_until:
                frame = render_volume(last_volume, last_muted)
                bitmap = image_to_bitmap(frame)
                send_frame(base_url, bitmap)

                if now - last_heartbeat > 5:
                    send_heartbeat(base_url)
                    last_heartbeat = now

                time.sleep(frame_interval)
                continue

            # Fetch media info periodically (not every frame)
            if now - last_media_fetch >= media_fetch_interval:
                cached_info = asyncio.run(get_media_info())
                last_media_fetch = now

            # Status 4 = Playing in the SMTC enum
            is_playing = (cached_info and cached_info["title"]
                          and cached_info["status"] is not None
                          and cached_info["status"] == 4)

            if is_playing:
                if cached_info["title"] != last_title or cached_info["artist"] != last_artist:
                    last_title = cached_info["title"]
                    last_artist = cached_info["artist"]

                frame = render_now_playing(cached_info)
            else:
                if last_title is not None:
                    last_title = None
                    last_artist = None
                frame = render_idle()

            bitmap = image_to_bitmap(frame)
            send_frame(base_url, bitmap)

            if now - last_heartbeat > 5:
                send_heartbeat(base_url)
                last_heartbeat = now

            time.sleep(frame_interval)

        except Exception as e:
            log.error(f"OLED loop error: {e}", exc_info=True)
            time.sleep(1)

    cleanup(base_url)
    log.info("OLED loop stopped")


def main():
    log.info("Starting SteelSeries Now Playing")
    base_url = get_steelseries_address()
    log.info(f"SteelSeries Engine at {base_url}")
    register_game(base_url)
    bind_screen_event(base_url)
    log.info("Registered with GameSense SDK")

    # Create the system tray icon
    icon = pystray.Icon(
        "now_playing",
        create_tray_icon(),
        "SteelSeries Now Playing",
        menu=pystray.Menu(
            pystray.MenuItem("SteelSeries Now Playing", None, enabled=False),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit", on_quit),
        ),
    )

    # Run the OLED loop in a background thread
    oled_thread = threading.Thread(target=oled_loop, args=(base_url,), daemon=True)
    oled_thread.start()

    # Run tray icon on main thread (blocks until quit)
    icon.run()

    # Signal OLED thread to stop and wait
    global running
    running = False
    oled_thread.join(timeout=3)


if __name__ == "__main__":
    main()

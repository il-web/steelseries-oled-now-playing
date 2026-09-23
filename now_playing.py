"""
SteelSeries OLED - Now Playing Display
Shows current media info (title, artist, progress bar, album art) on the keyboard
OLED screen, and a clock, system stats or the weather when nothing is playing.
"""

import asyncio
import ctypes
import functools
import io
import json
import logging
import logging.handlers
import math
import os
import subprocess
import sys
import threading
import time
import unicodedata
import winreg
from datetime import datetime, timezone

# comtypes (used by pycaw for the volume) reads this when first imported. The
# default single-threaded COM apartment needs a message pump that the OLED
# loop doesn't run, which makes WinRT calls such as reading album art hang.
sys.coinit_flags = 0  # COINIT_MULTITHREADED

import psutil
import pystray
import requests
from PIL import Image, ImageDraw, ImageFont, ImageOps

try:
    from bidi.algorithm import get_display
except ImportError:  # python-bidi not installed: RTL text renders in logical order
    get_display = None

APP_NAME = "SteelSeries Now Playing"
APP_ID = "SteelSeriesNowPlaying"  # registry value, mutex and data folder name

# Settings and logs live in %APPDATA% so the script and the packaged .exe behave the same
DATA_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), APP_ID)
SETTINGS_PATH = os.path.join(DATA_DIR, "settings.json")
LOG_PATH = os.path.join(DATA_DIR, "now_playing.log")

log = logging.getLogger("now_playing")


def setup_logging():
    """Log to a file (the .pyw/.exe has no console), rotated so it can't grow forever."""
    handler = logging.handlers.RotatingFileHandler(
        LOG_PATH, maxBytes=1_000_000, backupCount=1, encoding="utf-8"
    )
    handlers = [handler]
    if sys.stderr is not None:  # also echo to the console when there is one
        handlers.append(logging.StreamHandler())
    formatter = logging.Formatter(
        "%(asctime)s %(levelname)s: %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    )
    for h in handlers:
        h.setFormatter(formatter)
    logging.basicConfig(level=logging.INFO, handlers=handlers)


# --- Settings ---

DEFAULT_SETTINGS = {
    "time_display": "elapsed",   # "elapsed" (1:23  3:45), "remaining" (1:23  -2:22) or "off"
    "album_art": False,
    "idle_screen": "clock",      # "clock", "stats", "weather" or "cycle"
    "cycle_seconds": 10,         # how long each idle screen shows in "cycle" mode
    "clock_24h": True,
    "clock_show_date": True,
    "volume_overlay": True,
    "startup_animation": True,
    "weather_city": "",
    "weather_units": "celsius",  # "celsius" or "fahrenheit"
    "scroll_speed": 2,           # pixels per frame
    "title_font_size": 11,
    "artist_font_size": 10,
}

SETTING_CHOICES = {
    "time_display": ("elapsed", "remaining", "off"),
    "idle_screen": ("clock", "stats", "weather", "cycle"),
    "weather_units": ("celsius", "fahrenheit"),
}


def valid_setting(key, value):
    """Reject hand-edited values of the wrong type or outside the allowed choices."""
    default = DEFAULT_SETTINGS[key]
    if type(value) is not type(default):  # noqa: E721 - bool must not pass as int
        return False
    if key in SETTING_CHOICES:
        return value in SETTING_CHOICES[key]
    if isinstance(default, bool):  # before the int check: bool is an int subclass
        return True
    if isinstance(default, int):
        return 1 <= value <= 60
    return True


class Settings:
    """JSON-backed settings: written by the tray menu, read by the OLED loop.

    Until open() is called it only holds the defaults (nothing touches disk).
    """

    def __init__(self):
        self.path = None
        self._values = dict(DEFAULT_SETTINGS)
        self._lock = threading.Lock()
        self._mtime = None

    def open(self, path):
        self.path = path
        if os.path.exists(path):
            self.load()
        else:
            self.save()

    def load(self):
        try:
            with open(self.path, encoding="utf-8") as f:
                loaded = json.load(f)
            self._mtime = os.path.getmtime(self.path)
        except (OSError, ValueError) as e:
            log.warning(f"Could not read settings ({e}); keeping current values")
            return
        values = dict(DEFAULT_SETTINGS)
        for key, value in loaded.items():
            if key in DEFAULT_SETTINGS and valid_setting(key, value):
                values[key] = value
            else:
                log.warning(f"Ignoring invalid setting {key!r}: {value!r}")
        with self._lock:
            self._values = values

    def reload_if_changed(self):
        """Pick up edits made by hand in the settings file."""
        try:
            mtime = os.path.getmtime(self.path)
        except (OSError, TypeError):
            return
        if mtime != self._mtime:
            log.info("Settings file changed; reloading")
            self.load()

    def save(self):
        with self._lock:
            data = json.dumps(self._values, indent=2)
        try:
            tmp_path = self.path + ".tmp"
            with open(tmp_path, "w", encoding="utf-8") as f:
                f.write(data + "\n")
            os.replace(tmp_path, self.path)
            self._mtime = os.path.getmtime(self.path)
        except OSError as e:
            log.warning(f"Could not save settings: {e}")

    def __getitem__(self, key):
        return self._values[key]

    def set(self, key, value):
        with self._lock:
            self._values[key] = value
        if self.path:
            self.save()


settings = Settings()


# --- SteelSeries GameSense SDK ---

GAME_NAME = "NOWPLAYING"
EVENT_NAME = "SCREEN"
OLED_WIDTH = 128
OLED_HEIGHT = 40
REQUEST_TIMEOUT = 2  # seconds; never let a stuck Engine hang the loop

# Reuse one HTTP connection instead of opening a new one every frame
http = requests.Session()


def get_steelseries_address():
    """Read the SteelSeries Engine address from coreProps.json (None if unavailable)."""
    props_path = os.path.join(
        os.environ.get("PROGRAMDATA", "C:\\ProgramData"),
        "SteelSeries", "SteelSeries Engine 3", "coreProps.json"
    )
    try:
        with open(props_path, "r") as f:
            props = json.load(f)
        return f"http://{props['address']}"
    except FileNotFoundError:
        log.warning(f"SteelSeries GG/Engine 3 not found (expected {props_path})")
    except (json.JSONDecodeError, KeyError):
        log.warning("Could not parse SteelSeries coreProps.json")
    return None


def connect():
    """Wait for SteelSeries Engine and register with it.

    The Engine picks a new port every time it restarts, so this is called at
    startup and again whenever a request fails. Returns the base URL, or None
    if the app is quitting.
    """
    retry_delay = 2
    while not stop_event.is_set():
        base_url = get_steelseries_address()
        if base_url:
            try:
                register_game(base_url)
                bind_screen_event(base_url)
                log.info(f"Registered with GameSense SDK at {base_url}")
                return base_url
            except requests.RequestException as e:
                log.warning(f"SteelSeries Engine not reachable at {base_url}: {e}")
        stop_event.wait(retry_delay)
        retry_delay = min(retry_delay * 2, 30)
    return None


def register_game(base_url):
    """Register our app with SteelSeries Engine."""
    resp = http.post(f"{base_url}/game_metadata", timeout=REQUEST_TIMEOUT, json={
        "game": GAME_NAME,
        "game_display_name": "Now Playing",
        "developer": "il-web",
    })
    resp.raise_for_status()


def bind_screen_event(base_url):
    """Bind a screen handler for bitmap mode on the 128x40 OLED."""
    resp = http.post(f"{base_url}/bind_game_event", timeout=REQUEST_TIMEOUT, json={
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
    # Mode "1" packs 8 pixels per byte, MSB = leftmost, row by row: exactly
    # the layout the OLED expects (128 / 8 = 16 bytes per row, no padding).
    return list(img.tobytes())


def send_frame(base_url, bitmap_data):
    """Send a bitmap frame to the OLED."""
    resp = http.post(f"{base_url}/game_event", timeout=REQUEST_TIMEOUT, json={
        "game": GAME_NAME,
        "event": EVENT_NAME,
        "data": {
            "value": 0,
            "frame": {
                "image-data-128x40": bitmap_data,
            }
        }
    })
    resp.raise_for_status()


def send_heartbeat(base_url):
    """Prevent the 15-second timeout."""
    resp = http.post(f"{base_url}/game_heartbeat", timeout=REQUEST_TIMEOUT,
                     json={"game": GAME_NAME})
    resp.raise_for_status()


def stop_game(base_url):
    """Hand the OLED back to SteelSeries GG (used while the display is paused)."""
    try:
        http.post(f"{base_url}/stop_game", timeout=REQUEST_TIMEOUT, json={"game": GAME_NAME})
    except Exception:
        pass


def cleanup(base_url):
    """Remove our game from SteelSeries Engine."""
    try:
        http.post(f"{base_url}/remove_game", timeout=REQUEST_TIMEOUT,
                  json={"game": GAME_NAME})
    except Exception:
        pass


# --- Windows Volume ---

_endpoint_volume = None
_endpoint_fetched_at = 0
ENDPOINT_REFRESH_INTERVAL = 5  # re-resolve the default device (e.g. headset switch)


def get_system_volume():
    """Get the current system volume (0.0 - 1.0) and mute state."""
    global _endpoint_volume, _endpoint_fetched_at
    from pycaw.pycaw import AudioUtilities

    now = time.time()
    if _endpoint_volume is None or now - _endpoint_fetched_at > ENDPOINT_REFRESH_INTERVAL:
        _endpoint_volume = AudioUtilities.GetSpeakers().EndpointVolume
        _endpoint_fetched_at = now
    try:
        level = _endpoint_volume.GetMasterVolumeLevelScalar()  # 0.0 to 1.0
        muted = _endpoint_volume.GetMute()
    except Exception:
        _endpoint_volume = None  # device went away; re-resolve next call
        raise
    return level, bool(muted)


# --- Windows Media Info ---

ART_SIZE = OLED_HEIGHT  # album art is a square filling the screen height
MEDIA_QUERY_TIMEOUT = 5  # seconds

# Album art for the current track, so it is only downloaded once per song
_art_cache = {"key": None, "image": None}


def prepare_album_art(data):
    """Turn encoded cover art into a dithered 1-bit square for the OLED."""
    img = Image.open(io.BytesIO(data)).convert("L")
    img = ImageOps.fit(img, (ART_SIZE, ART_SIZE), Image.Resampling.LANCZOS)
    img = ImageOps.autocontrast(img, cutoff=2)
    return img.convert("1")  # Floyd-Steinberg dithering


async def read_thumbnail(thumbnail_ref):
    """Read a WinRT stream reference (the SMTC thumbnail) into bytes."""
    from winrt.windows.storage.streams import Buffer, InputStreamOptions

    stream = await thumbnail_ref.open_read_async()
    size = stream.size
    buffer = Buffer(size)
    await stream.read_async(buffer, size, InputStreamOptions.READ_AHEAD)
    return bytes(buffer)


async def get_album_art(media_props, key):
    """Album art for the current track (cached), or None."""
    # Players often publish the thumbnail a moment after the title, so keep
    # retrying while the cached image for this track is still missing.
    if _art_cache["key"] != key or _art_cache["image"] is None:
        image = None
        try:
            if media_props.thumbnail is not None:
                image = prepare_album_art(await read_thumbnail(media_props.thumbnail))
        except Exception as e:
            log.debug(f"Could not read album art: {e}")
        _art_cache.update(key=key, image=image)
    return _art_cache["image"]


async def get_media_info(fetch_art=False):
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
        duration_secs = (timeline.end_time - timeline.start_time).total_seconds()
        # When the player last reported its position. Many players (Spotify,
        # browsers) only report on play/pause/seek, so we extrapolate from it.
        updated = timeline.last_updated_time
        if updated.tzinfo is None:
            updated = updated.replace(tzinfo=timezone.utc)
        position_updated_at = updated.timestamp()
    except Exception:
        position_secs = 0
        duration_secs = 0
        position_updated_at = None

    art = await get_album_art(media_props, (title, artist)) if fetch_art else None

    return {
        "title": title,
        "artist": artist,
        "status": status,
        "position": position_secs,
        "position_updated_at": position_updated_at,
        "duration": duration_secs,
        "art": art,
    }


class MediaWatcher:
    """Runs media queries on their own thread, one at a time.

    The OLED loop only reads the latest result, so a slow or stuck Windows
    media call can never freeze the display.
    """

    def __init__(self):
        self.info = None
        self._busy = False
        self._started_at = 0
        self._stuck_logged = False

    def refresh(self, fetch_art):
        if self._busy:
            if not self._stuck_logged and time.time() - self._started_at > MEDIA_QUERY_TIMEOUT:
                log.warning("Media info query is stuck; showing the last known info")
                self._stuck_logged = True
            return
        self._busy = True
        self._started_at = time.time()
        threading.Thread(target=self._query, args=(fetch_art,), daemon=True).start()

    def _query(self, fetch_art):
        try:
            self.info = asyncio.run(get_media_info(fetch_art=fetch_art))
        except Exception as e:
            log.error(f"Media info query failed: {e}", exc_info=True)
        finally:
            self._busy = False
            self._stuck_logged = False


media_watcher = MediaWatcher()


def current_position(info, is_playing):
    """Estimate the live playback position from the last reported one."""
    position = info["position"]
    updated_at = info.get("position_updated_at")
    if is_playing and updated_at:
        elapsed = time.time() - updated_at
        # Ignore bogus timestamps (unset/zero dates, clock skew)
        if 0 <= elapsed <= info["duration"]:
            position += elapsed
    return min(position, info["duration"])


# --- System Stats ---

class SystemStats:
    """CPU / RAM / network usage, sampled at most once per second."""

    def __init__(self):
        self.cpu = 0.0
        self.ram = 0.0
        self.down = 0.0  # bytes per second
        self.up = 0.0
        self._sampled_at = 0
        self._net = None

    def update(self):
        now = time.time()
        if now - self._sampled_at < 1:
            return
        net = psutil.net_io_counters()
        if self._net is not None:
            elapsed = now - self._sampled_at
            self.down = max(net.bytes_recv - self._net.bytes_recv, 0) / elapsed
            self.up = max(net.bytes_sent - self._net.bytes_sent, 0) / elapsed
        self._net = net
        self._sampled_at = now
        self.cpu = psutil.cpu_percent(interval=None)  # since the previous call
        self.ram = psutil.virtual_memory().percent


system_stats = SystemStats()


# --- Weather (Open-Meteo, no API key needed) ---

GEOCODE_URL = "https://geocoding-api.open-meteo.com/v1/search"
FORECAST_URL = "https://api.open-meteo.com/v1/forecast"

# WMO weather interpretation codes -> short description
WEATHER_CODES = [
    ((0,), "Clear"),
    ((1,), "Mostly clear"),
    ((2,), "Partly cloudy"),
    ((3,), "Overcast"),
    ((45, 48), "Fog"),
    ((51, 53, 55, 56, 57), "Drizzle"),
    ((61, 63, 65, 66, 67), "Rain"),
    ((71, 73, 75, 77), "Snow"),
    ((80, 81, 82), "Showers"),
    ((85, 86), "Snow showers"),
    ((95, 96, 99), "Thunderstorm"),
]


def describe_weather(code):
    for codes, text in WEATHER_CODES:
        if code in codes:
            return text
    return ""


class WeatherError(Exception):
    """A weather problem worth showing on the screen as-is."""


class Weather:
    """Current weather for the configured city, fetched in the background."""

    REFRESH_INTERVAL = 15 * 60
    RETRY_INTERVAL = 60

    def __init__(self):
        self.data = None
        self.error = None
        self._key = None           # (city, units) the data/error belongs to
        self._next_fetch = 0
        self._fetching = False
        self._place = None         # cached geocoding result: (query, name, lat, lon)

    def get(self):
        """Return (data, error), starting a background refresh when stale."""
        key = (settings["weather_city"].strip(), settings["weather_units"])
        if key != self._key:
            self.data = None
            self.error = None
            self._next_fetch = 0
        if not self._fetching and time.time() >= self._next_fetch:
            self._fetching = True
            threading.Thread(target=self._fetch, args=(key,), daemon=True).start()
        return self.data, self.error

    def _fetch(self, key):
        city, units = key
        retry_after = self.REFRESH_INTERVAL
        data, error = None, None
        try:
            if not city:
                raise WeatherError("Set a city in the tray menu")
            if self._place is None or self._place[0] != city:
                resp = requests.get(GEOCODE_URL, params={"name": city, "count": 1},
                                    timeout=10)
                resp.raise_for_status()
                results = resp.json().get("results")
                if not results:
                    raise WeatherError(f"City not found: {city}")
                self._place = (city, results[0]["name"],
                               results[0]["latitude"], results[0]["longitude"])
            _, name, lat, lon = self._place
            resp = requests.get(FORECAST_URL, timeout=10, params={
                "latitude": lat,
                "longitude": lon,
                "current": "temperature_2m,weather_code",
                "daily": "temperature_2m_max,temperature_2m_min",
                "temperature_unit": units,
                "timezone": "auto",
                "forecast_days": 1,
            })
            resp.raise_for_status()
            forecast = resp.json()
            data = {
                "place": name,
                "temp": forecast["current"]["temperature_2m"],
                "condition": describe_weather(forecast["current"]["weather_code"]),
                "high": forecast["daily"]["temperature_2m_max"][0],
                "low": forecast["daily"]["temperature_2m_min"][0],
                "unit": "F" if units == "fahrenheit" else "C",
            }
        except WeatherError as e:
            error = str(e)
        except Exception as e:
            log.warning(f"Weather fetch failed: {e}")
            error = "Weather unavailable"
            retry_after = self.RETRY_INTERVAL
        # Keep showing the last good data through a temporary network error
        if data is not None or key != self._key or self.data is None:
            self.data = data
        self.error = error
        self._key = key
        self._next_fetch = time.time() + retry_after
        self._fetching = False


weather = Weather()


# --- Fonts ---

FONT_DIR = os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts")
PRIMARY_FONTS = ["segoeui.ttf", "arial.ttf", "tahoma.ttf"]
# Tried in order for characters the primary font can't draw
JAPANESE_FONTS = ["YuGothM.ttc", "meiryo.ttc", "msgothic.ttc"]
CHINESE_FONTS = ["msyh.ttc", "simsun.ttc"]
KOREAN_FONTS = ["malgun.ttf", "gulim.ttc"]
OTHER_FALLBACK_FONTS = ["Nirmala.ttc", "Nirmala.ttf", "LeelawUI.ttf",
                        "seguisym.ttf", "seguiemj.ttf"]


@functools.lru_cache(maxsize=None)
def font_codepoints(name):
    """The set of characters a font file can draw (empty if it's missing)."""
    from fontTools.ttLib import TTCollection, TTFont

    path = os.path.join(FONT_DIR, name)
    if not os.path.exists(path):
        return frozenset()
    try:
        if name.lower().endswith(".ttc"):
            collection = TTCollection(path, lazy=True)
            codepoints = frozenset(collection.fonts[0].getBestCmap() or ())
            collection.close()
        else:
            font = TTFont(path, lazy=True)
            codepoints = frozenset(font.getBestCmap() or ())
            font.close()
        return codepoints
    except Exception as e:
        log.warning(f"Could not read font {name}: {e}")
        return frozenset()


@functools.lru_cache(maxsize=None)
def load_font(name, size):
    return ImageFont.truetype(os.path.join(FONT_DIR, name), size)


@functools.lru_cache(maxsize=None)
def primary_font_name():
    for name in PRIMARY_FONTS:
        if os.path.exists(os.path.join(FONT_DIR, name)):
            return name
    return None


@functools.lru_cache(maxsize=None)
def get_font(size):
    """The primary UI font at a pixel size. Cached: called every frame."""
    name = primary_font_name()
    if name:
        return load_font(name, size)
    return ImageFont.load_default(size)


@functools.lru_cache(maxsize=None)
def get_small_font():
    """Font for small numbers (times, stats). Tahoma is hinted for tiny pixel
    sizes; Segoe UI at 9px drops colons and dots ("1:23" shows as "1 23")."""
    if os.path.exists(os.path.join(FONT_DIR, "tahoma.ttf")):
        return load_font("tahoma.ttf", 9)
    return get_font(10)


def is_kana(ch):
    return 0x3040 <= ord(ch) <= 0x30FF or 0x31F0 <= ord(ch) <= 0x31FF


def fallback_chain(text):
    """Fonts to try, in order. Chinese characters are shared with Japanese,
    so prefer a Japanese font when the text also contains kana."""
    if any(is_kana(ch) for ch in text):
        cjk = JAPANESE_FONTS + CHINESE_FONTS
    else:
        cjk = CHINESE_FONTS + JAPANESE_FONTS
    names = [primary_font_name()] + cjk + KOREAN_FONTS + OTHER_FALLBACK_FONTS
    return [name for name in names if name and font_codepoints(name)]


@functools.lru_cache(maxsize=256)
def text_runs(text, size):
    """Split text into (font, substring) runs so every character gets a font
    that can draw it (Segoe UI has no Japanese, Chinese, Korean or emoji)."""
    chain = fallback_chain(text)
    if not chain:
        return ((get_font(size), text),)
    runs = []
    current, chunk = None, ""
    for ch in text:
        # Stay in the current font when it has the glyph (spaces, punctuation)
        if current is None or ord(ch) not in font_codepoints(current):
            name = next((n for n in chain if ord(ch) in font_codepoints(n)), current or chain[0])
            if name != current and chunk:
                runs.append((current, chunk))
                chunk = ""
            current = name
        chunk += ch
    if chunk:
        runs.append((current, chunk))
    return tuple((load_font(name, size), chunk) for name, chunk in runs)


def to_visual_order(text):
    """Reorder right-to-left text (Hebrew, Arabic) for display.

    Pillow on Windows has no libraqm/FriBiDi, so it draws characters in
    logical order, which shows RTL words backwards.
    """
    if get_display is None:
        return text
    return get_display(text)


def is_rtl(text):
    """True if the text's first strongly-directional character is right-to-left."""
    for ch in text:
        direction = unicodedata.bidirectional(ch)
        if direction in ("R", "AL"):
            return True
        if direction == "L":
            return False
    return False


def layout_text(text, size):
    """Runs for arbitrary (e.g. song title) text: RTL reordering plus font fallback."""
    return text_runs(to_visual_order(text), size)


def runs_width(runs):
    return sum(font.getlength(chunk) for font, chunk in runs)


def draw_runs(draw, x, y, runs):
    """Draw text runs with the top of the tallest glyph at y."""
    tops = [font.getbbox(chunk, anchor="ls")[1] for font, chunk in runs if chunk.strip()]
    baseline = y - min(tops, default=0)
    for font, chunk in runs:
        draw.text((x, baseline), chunk, fill=1, font=font, anchor="ls")
        x += font.getlength(chunk)


def draw_text(img, x, y, text, size, align="left", max_width=None):
    """Draw arbitrary text with its top at y. align is relative to x."""
    runs = layout_text(text, size)
    width = runs_width(runs)
    if align == "center":
        x -= width / 2
    elif align == "right":
        x -= width
    if max_width is not None and width > max_width:
        # Clip instead of overflowing into neighbouring content
        clip = Image.new("1", (int(max_width), OLED_HEIGHT), 0)
        draw_runs(ImageDraw.Draw(clip), 0, y, runs)
        img.paste(clip, (int(x), 0), clip)
        return
    draw_runs(ImageDraw.Draw(img), x, y, runs)


# --- Rendering ---

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
        # The available width can change (album art toggled)
        self.scroll_offset = min(self.scroll_offset, max_scroll)

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


def draw_scrolling_text(img, scroller, text, size, x, y, max_width, line_height):
    """Draw text with horizontal scrolling if it exceeds max_width."""
    runs = layout_text(text, size)
    scroller.update_text(text)
    scroller.scroll_speed = settings["scroll_speed"]
    scroller.text_width = runs_width(runs)
    scroller.tick(max_width)

    if is_rtl(text):
        # Right-to-left text begins at its right edge: align it right and
        # scroll the other way, so the start of the title shows first
        text_x = max_width - scroller.text_width + scroller.scroll_offset
    else:
        text_x = -scroller.scroll_offset

    # Draw into a temporary image with fixed line height for clipping
    text_img = Image.new("1", (max_width, line_height), 0)
    draw_runs(ImageDraw.Draw(text_img), text_x, 0, runs)

    # Paste the clipped text onto the main image
    img.paste(text_img, (x, y))


def format_time(seconds):
    """Format seconds as m:ss, or h:mm:ss for long media."""
    seconds = max(int(seconds), 0)
    hours, rest = divmod(seconds, 3600)
    minutes, secs = divmod(rest, 60)
    if hours:
        return f"{hours}:{minutes:02}:{secs:02}"
    return f"{minutes}:{secs:02}"


def draw_progress(draw, duration, position, left, right):
    """Progress bar along the bottom, optionally with times on either side."""
    mode = settings["time_display"]
    bar_y, bar_height = 34, 5

    if mode != "off" and duration > 0:
        font = get_small_font()
        elapsed_str = format_time(position)
        if mode == "remaining":
            total_str = "-" + format_time(duration - position)
        else:
            total_str = format_time(duration)
        draw.text((left, OLED_HEIGHT - 1), elapsed_str, fill=1, font=font, anchor="ls")
        draw.text((right + 1, OLED_HEIGHT - 1), total_str, fill=1, font=font, anchor="rs")
        left = int(left + font.getlength(elapsed_str) + 3)
        right = int(right - font.getlength(total_str) - 3)
        bar_y, bar_height = 33, 4

    # Bar outline
    draw.rectangle([left, bar_y, right, bar_y + bar_height], outline=1, fill=0)

    # Bar fill
    if duration > 0:
        progress = min(position / duration, 1.0)
        fill_width = int((right - left - 2) * progress)
        if fill_width > 0:
            draw.rectangle(
                [left + 1, bar_y + 1, left + 1 + fill_width, bar_y + bar_height - 1],
                fill=1,
            )


def render_now_playing(info, position):
    """Render the now-playing screen as a PIL Image."""
    img = Image.new("1", (OLED_WIDTH, OLED_HEIGHT), 0)
    draw = ImageDraw.Draw(img)

    text_x = 1
    art = info.get("art")
    if art is not None and settings["album_art"]:
        img.paste(art, (0, 0))
        text_x = ART_SIZE + 3

    title = info["title"] if info["title"] else "Unknown"
    artist = info["artist"] if info["artist"] else "Unknown"

    max_text_width = OLED_WIDTH - text_x - 1
    title_height = settings["title_font_size"] + 2
    artist_y = title_height + 1
    artist_height = max(min(settings["artist_font_size"] + 4, 31 - artist_y), 1)

    # Draw title (line 1) - scrolls if too long
    draw_scrolling_text(img, title_scroller, title, settings["title_font_size"],
                        text_x, 0, max_text_width, title_height)

    # Draw artist (line 2) - scrolls if too long
    draw_scrolling_text(img, artist_scroller, artist, settings["artist_font_size"],
                        text_x, artist_y, max_text_width, artist_height)

    # Progress bar (pushed to bottom)
    draw_progress(draw, info["duration"], position, text_x, OLED_WIDTH - 2)

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
        fill_width = round((bar_right - bar_left - 2) * volume_pct)
        if fill_width > 0:
            draw.rectangle(
                [bar_left + 1, bar_y + 1, bar_left + 1 + fill_width, bar_y + bar_height - 1],
                fill=1,
            )

    # Percentage text
    # round(), not int(): Windows reports e.g. 0.4999999 for 50%
    pct_str = f"{round(volume_pct * 100)}%"
    bbox = draw.textbbox((0, 0), pct_str, font=pct_font)
    pct_w = bbox[2] - bbox[0]
    draw.text(((OLED_WIDTH - pct_w) // 2, 28 - bbox[1]), pct_str, fill=1, font=pct_font)

    return img


def render_clock():
    """Render the current time (and optionally the date)."""
    img = Image.new("1", (OLED_WIDTH, OLED_HEIGHT), 0)
    draw = ImageDraw.Draw(img)

    now = datetime.now()
    if settings["clock_24h"]:
        time_str = now.strftime("%H:%M")
    else:
        time_str = now.strftime("%I:%M %p").lstrip("0")

    time_font = get_font(18)
    bbox = draw.textbbox((0, 0), time_str, font=time_font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    x = (OLED_WIDTH - text_w) // 2 - bbox[0]

    if settings["clock_show_date"]:
        date_str = now.strftime("%a %d %b")
        date_font = get_font(10)
        date_bbox = draw.textbbox((0, 0), date_str, font=date_font)
        date_w = date_bbox[2] - date_bbox[0]
        date_h = date_bbox[3] - date_bbox[1]
        gap = 5
        top = (OLED_HEIGHT - (text_h + gap + date_h)) // 2
        draw.text((x, top - bbox[1]), time_str, fill=1, font=time_font)
        draw.text(((OLED_WIDTH - date_w) // 2 - date_bbox[0], top + text_h + gap - date_bbox[1]),
                  date_str, fill=1, font=date_font)
    else:
        y = (OLED_HEIGHT - text_h) // 2 - 2
        draw.text((x, y - bbox[1]), time_str, fill=1, font=time_font)

    return img


def format_rate(bytes_per_sec):
    """Compact transfer rate: 0 B/s, 40 KB/s, 1.2 MB/s."""
    for unit in ("B", "KB", "MB"):
        if bytes_per_sec < 1000:
            if unit == "B" or bytes_per_sec >= 10:
                return f"{bytes_per_sec:.0f} {unit}/s"
            return f"{bytes_per_sec:.1f} {unit}/s"
        bytes_per_sec /= 1000
    return f"{bytes_per_sec:.1f} GB/s"


def render_stats(stats):
    """Render CPU and RAM usage bars plus network throughput."""
    img = Image.new("1", (OLED_WIDTH, OLED_HEIGHT), 0)
    draw = ImageDraw.Draw(img)
    font = get_small_font()

    label_width = 24
    pct_width = 24
    for row, (label, pct) in enumerate((("CPU", stats.cpu), ("RAM", stats.ram))):
        baseline = 9 + row * 13
        draw.text((1, baseline), label, fill=1, font=font, anchor="ls")
        draw.text((OLED_WIDTH - 1, baseline), f"{pct:.0f}%", fill=1, font=font, anchor="rs")
        bar_left = label_width
        bar_right = OLED_WIDTH - pct_width - 2
        bar_top = baseline - 7
        draw.rectangle([bar_left, bar_top, bar_right, bar_top + 7], outline=1, fill=0)
        fill_width = int((bar_right - bar_left - 2) * min(pct, 100) / 100)
        if fill_width > 0:
            draw.rectangle([bar_left + 1, bar_top + 1, bar_left + 1 + fill_width, bar_top + 6],
                           fill=1)

    # Network: download on the left, upload on the right
    net_baseline = OLED_HEIGHT - 3
    down_str = format_rate(stats.down)
    up_str = format_rate(stats.up)
    up_x = OLED_WIDTH - 1 - int(font.getlength(up_str)) - 8
    draw_arrow(draw, 1, net_baseline, up=False)
    draw.text((8, net_baseline), down_str, fill=1, font=font, anchor="ls")
    draw_arrow(draw, up_x, net_baseline, up=True)
    draw.text((OLED_WIDTH - 1, net_baseline), up_str, fill=1, font=font, anchor="rs")
    return img


def draw_arrow(draw, x, baseline, up):
    """A 5x7 pixel arrow (the font's arrow glyphs are unreadable at this size)."""
    top, bottom = baseline - 7, baseline - 1
    draw.line([x + 2, top, x + 2, bottom], fill=1)
    if up:
        draw.line([x, top + 2, x + 2, top, x + 4, top + 2], fill=1)
    else:
        draw.line([x, bottom - 2, x + 2, bottom, x + 4, bottom - 2], fill=1)


def render_weather(data, error):
    """Render the current weather, or a status message."""
    img = Image.new("1", (OLED_WIDTH, OLED_HEIGHT), 0)
    draw = ImageDraw.Draw(img)

    if data is None:
        message = error or "Loading weather..."
        draw_text(img, OLED_WIDTH / 2, 7, "Weather", 11, align="center")
        draw_text(img, OLED_WIDTH / 2, 24, message, 10, align="center",
                  max_width=OLED_WIDTH - 2)
        return img

    draw_text(img, 1, 0, data["place"], 10, max_width=OLED_WIDTH - 2)

    temp_str = f"{data['temp']:.0f}\u00b0{data['unit']}"
    temp_font = get_font(18)
    draw.text((1, OLED_HEIGHT - 3), temp_str, fill=1, font=temp_font, anchor="ls")

    column_x = int(temp_font.getlength(temp_str)) + 8
    column_width = OLED_WIDTH - column_x - 1
    draw_text(img, column_x, 16, data["condition"], 10, max_width=column_width)
    draw.text((column_x, OLED_HEIGHT - 3), f"H {data['high']:.0f}\u00b0  L {data['low']:.0f}\u00b0",
              fill=1, font=get_small_font(), anchor="ls")
    return img


STARTUP_FPS = 20
STARTUP_FRAMES = 44  # 2.2 seconds


def ease(t):
    """Smooth 0..1 -> 0..1 (ease in-out)."""
    t = min(max(t, 0.0), 1.0)
    return t * t * (3 - 2 * t)


def lerp(a, b, t):
    return a + (b - a) * t


def render_startup_frame(i):
    """Frame i of the startup animation: bouncing equalizer bars that slide
    left while "Now Playing" wipes in and a loading bar fills."""
    img = Image.new("1", (OLED_WIDTH, OLED_HEIGHT), 0)
    draw = ImageDraw.Draw(img)

    # Phases (in frames): bars alone -> bars slide left -> text wipes in
    # -> loading bar fills -> short hold on the finished frame
    move = ease((i - 10) / 8)
    reveal_t = ease((i - 16) / 9)
    fill = ease((i - 24) / 14)

    # Equalizer geometry, interpolated from big + centred to small + left
    bar_count = 5
    bar_width = round(lerp(5, 3, move))
    gap = round(lerp(3, 2, move))
    max_height = lerp(30, 20, move)
    bottom = round(lerp(36, 31, move))
    group_width = bar_count * bar_width + (bar_count - 1) * gap
    left = round(lerp((OLED_WIDTH - group_width) / 2, 6, move))
    grow = min(i / 5, 1.0)  # bars rise from nothing at the start

    for k in range(bar_count):
        bounce = 0.25 + 0.75 * abs(math.sin(i * 0.45 + k * 1.3))
        height = max(round(max_height * bounce * grow), 1)
        x = left + k * (bar_width + gap)
        draw.rectangle([x, bottom - height + 1, x + bar_width - 1, bottom], fill=1)

    # "Now Playing" revealed left to right
    if reveal_t > 0:
        text_x = 34
        font = get_font(14)
        text = "Now Playing"
        text_img = Image.new("1", (OLED_WIDTH - text_x, OLED_HEIGHT), 0)
        ImageDraw.Draw(text_img).text((0, 22), text, fill=1, font=font, anchor="ls")
        reveal = round(font.getlength(text) * reveal_t) + 1
        visible = text_img.crop((0, 0, reveal, OLED_HEIGHT))
        img.paste(visible, (text_x, 0), visible)  # masked: don't erase the bars

        # Loading bar under the text
        if fill > 0:
            bar_right = text_x + round((OLED_WIDTH - 4 - text_x) * fill)
            draw.rectangle([text_x, 29, bar_right, 30], fill=1)

    return img


def play_startup_animation(base_url):
    """Show the startup animation once (about 2 seconds)."""
    start = time.monotonic()
    for i in range(STARTUP_FRAMES):
        if stop_event.is_set() or display_paused.is_set():
            return
        send_frame(base_url, image_to_bitmap(render_startup_frame(i)))
        # Wait until this frame's slot ends, so sending time doesn't slow the animation
        stop_event.wait(max(start + (i + 1) / STARTUP_FPS - time.monotonic(), 0))


IDLE_SCREENS = ("clock", "stats", "weather")


def render_idle():
    """Render the idle screen chosen in the tray menu."""
    screen = settings["idle_screen"]
    if screen == "cycle":
        # Skip the weather until a city is set
        screens = [s for s in IDLE_SCREENS
                   if s != "weather" or settings["weather_city"].strip()]
        screen = screens[int(time.time() // settings["cycle_seconds"]) % len(screens)]

    if screen == "stats":
        system_stats.update()
        return render_stats(system_stats)
    if screen == "weather":
        return render_weather(*weather.get())
    return render_clock()


# --- Windows integration ---

RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
ERROR_ALREADY_EXISTS = 183
_instance_mutex = None


def acquire_single_instance():
    """Return False if another copy is already running.

    Holds a named mutex for the life of the process; Windows frees it on exit,
    even after a crash.
    """
    global _instance_mutex
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateMutexW.restype = ctypes.c_void_p
    _instance_mutex = kernel32.CreateMutexW(None, False, f"Local\\{APP_ID}")
    return ctypes.get_last_error() != ERROR_ALREADY_EXISTS


def show_message(text):
    """Show a message box (there's no console to print to)."""
    MB_ICONINFORMATION, MB_SETFOREGROUND = 0x40, 0x10000
    ctypes.windll.user32.MessageBoxW(None, text, APP_NAME, MB_ICONINFORMATION | MB_SETFOREGROUND)


def startup_command():
    """The command Windows should run at login to start this app."""
    if getattr(sys, "frozen", False):  # packaged .exe
        return f'"{sys.executable}"'
    pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    if not os.path.exists(pythonw):
        pythonw = sys.executable
    return f'"{pythonw}" "{os.path.abspath(__file__)}"'


def get_startup_entry():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as key:
            return winreg.QueryValueEx(key, APP_ID)[0]
    except OSError:
        return None


def set_startup_enabled(enabled):
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as key:
        if enabled:
            winreg.SetValueEx(key, APP_ID, 0, winreg.REG_SZ, startup_command())
        else:
            try:
                winreg.DeleteValue(key, APP_ID)
            except FileNotFoundError:
                pass


def refresh_startup_entry():
    """Keep the login entry pointing here if the app was moved since it was enabled."""
    entry = get_startup_entry()
    if entry is not None and entry != startup_command():
        log.info("Updating Start with Windows entry to the current location")
        set_startup_enabled(True)


# --- System Tray ---

# Set when the user quits; wait()-able so sleeps end immediately on quit
stop_event = threading.Event()
# Set while the user has paused the display (the OLED goes back to GG)
display_paused = threading.Event()
_dialog_lock = threading.Lock()


def create_tray_icon(size=64, paused=False):
    """Create a simple tray icon image (music note); grey while paused."""
    s = size / 64
    note = (130, 130, 130, 255) if paused else (255, 255, 255, 255)
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    def box(x0, y0, x1, y1):
        return [x0 * s, y0 * s, x1 * s, y1 * s]

    # Background circle
    draw.ellipse(box(4, 4, 60, 60), fill=(30, 30, 30, 255))
    # Music note shape
    draw.ellipse(box(14, 36, 30, 50), fill=note)
    draw.ellipse(box(34, 30, 50, 44), fill=note)
    draw.rectangle(box(28, 12, 32, 40), fill=note)
    draw.rectangle(box(48, 6, 52, 34), fill=note)
    draw.rectangle(box(28, 8, 52, 14), fill=note)
    return img


def on_quit(icon, item):
    """Handle quit from tray menu."""
    stop_event.set()
    icon.stop()


def on_toggle_pause(icon, item):
    if display_paused.is_set():
        display_paused.clear()
    else:
        display_paused.set()
    paused = display_paused.is_set()
    icon.icon = create_tray_icon(paused=paused)
    icon.title = f"{APP_NAME} (paused)" if paused else APP_NAME


def on_toggle_startup(icon, item):
    try:
        set_startup_enabled(get_startup_entry() is None)
    except OSError as e:
        log.error(f"Could not change Start with Windows: {e}")
        show_message(f"Could not change the Start with Windows setting:\n{e}")


def on_open_settings(icon, item):
    subprocess.Popen(["notepad.exe", settings.path])


def on_open_log(icon, item):
    subprocess.Popen(["notepad.exe", LOG_PATH])


def ask_weather_city():
    """Prompt for the weather city (runs on its own thread; Tk lives only here)."""
    if not _dialog_lock.acquire(blocking=False):
        return  # a dialog is already open
    try:
        import tkinter as tk
        from tkinter import simpledialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        city = simpledialog.askstring(
            "Weather city", "City name (e.g. London, Tel Aviv):",
            initialvalue=settings["weather_city"], parent=root,
        )
        root.destroy()
        if city is not None:
            settings.set("weather_city", city.strip())
    except Exception as e:
        log.error(f"Weather city dialog failed: {e}", exc_info=True)
    finally:
        _dialog_lock.release()


def on_set_weather_city(icon, item):
    threading.Thread(target=ask_weather_city, daemon=True).start()


def build_menu():
    """The tray menu; checkmarks read the live settings every time it opens."""

    def choice(text, key, value):
        def select(icon, item):
            settings.set(key, value)
        return pystray.MenuItem(text, select, radio=True,
                                checked=lambda item: settings[key] == value)

    def toggle(text, key):
        def flip(icon, item):
            settings.set(key, not settings[key])
        return pystray.MenuItem(text, flip, checked=lambda item: settings[key])

    separator = pystray.Menu.SEPARATOR
    return pystray.Menu(
        pystray.MenuItem(APP_NAME, None, enabled=False),
        separator,
        pystray.MenuItem("Now playing", pystray.Menu(
            toggle("Show album art", "album_art"),
            separator,
            choice("Time: elapsed / total", "time_display", "elapsed"),
            choice("Time: elapsed / remaining", "time_display", "remaining"),
            choice("Time: hidden", "time_display", "off"),
        )),
        pystray.MenuItem("When idle", pystray.Menu(
            choice("Clock", "idle_screen", "clock"),
            choice("System stats", "idle_screen", "stats"),
            choice("Weather", "idle_screen", "weather"),
            choice("Cycle through all", "idle_screen", "cycle"),
            separator,
            toggle("24-hour clock", "clock_24h"),
            toggle("Show date", "clock_show_date"),
            separator,
            pystray.MenuItem("Set weather city...", on_set_weather_city),
            choice("Celsius", "weather_units", "celsius"),
            choice("Fahrenheit", "weather_units", "fahrenheit"),
        )),
        toggle("Volume overlay", "volume_overlay"),
        toggle("Startup animation", "startup_animation"),
        separator,
        pystray.MenuItem("Pause display", on_toggle_pause,
                         checked=lambda item: display_paused.is_set()),
        pystray.MenuItem("Start with Windows", on_toggle_startup,
                         checked=lambda item: get_startup_entry() is not None),
        pystray.MenuItem("Open settings file", on_open_settings),
        pystray.MenuItem("Open log", on_open_log),
        separator,
        pystray.MenuItem("Quit", on_quit),
    )


# --- Main Loop ---

def oled_loop():
    """Main OLED update loop running in a background thread."""
    base_url = None
    last_url = None  # for cleanup, even if we're paused/disconnected at quit
    last_heartbeat = 0
    last_bitmap = None
    last_frame_sent = 0
    last_media_fetch = 0
    last_settings_check = 0
    cached_info = None
    startup_played = False
    frame_interval = 0.1  # 10 FPS for smooth scrolling
    media_fetch_interval = 2  # fetch media info every 2 seconds
    resend_interval = 1  # resend an unchanged frame at most this often

    # Volume tracking
    last_volume = None
    last_muted = None
    volume_display_until = 0  # timestamp when volume overlay should disappear
    volume_overlay_duration = 2  # seconds to show volume overlay
    volume_error_logged = False

    while not stop_event.is_set():
        try:
            now = time.time()

            if now - last_settings_check >= 2:
                settings.reload_if_changed()
                last_settings_check = now

            if display_paused.is_set():
                if base_url is not None:
                    stop_game(base_url)
                    log.info("Display paused")
                    base_url = None
                stop_event.wait(0.5)
                continue

            if base_url is None:
                base_url = connect()
                if base_url is None:
                    break  # quitting
                last_url = base_url
                last_bitmap = None
                last_heartbeat = 0
                # Only on app start, not when reconnecting after a GG restart
                if not startup_played:
                    startup_played = True
                    if settings["startup_animation"]:
                        play_startup_animation(base_url)

            # Check volume every frame (cheap call)
            try:
                current_volume, current_muted = get_system_volume()
                if last_volume is None:
                    last_volume = current_volume
                    last_muted = current_muted
                elif abs(current_volume - last_volume) > 0.005 or current_muted != last_muted:
                    last_volume = current_volume
                    last_muted = current_muted
                    if settings["volume_overlay"]:
                        volume_display_until = now + volume_overlay_duration
            except Exception as e:
                if not volume_error_logged:  # once, not 10 times a second
                    log.warning(f"Could not read system volume: {e}")
                    volume_error_logged = True

            if now < volume_display_until:
                # Volume overlay takes priority
                frame = render_volume(last_volume, last_muted)
            else:
                # Fetch media info periodically (not every frame)
                if now - last_media_fetch >= media_fetch_interval:
                    # Set before fetching so a failing query also waits the full interval
                    last_media_fetch = now
                    media_watcher.refresh(fetch_art=settings["album_art"])
                cached_info = media_watcher.info

                # Status 4 = Playing in the SMTC enum
                is_playing = (cached_info and cached_info["title"]
                              and cached_info["status"] == 4)

                if is_playing:
                    position = current_position(cached_info, is_playing)
                    frame = render_now_playing(cached_info, position)
                else:
                    frame = render_idle()

            # Skip identical frames (idle clock, paused scroll) to spare the
            # Engine ~10 requests/s, but resend periodically so it never goes stale
            bitmap = image_to_bitmap(frame)
            if bitmap != last_bitmap or now - last_frame_sent >= resend_interval:
                send_frame(base_url, bitmap)
                last_bitmap = bitmap
                last_frame_sent = now

            if now - last_heartbeat > 5:
                send_heartbeat(base_url)
                last_heartbeat = now

            stop_event.wait(frame_interval)

        except requests.RequestException as e:
            # Usually GG was restarted (new port) or closed: re-discover and re-register
            log.warning(f"Lost connection to SteelSeries Engine: {e}")
            base_url = None
            stop_event.wait(1)
        except Exception as e:
            log.error(f"OLED loop error: {e}", exc_info=True)
            stop_event.wait(1)

    if last_url:
        cleanup(last_url)
    log.info("OLED loop stopped")


def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    setup_logging()

    if not acquire_single_instance():
        log.info("Another instance is already running; exiting")
        show_message(f"{APP_NAME} is already running.\n\nLook for its icon in the system tray.")
        return

    log.info(f"Starting {APP_NAME}")
    settings.open(SETTINGS_PATH)
    try:
        refresh_startup_entry()
    except OSError as e:
        log.warning(f"Could not update Start with Windows entry: {e}")

    icon = pystray.Icon(APP_ID, create_tray_icon(), APP_NAME, menu=build_menu())

    # Run the OLED loop in a background thread. It waits for SteelSeries
    # Engine itself, so starting before GG (e.g. at login) is fine.
    oled_thread = threading.Thread(target=oled_loop, daemon=True)
    oled_thread.start()

    # Run tray icon on main thread (blocks until quit)
    icon.run()

    # Signal OLED thread to stop and wait
    stop_event.set()
    oled_thread.join(timeout=3)


if __name__ == "__main__":
    main()

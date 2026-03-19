#!/usr/bin/env python3
"""
extract_metadata.py — Car Mileage Photo Metadata Extractor

Reads EXIF data from photos in the photos/ directory and writes a
mileage_log.csv suitable for import into Excel or a database.

Usage:
    export ANTHROPIC_API_KEY=sk-ant-...
    python extract_metadata.py
"""

import base64
import csv
import json
import os
import sys
import time
import urllib.request
import urllib.error
from pathlib import Path

try:
    from PIL import Image
    import piexif
    import anthropic
except ImportError as e:
    print(f"Missing dependency: {e}")
    print("Run: pip install -r requirements.txt")
    sys.exit(1)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

BASE_DIR = Path(__file__).parent
PHOTOS_DIR = BASE_DIR / "photos"
OUTPUT_CSV = BASE_DIR / "mileage_log.csv"

SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".heic"}

NOMINATIM_URL = "https://nominatim.openstreetmap.org/reverse?format=json&lat={lat}&lon={lon}"
NOMINATIM_USER_AGENT = "millagelog/1.0"

CLAUDE_MODEL = "claude-haiku-4-5-20251001"

CSV_HEADERS = [
    "Filename",
    "Date",
    "Time",
    "Latitude",
    "Longitude",
    "Location",
    "Mileage_Reading",
    "File_Size_MB",
    "Camera_Make",
    "Camera_Model",
    "Notes",
]

ODOMETER_SYSTEM_PROMPT = (
    "You are an odometer reading assistant. Your only job is to read the numeric "
    "mileage value shown on the odometer or trip meter display in the provided photo.\n\n"
    "Rules:\n"
    "- Return ONLY the numeric value (e.g. '124853' or '124853.2'). No units, no labels.\n"
    "- If the odometer is not visible, not in focus, or the number cannot be determined "
    "with confidence, return exactly: NONE\n"
    "- Do not return any other text."
)

MEDIA_TYPES = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".heic": "image/webp",  # Claude does not natively support HEIC; fallback
}

# ---------------------------------------------------------------------------
# Validation & setup
# ---------------------------------------------------------------------------


def validate_api_key() -> str:
    key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not key:
        print("Error: ANTHROPIC_API_KEY environment variable is not set.")
        print("Get your key at: https://console.anthropic.com")
        print("Then run:  export ANTHROPIC_API_KEY=sk-ant-...")
        sys.exit(1)
    return key


def ensure_photos_dir() -> None:
    PHOTOS_DIR.mkdir(exist_ok=True)


def collect_image_paths() -> list[Path]:
    paths = sorted(
        p for p in PHOTOS_DIR.iterdir()
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS
    )
    return paths


# ---------------------------------------------------------------------------
# EXIF helpers
# ---------------------------------------------------------------------------


def extract_exif(path: Path) -> dict:
    """Open image with Pillow and load raw piexif dict. Returns {} on any failure."""
    try:
        img = Image.open(path)
        raw_exif = img.info.get("exif", b"")
        if not raw_exif:
            return {}
        return piexif.load(raw_exif)
    except Exception:
        return {}


def parse_datetime(exif: dict) -> tuple[str, str]:
    """Extract date and time from EXIF. Returns ('YYYY-MM-DD', 'HH:MM:SS') or ('N/A', 'N/A')."""
    try:
        exif_ifd = exif.get("Exif", {})
        raw = exif_ifd.get(piexif.ExifIFD.DateTimeOriginal, b"")
        if not raw:
            # Fall back to DateTime in the 0th IFD
            raw = exif.get("0th", {}).get(piexif.ImageIFD.DateTime, b"")
        if not raw:
            return "N/A", "N/A"
        dt_str = raw.decode("utf-8", errors="replace").strip("\x00")
        date_part, time_part = dt_str.split(" ", 1)
        date_formatted = date_part.replace(":", "-")
        return date_formatted, time_part
    except Exception:
        return "N/A", "N/A"


def dms_to_decimal(dms: tuple, ref: bytes) -> float:
    """Convert DMS rational tuple from piexif to decimal degrees."""
    def rational(r):
        num, den = r
        if den == 0:
            return 0.0
        return num / den

    degrees = rational(dms[0])
    minutes = rational(dms[1])
    seconds = rational(dms[2])
    decimal = degrees + minutes / 60.0 + seconds / 3600.0
    if ref in (b"S", b"W"):
        decimal = -decimal
    return round(decimal, 6)


def parse_gps(exif: dict) -> tuple[float | None, float | None]:
    """Extract GPS lat/lon as decimal degrees. Returns (None, None) if unavailable."""
    try:
        gps = exif.get("GPS", {})
        if not gps:
            return None, None
        lat_dms = gps.get(piexif.GPSIFD.GPSLatitude)
        lat_ref = gps.get(piexif.GPSIFD.GPSLatitudeRef, b"N")
        lon_dms = gps.get(piexif.GPSIFD.GPSLongitude)
        lon_ref = gps.get(piexif.GPSIFD.GPSLongitudeRef, b"E")
        if not lat_dms or not lon_dms:
            return None, None
        lat = dms_to_decimal(lat_dms, lat_ref)
        lon = dms_to_decimal(lon_dms, lon_ref)
        return lat, lon
    except Exception:
        return None, None


def parse_camera(exif: dict) -> tuple[str, str]:
    """Extract camera make and model from EXIF 0th IFD."""
    try:
        ifd0 = exif.get("0th", {})
        make = ifd0.get(piexif.ImageIFD.Make, b"")
        model = ifd0.get(piexif.ImageIFD.Model, b"")
        make_str = make.decode("utf-8", errors="replace").strip("\x00").strip()
        model_str = model.decode("utf-8", errors="replace").strip("\x00").strip()
        return make_str or "N/A", model_str or "N/A"
    except Exception:
        return "N/A", "N/A"


# ---------------------------------------------------------------------------
# File size
# ---------------------------------------------------------------------------


def get_file_size(path: Path) -> str:
    try:
        size_mb = path.stat().st_size / (1024 * 1024)
        return str(round(size_mb, 3))
    except Exception:
        return "N/A"


# ---------------------------------------------------------------------------
# Reverse geocoding
# ---------------------------------------------------------------------------


def reverse_geocode(lat: float, lon: float) -> str:
    """Look up a human-readable address from coordinates via Nominatim."""
    time.sleep(1)  # Respect Nominatim rate limit: 1 req/sec
    try:
        url = NOMINATIM_URL.format(lat=lat, lon=lon)
        req = urllib.request.Request(url, headers={"User-Agent": NOMINATIM_USER_AGENT})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return data.get("display_name", "N/A")
    except Exception as exc:
        return f"Geocoding failed: {exc}"


# ---------------------------------------------------------------------------
# Odometer OCR via Claude vision
# ---------------------------------------------------------------------------


def read_odometer(path: Path, client: anthropic.Anthropic) -> str:
    """Use Claude vision to read odometer value from photo. Returns numeric string or 'N/A'."""
    try:
        suffix = path.suffix.lower()
        media_type = MEDIA_TYPES.get(suffix, "image/jpeg")
        image_data = base64.standard_b64encode(path.read_bytes()).decode("utf-8")

        response = client.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=50,
            system=ODOMETER_SYSTEM_PROMPT,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": image_data,
                            },
                        },
                        {
                            "type": "text",
                            "text": "What is the odometer reading in this photo?",
                        },
                    ],
                }
            ],
        )
        reading = response.content[0].text.strip()
        if reading.upper() == "NONE" or not reading:
            return "N/A"
        return reading
    except anthropic.APIError as exc:
        return f"OCR_ERROR: {exc}"
    except Exception as exc:
        return f"OCR_ERROR: {exc}"


# ---------------------------------------------------------------------------
# Core processing
# ---------------------------------------------------------------------------


def process_image(path: Path, client: anthropic.Anthropic) -> dict:
    """Extract all metadata from a single image. Never raises — errors go into Notes."""
    notes: list[str] = []

    # File size (always available if the file exists)
    file_size = get_file_size(path)

    # EXIF extraction
    try:
        exif = extract_exif(path)
    except Exception as exc:
        exif = {}
        notes.append(f"Cannot open image: {exc}")

    if not exif:
        notes.append("No EXIF data")

    # Date / time
    date_str, time_str = parse_datetime(exif)

    # GPS
    lat, lon = parse_gps(exif)
    lat_str = str(lat) if lat is not None else "N/A"
    lon_str = str(lon) if lon is not None else "N/A"

    # Location
    if lat is not None and lon is not None:
        location = reverse_geocode(lat, lon)
    else:
        location = "N/A"
        if not exif:
            pass  # already noted no EXIF
        else:
            notes.append("No GPS data in EXIF")

    # Odometer
    mileage = read_odometer(path, client)

    # Camera
    make, model = parse_camera(exif)

    return {
        "Filename": path.name,
        "Date": date_str,
        "Time": time_str,
        "Latitude": lat_str,
        "Longitude": lon_str,
        "Location": location,
        "Mileage_Reading": mileage,
        "File_Size_MB": file_size,
        "Camera_Make": make,
        "Camera_Model": model,
        "Notes": "; ".join(notes),
    }


# ---------------------------------------------------------------------------
# CSV output
# ---------------------------------------------------------------------------


def write_csv(rows: list[dict]) -> None:
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_HEADERS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    print(f"\nWrote {len(rows)} record(s) to {OUTPUT_CSV}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    api_key = validate_api_key()
    client = anthropic.Anthropic(api_key=api_key)

    ensure_photos_dir()
    paths = collect_image_paths()

    if not paths:
        print(f"No images found in {PHOTOS_DIR}")
        print("Drop your .jpg / .jpeg / .png / .heic photos into the photos/ folder and re-run.")
        sys.exit(0)

    print(f"Processing {len(paths)} image(s) from {PHOTOS_DIR} ...")
    rows: list[dict] = []
    for i, path in enumerate(paths, 1):
        print(f"  [{i}/{len(paths)}] {path.name}")
        record = process_image(path, client)
        rows.append(record)

    write_csv(rows)
    print("Done.")


if __name__ == "__main__":
    main()

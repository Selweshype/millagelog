# millagelog

Automatically extract metadata from car mileage photos and produce a CSV log for tax compliance tracking (500 km private use limit).

## What it does

For each photo in the `photos/` folder the script extracts:

| Field | Source |
|---|---|
| Date / Time | EXIF data |
| GPS coordinates | EXIF GPS tags |
| Location | Reverse geocoded via OpenStreetMap Nominatim |
| Mileage reading | Claude vision AI reads the odometer display |
| File size | File system |
| Camera make / model | EXIF data |

Output is written to `mileage_log.csv`, ready to import into Excel or a database.

## Requirements

- Python 3.11+
- An [Anthropic API key](https://console.anthropic.com) (for odometer OCR)

## Setup

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Set your API key
export ANTHROPIC_API_KEY=sk-ant-...
```

## Usage

1. Drop your `.jpg`, `.jpeg`, `.png`, or `.heic` photos into the `photos/` folder
2. Run the script:

```bash
python extract_metadata.py
```

3. Open `mileage_log.csv` in Excel or import it into your database

## CSV columns

```
Filename, Date, Time, Latitude, Longitude, Location,
Mileage_Reading, File_Size_MB, Camera_Make, Camera_Model, Notes
```

## Notes

- Missing fields are marked `N/A`
- GPS location lookup respects Nominatim's 1 request/second rate limit
- Odometer reading uses `claude-haiku-4-5` (cheapest Claude vision model)
- The `photos/` folder and `mileage_log.csv` are excluded from git

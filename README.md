# ChronoFix

Fix incorrect or missing date metadata on your photos, RAW files, and videos. Works interactively, doesn't re-encode anything, and won't change a single file without showing you what it's doing.

---

## Table of Contents

- [Why this exists](#why-this-exists)
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Supported File Types](#supported-file-types)
- [How the Date Gets Chosen](#how-the-date-gets-chosen)
- [Filename Patterns Recognised](#filename-patterns-recognised)
- [What Gets Written to Files](#what-gets-written-to-files)
- [Configuration](#configuration)
- [Report File](#report-file)
- [Safety](#safety)
- [Roadmap](#roadmap)
- [Contributing](#contributing)
- [License](#license)

---

## Why this exists

When you copy photos off a phone, restore from a backup, or receive files over WhatsApp, the "Date Taken" or creation date often ends up wrong or completely missing. ChronoFix goes through every media file in a folder (including subfolders), figures out the correct date from whatever sources it can find, and writes it back into the file's metadata and filesystem timestamps. You get to confirm or choose the date for each file before anything is written.

---

## Features

- Supports images, RAW files, and videos (JPG, PNG, HEIC, WEBP, CR2, NEF, ARW, MP4, MOV, MKV, and more)
- Reads dates from four places: EXIF/video metadata, folder name, filename, and filesystem Date Modified
- Auto-selects the date when all sources agree; asks you when they conflict
- Nothing is changed silently; every decision is printed to the screen
- Videos are never re-encoded. All streams are copied losslessly with `-c copy` and `-map_metadata 0`; only the date tags change
- The original video is only replaced after the new file has been fully and successfully written
- On Windows, reads and writes the "Date Taken" property via the `propsys` API, which works correctly regardless of your Windows display language
- Detects when a file's extension doesn't match its actual content (like a HEIC file named `.jpg`) and offers to rename it
- Recognises WhatsApp filename patterns and flags those dates as unreliable
- Saves a detailed report after each run. If a report from a previous run already exists, it gets renamed with a timestamp so you don't lose it

---

## Requirements

**Python 3.13 or higher**

You also need these tools installed and available on your PATH:

| Tool | What it's used for | How to install |
|---|---|---|
| [exiftool](https://exiftool.org/) | Read/write EXIF metadata for images and RAW files | `brew install exiftool` on Mac, `sudo apt install libimage-exiftool-perl` on Linux, or download the .exe on Windows |
| [ffmpeg](https://ffmpeg.org/) | Write date metadata into video files | `brew install ffmpeg` on Mac, `sudo apt install ffmpeg` on Linux, or download from ffmpeg.org |
| [ffprobe](https://ffmpeg.org/) | Read date metadata from video files | Comes bundled with ffmpeg |
| [pywin32](https://github.com/mhammond/pywin32) | Read/write the Windows "Date Taken" property (Windows only) | `pip install pywin32` (build 306 or higher required for Python 3.13) |

---

## Installation

```bash
git clone https://github.com/bhaumik-solanki/chronofix.git
cd chronofix
```

No additional Python packages to install unless you're on Windows (see above).

---

## Usage

```bash
python chronofix.py "path/to/your/folder"
```

Examples:

```bash
# Windows
python chronofix.py "C:\Users\YourName\Pictures\Vacation 2024"

# Mac or Linux
python chronofix.py "/home/yourname/Photos/12 September 2023"
```

The script processes all media files recursively. For each file it will:

1. Show you every date it found and where it came from
2. Either auto-select the date (when all sources agree) or ask you to pick one
3. Write the chosen date into the file's metadata and filesystem timestamps

---

## Supported File Types

| Category | Extensions |
|---|---|
| Images | `.jpg` `.jpeg` `.png` `.gif` `.bmp` `.tiff` `.tif` `.heic` `.heif` `.webp` |
| RAW | `.cr2` `.nef` `.arw` `.dng` `.raw` `.orf` `.sr2` `.pef` `.raf` `.rw2` `.rwl` `.srw` `.x3f` |
| Video | `.mp4` `.mov` `.avi` `.mkv` `.wmv` `.flv` `.webm` `.m4v` `.mpg` `.mpeg` `.3gp` `.mts` `.m2ts` |

---

## How the Date Gets Chosen

The script checks four sources for every file:

1. **File metadata** (highest priority): EXIF tags for images and RAW (`DateTimeOriginal`, `DateTime`, `DateTimeDigitized`), and `creation_time` / `date` / `date_recorded` for video
2. **Folder name**: a folder named `12 September 2023` is treated as the date `2023-09-12`
3. **Filename**: patterns like `IMG_20230912_143022`, `Screenshot_2023-09-12-14-30-22`, WhatsApp names, etc.
4. **Filesystem Date Modified**: always shown as an option but never auto-selected

Here's what happens in each scenario:

| Situation | What happens |
|---|---|
| All sources agree on date, hour, and minute | Auto-selected, no prompt |
| Folder and metadata agree | Uses metadata (it has the precise time) |
| Folder and metadata disagree | Asks you to choose |
| Metadata and filename, same time | Uses metadata |
| Metadata and filename, different time | Asks you to choose |
| Only one source found | Asks you to confirm |
| Nothing found anywhere | Asks you to type the date manually |

When all three main sources agree on date, hour, and minute, the seconds are resolved automatically:

| Metadata seconds | Filename seconds | What gets used |
|---|---|---|
| `00` | `00` | Metadata |
| `00` | nonzero | Filename (it has real seconds, metadata doesn't) |
| nonzero | `00` | Metadata (it has real seconds, filename doesn't) |
| same value | same value | Metadata |
| different values | different values | Metadata |

---

## Filename Patterns Recognised

| Pattern | Example |
|---|---|
| `IMG_YYYYMMDD_HHMMSS` | `IMG_20230912_143022.jpg` |
| `IMG_YYYYMMDD_HHMMSSMMM` | `IMG_20260515_182551618.jpg` |
| `VID_YYYYMMDD_HHMMSS` | `VID_20230912_143022.mp4` |
| `Screenshot_YYYY-MM-DD-HH-MM-SS` | `Screenshot_2023-09-12-14-30-22.png` |
| `Screenshot_YYYYMMDD_HHMMSS` | `Screenshot_20230912_143022.png` |
| `YYYYMMDD_HHMMSS` | `20230912_143022.jpg` |
| `YYYYMMDDHHMMSS` | `20230912143022.jpg` |
| `YYYY-MM-DD_HH-MM-SS` | `2023-09-12_14-30-22.jpg` |
| `YYYY-MM-DD` | `2023-09-12.jpg` |
| `YYYYMMDD` | `20230912.jpg` |
| `IMG-YYYYMMDD-WA` | `IMG-20230912-WA0016.jpg` (WhatsApp) |
| WhatsApp full format | `WhatsApp Image 2023-04-09 at 6.26.53 PM.jpeg` |

---

## What Gets Written to Files

| File type | What gets updated |
|---|---|
| Images | `DateTimeOriginal`, `CreateDate`, `ModifyDate` via exiftool |
| RAW | `DateTimeOriginal`, `CreateDate` via exiftool |
| Video | `creation_time`, `date`, `date_recorded` via ffmpeg (stored as UTC, which is standard for video containers) |
| All files on Windows | The "Date Taken" Windows property via the propsys API |
| All files | Filesystem `Date Modified` via `os.utime` |

---

## Configuration

Two settings near the top of `chronofix.py` you might want to change:

```python
# The earliest date the script will accept without asking you to confirm
MIN_DATE = datetime(2022, 1, 1)

# Set to True if you want metadata rewritten even when it already matches the chosen date
FORCE_METADATA_UPDATE = False
```

---

## Report File

After each run, a report is written to the folder you processed:

```
media_date_update_report.txt
```

If a report from a previous run is already there, it gets renamed before being overwritten:

```
media_date_update_report_20260519_143022.txt   (previous run, kept)
media_date_update_report.txt                   (current run)
```

The report covers which dependencies were available, what extraction priority was used, the decision rules that applied, what properties were updated, and a full count of files processed, updated, skipped, and errored.

---

## Safety

Images are never re-encoded. Only metadata tags are modified in place by exiftool.

Videos are never re-encoded. ffmpeg copies every stream losslessly and preserves all existing metadata tags; it only overrides the date-related ones. The original file is only replaced after the new file has been fully written.

Every auto-selection is printed to the screen with a reason, so you always know what was decided and why. Nothing changes silently.

---

## Roadmap

- **Dry-run mode (`--dry-run`)**: show what would be changed without writing anything. Good for checking a large folder before committing.
- **Batch mode (`--auto`)**: skip all prompts and auto-select the best date for every file. Useful when you're confident in the sources and don't want to answer for each file individually.
- **CLI date range flags (`--min-date`, `--max-date`)**: pass the valid date range as arguments instead of editing the script.
- **Google Photos JSON sidecar support**: when you export from Google Takeout, each file comes with a `.json` sidecar containing the original date. Parsing those would give very accurate results for Takeout exports.
- **Progress bar**: for large folders in batch mode, a progress bar via `tqdm` would be much nicer than thousands of lines scrolling past.
- **Config file support**: store settings like `MIN_DATE` and `FORCE_METADATA_UPDATE` in a `config.toml` instead of editing the script directly.
- **Undo support**: save a `before_state.json` before making any changes, so a future `--restore` run can put everything back.

---

## Contributing

Pull requests are welcome. For anything large, please open an issue first so we can talk through the approach before you put in the work.

---

## License

MIT. See [LICENSE](LICENSE) for details.

---

*Written for anyone who has had hundreds of photos show up with the wrong date after restoring a backup, or whose WhatsApp downloads all show today's date instead of when the photo was actually taken.*

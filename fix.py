import os
import sys
from datetime import datetime
from pathlib import Path
import re
import subprocess
import json

# Check Python version — requires 3.13+
if sys.version_info < (3, 13):
    print("This script requires Python 3.13 or higher")
    sys.exit(1)

# Optional imports for Windows
# Note: pywin32 build 306 or higher is required for Python 3.13 compatibility
if sys.platform == 'win32':
    try:
        import win32com.client
        from win32com.propsys import propsys, pscon
        from win32com.shell import shellcon
        import pythoncom
        from datetime import timezone
        has_win32 = True
    except ImportError:
        has_win32 = False
else:
    has_win32 = False

# Global settings
MIN_DATE     = datetime(2022, 1, 1)
MAX_DATE     = datetime.now().replace(hour=23, minute=59, second=59)
CURRENT_DATE = datetime.now()

# Set to True to always update metadata even when it already exists and matches
FORCE_METADATA_UPDATE = False

# Dependency availability — set once at startup in main(), referenced throughout
EXIFTOOL_AVAILABLE = False
FFMPEG_AVAILABLE   = False
FFPROBE_AVAILABLE  = False

# File extension categories
IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif',
                    '.heic', '.heif', '.webp'}
RAW_EXTENSIONS   = {'.cr2', '.nef', '.arw', '.dng', '.raw', '.orf', '.sr2', '.pef',
                    '.raf', '.rw2', '.rwl', '.srw', '.x3f'}
VIDEO_EXTENSIONS = {'.mp4', '.mov', '.avi', '.mkv', '.wmv', '.flv', '.webm',
                    '.m4v', '.mpg', '.mpeg', '.3gp', '.mts', '.m2ts'}

print(f"Current system date/time: {CURRENT_DATE}")
print(f"Valid date range: {MIN_DATE.date()} to {MAX_DATE.date()}\n")


# =============================================================================
# Utility & classification
# =============================================================================

def is_date_valid(date: datetime) -> bool:
    return MIN_DATE <= date <= MAX_DATE


def is_whatsapp_file(filepath: Path) -> bool:
    filename = filepath.name.lower()
    return 'whatsapp' in filename or '-wa' in filename or '_wa' in filename


def is_raw_file(filepath: Path) -> bool:
    return filepath.suffix.lower() in RAW_EXTENSIONS


def is_image_file(filepath: Path) -> bool:
    return filepath.suffix.lower() in IMAGE_EXTENSIONS


def is_video_file(filepath: Path) -> bool:
    return filepath.suffix.lower() in VIDEO_EXTENSIONS


def get_file_type_name(filepath: Path) -> str:
    if is_raw_file(filepath):
        return "RAW"
    elif is_image_file(filepath):
        return "Image"
    elif is_video_file(filepath):
        return "Video"
    return "Unknown"


def get_file_date_modified(filepath: Path) -> datetime | None:
    try:
        mtime = filepath.stat().st_mtime
        return datetime.fromtimestamp(mtime)
    except Exception as e:
        print(f"  ! Error getting Date Modified: {e}")
        return None


# =============================================================================
# File-type mismatch detection
# =============================================================================

def detect_actual_file_type(filepath: Path) -> str | None:
    """Read first 32 bytes and return correct extension, or None if unknown."""
    try:
        with open(filepath, 'rb') as f:
            header = f.read(32)

        if header[:3] == b'\xff\xd8\xff':
            return '.jpg'
        if header[:8] == b'\x89PNG\r\n\x1a\n':
            return '.png'
        if b'ftyp' in header[:12]:
            if b'heic' in header or b'heix' in header or b'mif1' in header or b'hevc' in header:
                return '.heic'
            if b'avif' in header:
                return '.avif'
            if b'mp4' in header or b'isom' in header or b'M4V' in header:
                return '.mp4'
            if b'qt' in header:
                return '.mov'
        if header[:6] in (b'GIF87a', b'GIF89a'):
            return '.gif'
        if header[:2] == b'BM':
            return '.bmp'
        if header[:4] in (b'II\x2a\x00', b'MM\x00\x2a'):
            if b'CR' in header:
                return '.cr2'
            return '.tiff'
        if header[:4] == b'RIFF' and header[8:12] == b'WEBP':
            return '.webp'
        return None

    except Exception as e:
        print(f"    ! Error detecting file type: {e}")
        return None


def fix_file_extension_mismatch(filepath: Path) -> Path | None:
    """
    Check if extension matches actual file type.
    Returns new path if renamed, original path if no change, None if skipped.
    """
    actual_type   = detect_actual_file_type(filepath)
    current_ext   = filepath.suffix.lower()

    if actual_type is None:
        return filepath

    ext_mapping        = {'.jpeg': '.jpg', '.tif': '.tiff'}
    normalized_current = ext_mapping.get(current_ext, current_ext)
    normalized_actual  = ext_mapping.get(actual_type, actual_type)

    if normalized_current == normalized_actual:
        return filepath

    print(f"\n  WARNING: FILE TYPE MISMATCH DETECTED!")
    print(f"     Current extension: {current_ext}")
    print(f"     Actual file type:  {actual_type}")
    print(f"     Options:")
    print(f"     1. Rename file to correct extension ({filepath.stem}{actual_type})")
    print(f"     2. Keep current name and try to process anyway")
    print(f"     3. Skip this file")

    while True:
        choice = input("     Choose (1-3): ").strip()
        if choice == '1':
            new_path = filepath.parent / f"{filepath.stem}{actual_type}"
            if new_path.exists():
                print(f"     Cannot rename: {new_path.name} already exists!")
                continue
            try:
                filepath.rename(new_path)
                print(f"     Renamed to: {new_path.name}")
                return new_path
            except Exception as e:
                print(f"     Error renaming file: {e}")
                continue
        elif choice == '2':
            return filepath
        elif choice == '3':
            return None
        else:
            print("     Invalid choice. Please enter 1, 2, or 3.")


# =============================================================================
# Dependency checks  (called once in main(); results stored in module globals)
# =============================================================================

def check_exiftool_available() -> bool:
    try:
        result = subprocess.run(['exiftool', '-ver'], capture_output=True, text=True, timeout=10)
        return result.returncode == 0
    except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired):
        return False


def check_ffmpeg_available() -> bool:
    try:
        subprocess.run(['ffmpeg', '-version'], capture_output=True, check=True, timeout=10)
        return True
    except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired):
        return False


def check_ffprobe_available() -> bool:
    try:
        subprocess.run(['ffprobe', '-version'], capture_output=True, check=True, timeout=10)
        return True
    except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired):
        return False


# =============================================================================
# Date extraction
# =============================================================================

def extract_date_from_folder(folder_path: Path) -> datetime | None:
    """Extract date from folder name in 'DD Month YYYY' format."""
    folder_name = folder_path.name
    months = {
        'january': 1, 'february': 2, 'march': 3, 'april': 4,
        'may': 5, 'june': 6, 'july': 7, 'august': 8,
        'september': 9, 'october': 10, 'november': 11, 'december': 12
    }
    match = re.search(r'(\d{1,2})\s+([a-zA-Z]+)\s+(\d{4})', folder_name)
    if match:
        day        = int(match.group(1))
        month_name = match.group(2).lower()
        year       = int(match.group(3))
        if month_name in months:
            try:
                return datetime(year, months[month_name], day, 12, 0, 0)
            except ValueError:
                pass
    return None


def extract_date_from_filename(
    filepath: Path,
    include_whatsapp: bool = False
) -> list[tuple[datetime, str, str, bool]]:
    """
    Extract date from filename using regex patterns.
    Returns list of (datetime, format_name, matched_text, has_time).
    Patterns: time-bearing first (most to least specific), then date-only.
    """
    if not include_whatsapp and is_whatsapp_file(filepath):
        return []

    filename          = filepath.stem
    found_dates: list[tuple[datetime, str, str, bool]] = []
    matched_positions: set[int] = set()

    patterns = [
        # ── WITH TIME ─────────────────────────────────────────────────────────

        # WhatsApp Image 2023-04-09 at 6.26.53 PM.jpeg
        (r'WhatsApp\s+(?:Image|Video)\s+(\d{4})-(\d{2})-(\d{2})\s+at\s+(\d{1,2})\.(\d{2})\.(\d{2})\s+(AM|PM)',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                            (int(m.group(4)) % 12) + (12 if m.group(7).upper() == 'PM' else 0),
                            int(m.group(5)), int(m.group(6))),
         "WhatsApp YYYY-MM-DD at H.MM.SS AM/PM", True),

        # IMG_20260515_182551618  (milliseconds appended directly, no separator)
        (r'IMG[-_](\d{4})(\d{2})(\d{2})[-_](\d{2})(\d{2})(\d{2})\d{1,3}(?!\d)',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                            int(m.group(4)), int(m.group(5)), int(m.group(6))),
         "IMG_YYYYMMDD_HHMMSSMMM", True),

        # IMG_20250730_002711 or IMG_20250730_002711_483
        (r'IMG[-_](\d{4})(\d{2})(\d{2})[-_](\d{2})(\d{2})(\d{2})',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                            int(m.group(4)), int(m.group(5)), int(m.group(6))),
         "IMG_YYYYMMDD_HHMMSS", True),

        # VID_20240329_143022
        (r'VID[-_](\d{4})(\d{2})(\d{2})[-_](\d{2})(\d{2})(\d{2})',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                            int(m.group(4)), int(m.group(5)), int(m.group(6))),
         "VID_YYYYMMDD_HHMMSS", True),

        # IMG20230902153011
        (r'IMG(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                            int(m.group(4)), int(m.group(5)), int(m.group(6))),
         "IMGYYYYMMDDHHMMSS", True),

        # VID20230902153011
        (r'VID(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                            int(m.group(4)), int(m.group(5)), int(m.group(6))),
         "VIDYYYYMMDDHHMMSS", True),

        # Screenshot_2025-03-12-11-36-14-787
        (r'[Ss]creenshot[_-](\d{4})[-_](\d{2})[-_](\d{2})[-_](\d{2})[-_](\d{2})[-_](\d{2})',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                            int(m.group(4)), int(m.group(5)), int(m.group(6))),
         "Screenshot_YYYY-MM-DD-HH-MM-SS", True),

        # Screenshot_20250422_214001
        (r'[Ss]creenshot[_-](\d{4})(\d{2})(\d{2})[_-](\d{2})(\d{2})(\d{2})',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                            int(m.group(4)), int(m.group(5)), int(m.group(6))),
         "Screenshot_YYYYMMDD_HHMMSS", True),

        # 20230902_160253
        (r'(\d{4})(\d{2})(\d{2})[-_](\d{2})(\d{2})(\d{2})',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                            int(m.group(4)), int(m.group(5)), int(m.group(6))),
         "YYYYMMDD_HHMMSS", True),

        # 20230815143022 (14-digit block)
        (r'(?<![A-Za-z])(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})(?!\d)',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                            int(m.group(4)), int(m.group(5)), int(m.group(6))),
         "YYYYMMDDHHMMSS", True),

        # 2023-08-15_14-30-22
        (r'(\d{4})[-_](\d{2})[-_](\d{2})[-_](\d{2})[-_](\d{2})[-_](\d{2})',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                            int(m.group(4)), int(m.group(5)), int(m.group(6))),
         "YYYY-MM-DD-HH-MM-SS", True),

        # ── DATE ONLY ─────────────────────────────────────────────────────────

        # IMG-20230409-WA0016 (WhatsApp)
        (r'IMG-(\d{4})(\d{2})(\d{2})-WA\d+',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), 12, 0, 0),
         "IMG-YYYYMMDD-WA (WhatsApp, no time)", False),

        # VID-20230409-WA0010 (WhatsApp)
        (r'VID-(\d{4})(\d{2})(\d{2})-WA\d+',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), 12, 0, 0),
         "VID-YYYYMMDD-WA (WhatsApp, no time)", False),

        # IMG-20250522 / IMG_20250522 (date only, no time block following)
        (r'IMG[-_](\d{4})(\d{2})(\d{2})(?![-_]?\d{2}[-_]?\d{2}[-_]?\d{2})',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), 12, 0, 0),
         "IMG_YYYYMMDD (no time)", False),

        # VID-20240329 (date only)
        (r'VID[-_](\d{4})(\d{2})(\d{2})(?![-_]?\d{2}[-_]?\d{2}[-_]?\d{2})',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), 12, 0, 0),
         "VID_YYYYMMDD (no time)", False),

        # 2023-08-15 / 2023_08_15
        (r'(\d{4})[-_](\d{2})[-_](\d{2})(?![-_]\d)',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), 12, 0, 0),
         "YYYY-MM-DD (no time)", False),

        # 20230815 (8-digit block)
        (r'(?<![A-Za-z\d])(\d{4})(\d{2})(\d{2})(?!\d)',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), 12, 0, 0),
         "YYYYMMDD (no time)", False),
    ]

    for pattern, date_func, format_name, has_time in patterns:
        for match in re.finditer(pattern, filename, re.IGNORECASE):
            match_range = set(range(match.start(), match.end()))
            if match_range & matched_positions:
                continue
            try:
                date = date_func(match)
                if 1 <= date.month <= 12 and 1 <= date.day <= 31:
                    if has_time:
                        if 0 <= date.hour <= 23 and 0 <= date.minute <= 59 and 0 <= date.second <= 59:
                            found_dates.append((date, format_name, match.group(0), has_time))
                            matched_positions.update(match_range)
                    else:
                        found_dates.append((date, format_name, match.group(0), has_time))
                        matched_positions.update(match_range)
            except (ValueError, OSError):
                continue

    # Deduplicate: prefer time-bearing variants
    unique_dates: list[tuple[datetime, str, str, bool]] = []
    seen_dates: set[str] = set()
    for date_info in sorted(found_dates, key=lambda x: (not x[3], x[0])):
        date_key = date_info[0].strftime("%Y-%m-%d")
        if date_key in seen_dates:
            continue
        seen_dates.add(date_key)
        unique_dates.append(date_info)

    return unique_dates


# =============================================================================
# Date parsing
# =============================================================================

def parse_datetime_with_timezone(date_str: str) -> datetime | None:
    """
    Parse a datetime string with full timezone handling.
    UTC timestamps and offset-bearing timestamps are converted to local time.
    Always returns a timezone-naive datetime in local time, or None.
    """
    if not date_str or not date_str.strip():
        return None

    date_str = date_str.strip()

    try:
        from datetime import timezone as tz, timedelta

        is_utc = (
            date_str.endswith('Z') or
            '+00:00' in date_str or
            date_str.endswith('+0000') or
            '-00:00' in date_str or
            date_str.endswith('-0000')
        )

        has_tz_offset     = False
        tz_offset_hours   = 0
        tz_offset_minutes = 0
        clean_str         = date_str.replace('Z', '')

        if '.' in clean_str:
            clean_str = clean_str.split('.')[0]

        tz_match = re.search(r'([+-])(\d{2}):?(\d{2})$', clean_str)
        if tz_match:
            has_tz_offset     = True
            sign              = 1 if tz_match.group(1) == '+' else -1
            tz_offset_hours   = int(tz_match.group(2)) * sign
            tz_offset_minutes = int(tz_match.group(3)) * sign
            clean_str         = clean_str[:tz_match.start()]

        parsed = None

        try:
            parsed = datetime.fromisoformat(clean_str)
        except ValueError:
            pass

        if not parsed:
            try:
                parsed = datetime.strptime(clean_str, '%Y:%m:%d %H:%M:%S')
            except ValueError:
                pass

        if not parsed:
            for fmt in ('%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S',
                        '%Y-%m-%dT%H:%M:%S', '%Y:%m:%d %H:%M',
                        '%Y-%m-%d %H:%M', '%Y:%m:%d', '%Y-%m-%d'):
                try:
                    parsed = datetime.strptime(clean_str, fmt)
                    break
                except ValueError:
                    continue

        if not parsed:
            return None

        if is_utc:
            utc_dt   = parsed.replace(tzinfo=tz.utc)
            local_dt = utc_dt.astimezone()
            print(f"    Converted UTC {parsed.strftime('%Y-%m-%d %H:%M:%S')} "
                  f"to local {local_dt.strftime('%Y-%m-%d %H:%M:%S')}")
            return local_dt.replace(tzinfo=None)

        elif has_tz_offset:
            offset    = timedelta(hours=tz_offset_hours, minutes=tz_offset_minutes)
            source_tz = tz(offset)
            source_dt = parsed.replace(tzinfo=source_tz)
            local_dt  = source_dt.astimezone()
            if abs(offset.total_seconds()) > 0:
                print(f"    Converted {parsed.strftime('%Y-%m-%d %H:%M:%S')} "
                      f"(UTC{tz_match.group(0)}) "
                      f"to local {local_dt.strftime('%Y-%m-%d %H:%M:%S')}")
            return local_dt.replace(tzinfo=None)

        else:
            return parsed

    except Exception as e:
        print(f"    ! Error parsing datetime '{date_str}': {e}")
        return None


def parse_exif_date(date_str: str) -> datetime | None:
    """Parse EXIF date string; delegates to timezone-aware parser when needed."""
    if not date_str or not date_str.strip():
        return None

    date_str = date_str.strip()

    has_tz = ('Z' in date_str or
              (len(date_str) > 10 and '+' in date_str[10:]) or
              (date_str.count('-') > 2 and len(date_str) > 19))
    if has_tz:
        result = parse_datetime_with_timezone(date_str)
        if result:
            return result

    try:
        return datetime.strptime(date_str, '%Y:%m:%d %H:%M:%S')
    except ValueError:
        pass

    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S',
                '%Y-%m-%dT%H:%M:%S', '%Y:%m:%d %H:%M',
                '%Y-%m-%d %H:%M', '%Y:%m:%d', '%Y-%m-%d'):
        try:
            parsed = datetime.strptime(date_str, fmt)
            if parsed.hour == 0 and parsed.minute == 0 and parsed.second == 0:
                if '%H' not in fmt and '%I' not in fmt:
                    parsed = parsed.replace(hour=12)
            return parsed
        except ValueError:
            continue

    return parse_datetime_with_timezone(date_str)


# =============================================================================
# Windows property system  (read & write via propsys API — locale-independent)
# =============================================================================

def get_windows_date_taken(filepath: Path) -> datetime | None:
    """
    Read System.Photo.DateTaken via propsys API.
    Returns timezone-naive datetime in local time, or None.
    Uses raw PROPVARIANT — no string parsing, no locale dependency.
    """
    if not has_win32:
        return None
    try:
        pythoncom.CoInitialize()
        ps    = propsys.SHGetPropertyStoreFromParsingName(
                    str(filepath), None,
                    shellcon.GPS_BESTEFFORT,
                    propsys.IID_IPropertyStore)
        pkey  = propsys.PSGetPropertyKeyFromName("System.Photo.DateTaken")
        value = ps.GetValue(pkey).GetValue()
        if value is None:
            return None
        from datetime import timezone as tz
        if hasattr(value, 'tzinfo') and value.tzinfo is not None:
            local_dt = value.astimezone()
            result   = datetime(local_dt.year, local_dt.month, local_dt.day,
                                local_dt.hour, local_dt.minute, local_dt.second)
        else:
            result = datetime(value.year, value.month, value.day,
                              value.hour, value.minute, value.second)
        print(f"    Windows Date Taken (propsys): {result}")
        return result
    except Exception:
        return None
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def get_windows_media_created(filepath: Path) -> datetime | None:
    """
    Read System.Media.DateEncoded via propsys API.
    Returns timezone-naive datetime in local time, or None.
    Uses raw PROPVARIANT — no string parsing, no locale dependency.
    """
    if not has_win32:
        return None
    try:
        pythoncom.CoInitialize()
        ps    = propsys.SHGetPropertyStoreFromParsingName(
                    str(filepath), None,
                    shellcon.GPS_BESTEFFORT,
                    propsys.IID_IPropertyStore)
        pkey  = propsys.PSGetPropertyKeyFromName("System.Media.DateEncoded")
        value = ps.GetValue(pkey).GetValue()
        if value is None:
            return None
        from datetime import timezone as tz
        if hasattr(value, 'tzinfo') and value.tzinfo is not None:
            local_dt = value.astimezone()
            result   = datetime(local_dt.year, local_dt.month, local_dt.day,
                                local_dt.hour, local_dt.minute, local_dt.second)
        else:
            result = datetime(value.year, value.month, value.day,
                              value.hour, value.minute, value.second)
        print(f"    Windows Media Created (propsys): {result}")
        return result
    except Exception:
        return None
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def set_windows_date_taken(filepath: Path, new_date: datetime) -> bool:
    """Write System.Photo.DateTaken via propsys API."""
    if not has_win32:
        return False
    try:
        pythoncom.CoInitialize()
        ps   = propsys.SHGetPropertyStoreFromParsingName(
                   str(filepath), None,
                   shellcon.GPS_READWRITE,
                   propsys.IID_IPropertyStore)
        pkey = propsys.PSGetPropertyKeyFromName("System.Photo.DateTaken")
        import pywintypes
        pytime = pywintypes.Time(new_date)
        pv     = propsys.PROPVARIANTType(pytime)
        ps.SetValue(pkey, pv)
        ps.Commit()
        print(f"  -> Updated Windows Date Taken = {new_date}")
        return True
    except Exception:
        pass
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return False


# =============================================================================
# Metadata extraction  (per file type)
# =============================================================================

def get_image_datetime(filepath: Path) -> tuple[datetime | None, str | None, str | None]:
    """
    Extract datetime from regular image files (non-RAW).

    Priority:
      1. DateTimeOriginal  (Exif.Photo.DateTimeOriginal  0x9003)
      2. DateTime          (Exif.Image.DateTime          0x0132)
      3. DateTimeDigitized (Exif.Photo.DateTimeDigitized 0x9004)
      4. Windows Date Taken  (last resort)

    Returns: (datetime, unparsed_string, source_name)
    """
    if EXIFTOOL_AVAILABLE:
        try:
            cmd = [
                'exiftool', '-json', '-m',
                '-DateTimeOriginal', '-ModifyDate', '-CreateDate',
                '-OffsetTimeOriginal', '-OffsetTime', '-OffsetTimeDigitized',
                str(filepath)
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            if result.returncode == 0:
                data = json.loads(result.stdout)
                if data:
                    tags             = data[0]
                    offset_original  = tags.get('OffsetTimeOriginal', '')
                    offset_modify    = tags.get('OffsetTime', '')
                    offset_digitized = tags.get('OffsetTimeDigitized', '')

                    # 1. DateTimeOriginal (0x9003)
                    if tags.get('DateTimeOriginal'):
                        value = str(tags['DateTimeOriginal'])
                        if offset_original:
                            value += offset_original
                        print(f"  -> exiftool DateTimeOriginal (0x9003): '{value}'")
                        parsed = parse_exif_date(value)
                        return (parsed, None, "DateTimeOriginal (0x9003)") if parsed \
                               else (None, value, "DateTimeOriginal (0x9003) (unparsed)")

                    # 2. ModifyDate (0x0132)
                    if tags.get('ModifyDate'):
                        value = str(tags['ModifyDate'])
                        if offset_modify:
                            value += offset_modify
                        print(f"  -> exiftool DateTime (0x0132): '{value}'")
                        parsed = parse_exif_date(value)
                        return (parsed, None, "DateTime (0x0132)") if parsed \
                               else (None, value, "DateTime (0x0132) (unparsed)")

                    # 3. CreateDate (0x9004)
                    if tags.get('CreateDate'):
                        value = str(tags['CreateDate'])
                        if offset_digitized:
                            value += offset_digitized
                        print(f"  -> exiftool DateTimeDigitized (0x9004): '{value}'")
                        parsed = parse_exif_date(value)
                        return (parsed, None, "DateTimeDigitized (0x9004)") if parsed \
                               else (None, value, "DateTimeDigitized (0x9004) (unparsed)")

        except subprocess.TimeoutExpired:
            print("  ! exiftool timed out")
        except json.JSONDecodeError as e:
            print(f"  ! exiftool JSON error: {e}")
        except Exception as e:
            print(f"  ! exiftool error: {e}")
    else:
        print("  -> exiftool not available, skipping EXIF extraction")

    # 4. Windows Date Taken (last resort)
    print("  -> Trying Windows Date Taken (last resort)...")
    win_date = get_windows_date_taken(filepath)
    if win_date:
        return win_date, None, "Windows Date Taken"

    return None, None, None


def get_raw_datetime(filepath: Path) -> tuple[datetime | None, str | None, str | None]:
    """
    Extract datetime from RAW image files.

    Priority:
      1. DateTimeOriginal (exiftool)
      2. CreateDate (exiftool)
      3. Windows Date Taken (last resort)

    Returns: (datetime, unparsed_string, source_name)
    """
    if EXIFTOOL_AVAILABLE:
        try:
            cmd = [
                'exiftool', '-json', '-m',
                '-DateTimeOriginal', '-CreateDate',
                '-OffsetTimeOriginal', '-OffsetTimeDigitized',
                str(filepath)
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            if result.returncode == 0:
                data = json.loads(result.stdout)
                if data:
                    tags             = data[0]
                    offset_original  = tags.get('OffsetTimeOriginal', '')
                    offset_digitized = tags.get('OffsetTimeDigitized', '')

                    # 1. DateTimeOriginal
                    if tags.get('DateTimeOriginal'):
                        value = str(tags['DateTimeOriginal'])
                        if offset_original:
                            value += offset_original
                        print(f"  -> exiftool DateTimeOriginal: '{value}'")
                        parsed = parse_exif_date(value)
                        return (parsed, None, "DateTimeOriginal") if parsed \
                               else (None, value, "DateTimeOriginal (unparsed)")

                    # 2. CreateDate
                    if tags.get('CreateDate'):
                        value = str(tags['CreateDate'])
                        if offset_digitized:
                            value += offset_digitized
                        print(f"  -> exiftool CreateDate: '{value}'")
                        parsed = parse_exif_date(value)
                        return (parsed, None, "CreateDate") if parsed \
                               else (None, value, "CreateDate (unparsed)")

        except subprocess.TimeoutExpired:
            print("  ! exiftool timed out")
        except json.JSONDecodeError as e:
            print(f"  ! exiftool JSON error: {e}")
        except Exception as e:
            print(f"  ! exiftool error: {e}")
    else:
        print("  -> exiftool not available, skipping EXIF extraction")

    # 3. Windows Date Taken (last resort)
    print("  -> Trying Windows Date Taken (last resort)...")
    win_date = get_windows_date_taken(filepath)
    if win_date:
        return win_date, None, "Windows Date Taken"

    return None, None, None


def get_video_datetime(filepath: Path) -> tuple[datetime | None, str | None, str | None]:
    """
    Extract datetime from video files.

    Priority:
      1. creation_time (ffprobe format tags, stream tags as fallback)
      2. date (ffprobe)
      3. date_recorded (ffprobe)
      4. Windows Media Created (last resort)
      5. Windows Date Taken (last resort)

    Returns: (datetime, unparsed_string, source_name)
    """
    if FFPROBE_AVAILABLE:
        try:
            # Request both format-level and stream-level tags in one call
            cmd = [
                'ffprobe', '-v', 'quiet', '-print_format', 'json',
                '-show_entries',
                'format_tags=creation_time,date,date_recorded:stream_tags=creation_time',
                str(filepath)
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            if result.returncode == 0:
                data          = json.loads(result.stdout)
                creation_time = None
                date_val      = None
                date_recorded = None

                # Check format tags first
                if 'format' in data and 'tags' in data['format']:
                    fmt_tags      = data['format']['tags']
                    creation_time = fmt_tags.get('creation_time')
                    date_val      = fmt_tags.get('date')
                    date_recorded = fmt_tags.get('date_recorded')

                # Fallback to stream tags if creation_time absent at format level
                if not creation_time and 'streams' in data:
                    for stream in data['streams']:
                        ct = stream.get('tags', {}).get('creation_time')
                        if ct:
                            creation_time = ct
                            print("  -> creation_time found in stream tags (not format tags)")
                            break

                # 1. creation_time
                if creation_time:
                    print(f"  -> ffprobe creation_time: '{creation_time}'")
                    parsed = parse_datetime_with_timezone(creation_time)
                    return (parsed, None, "creation_time") if parsed \
                           else (None, creation_time, "creation_time (unparsed)")

                # 2. date
                if date_val:
                    print(f"  -> ffprobe date: '{date_val}'")
                    parsed = parse_datetime_with_timezone(date_val)
                    return (parsed, None, "date") if parsed \
                           else (None, date_val, "date (unparsed)")

                # 3. date_recorded
                if date_recorded:
                    print(f"  -> ffprobe date_recorded: '{date_recorded}'")
                    parsed = parse_datetime_with_timezone(date_recorded)
                    return (parsed, None, "date_recorded") if parsed \
                           else (None, date_recorded, "date_recorded (unparsed)")

        except subprocess.TimeoutExpired:
            print("  ! ffprobe timed out")
        except Exception as e:
            print(f"  ! ffprobe error: {e}")
    else:
        print("  -> ffprobe not available, skipping video metadata extraction")

    # 4. Windows Media Created (last resort)
    print("  -> Trying Windows Media Created (last resort)...")
    win_date = get_windows_media_created(filepath)
    if win_date:
        return win_date, None, "Windows Media Created"

    # 5. Windows Date Taken (last resort)
    print("  -> Trying Windows Date Taken (last resort)...")
    win_date = get_windows_date_taken(filepath)
    if win_date:
        return win_date, None, "Windows Date Taken"

    return None, None, None


# =============================================================================
# Metadata writing
# =============================================================================

def update_image_metadata(filepath: Path, new_date: datetime) -> bool:
    """Update DateTimeOriginal, CreateDate, ModifyDate via exiftool."""
    if not EXIFTOOL_AVAILABLE:
        print("  ! exiftool not available - cannot update image metadata")
        return False
    try:
        exif_date = new_date.strftime("%Y:%m:%d %H:%M:%S")
        cmd = [
            'exiftool', '-m', '-overwrite_original',
            f'-DateTimeOriginal={exif_date}',
            f'-CreateDate={exif_date}',
            f'-ModifyDate={exif_date}',
            str(filepath)
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        if result.returncode == 0:
            if "0 image files updated" in result.stdout:
                print("  ! exiftool reported 0 files updated")
                return False
            print(f"  -> Updated image metadata: DateTimeOriginal, CreateDate, ModifyDate = {new_date}")
            return True
        else:
            error_msg = (result.stderr or result.stdout).strip()
            if "Not a valid" in error_msg:
                print(f"  ! File format mismatch: {error_msg[:100]}")
            else:
                print(f"  ! exiftool failed: {error_msg[:100]}")
            return False
    except subprocess.TimeoutExpired:
        print("  ! exiftool timed out")
        return False
    except Exception as e:
        print(f"  ! Error updating image metadata: {e}")
        return False


def update_raw_metadata(filepath: Path, new_date: datetime) -> bool:
    """Update DateTimeOriginal, CreateDate via exiftool."""
    if not EXIFTOOL_AVAILABLE:
        print("  ! exiftool not available - cannot update RAW metadata")
        return False
    try:
        exif_date = new_date.strftime("%Y:%m:%d %H:%M:%S")
        cmd = [
            'exiftool', '-m', '-overwrite_original',
            f'-DateTimeOriginal={exif_date}',
            f'-CreateDate={exif_date}',
            str(filepath)
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        if result.returncode == 0:
            if "0 image files updated" in result.stdout:
                print("  ! exiftool reported 0 files updated")
                return False
            print(f"  -> Updated RAW metadata: DateTimeOriginal, CreateDate = {new_date}")
            return True
        else:
            error_msg = (result.stderr or result.stdout).strip()
            if "Not a valid" in error_msg:
                print(f"  ! File format mismatch: {error_msg[:100]}")
            else:
                print(f"  ! exiftool failed: {error_msg[:100]}")
            return False
    except subprocess.TimeoutExpired:
        print("  ! exiftool timed out")
        return False
    except Exception as e:
        print(f"  ! Error updating RAW metadata: {e}")
        return False


def update_video_metadata(filepath: Path, new_date: datetime) -> bool:
    """
    Update date metadata in video files via ffmpeg.

    Quality & integrity guarantees:
      - -map 0          : all streams copied (video, audio, subtitles, attachments)
      - -c copy         : lossless, no re-encoding
      - -map_metadata 0 : all existing metadata tags preserved
      - Only creation_time, date, date_recorded are overridden
      - Atomic replace  : original replaced only after fully successful write
    """
    if not FFMPEG_AVAILABLE:
        print("  ! ffmpeg not available - cannot update video metadata")
        return False

    temp_file = filepath.parent / f"{filepath.stem}_temp{filepath.suffix}"

    try:
        from datetime import timezone as tz

        local_tz = datetime.now().astimezone().tzinfo
        local_dt = new_date.replace(tzinfo=local_tz)
        utc_dt   = local_dt.astimezone(tz.utc)
        iso_date = utc_dt.strftime("%Y-%m-%dT%H:%M:%S.000000Z")

        print(f"  -> Converting local {new_date.strftime('%Y-%m-%d %H:%M:%S')} "
              f"to UTC {utc_dt.strftime('%Y-%m-%d %H:%M:%S')}Z for storage")

        def build_cmd(map_args: list[str]) -> list[str]:
            return [
                'ffmpeg',
                '-i',            str(filepath),
                *map_args,
                '-c',            'copy',
                '-map_metadata', '0',
                '-metadata',     f'creation_time={iso_date}',
                '-metadata',     f'date={iso_date}',
                '-metadata',     f'date_recorded={iso_date}',
                '-y',
                '-loglevel',     'error',
                str(temp_file)
            ]

        # Attempt 1: copy every stream, but don't abort if a stream has a
        # type ffmpeg can't map into the container (e.g. a "tmcd" timecode
        # track some phones/cameras embed as stream 0 with codec_id "none",
        # or other data/metadata tracks) — skip it instead of failing.
        cmd = build_cmd(['-map', '0', '-ignore_unknown'])
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)

        # Attempt 2 (fallback): if that still failed on a stream/tag/codec
        # issue, retry keeping only video/audio/subtitle streams, explicitly
        # dropping whatever unmappable stream tripped up attempt 1.
        retried = False
        if result.returncode != 0:
            err = (result.stderr or '').lower()
            if any(s in err for s in ('tag for codec', 'not currently supported in container', 'unsupported codec')):
                if temp_file.exists():
                    temp_file.unlink()
                cmd = build_cmd(['-map', '0:v?', '-map', '0:a?', '-map', '0:s?'])
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
                retried = True

        if result.returncode == 0 and temp_file.exists() and temp_file.stat().st_size > 0:
            try:
                # Atomic: original replaced only if this succeeds
                temp_file.replace(filepath)
                note = " (dropped an unsupported stream, e.g. a timecode track)" if retried else ""
                print(f"  -> Updated video metadata: creation_time, date, date_recorded = "
                      f"{new_date} (stored as UTC){note}")
                return True
            except Exception as e:
                print(f"  ! Error replacing file: {e}")
                if temp_file.exists():
                    temp_file.unlink()
                return False
        else:
            if temp_file.exists():
                temp_file.unlink()
            if result.stderr:
                print(f"  ! ffmpeg failed: {result.stderr.strip()[:200]}")
            return False

    except subprocess.TimeoutExpired:
        print("  ! ffmpeg timed out")
        if temp_file.exists():
            temp_file.unlink()
        return False
    except Exception as e:
        if temp_file.exists():
            temp_file.unlink()
        print(f"  ! Error updating video metadata: {e}")
        return False


def set_file_system_dates(filepath: Path, new_date: datetime) -> bool:
    """Set file system mtime (Date Modified)."""
    try:
        timestamp = new_date.timestamp()
        os.utime(filepath, (timestamp, timestamp))
        print(f"  -> Updated file system Date Modified = {new_date}")
        return True
    except Exception as e:
        print(f"  ! Error setting file system date: {e}")
        return False


# =============================================================================
# Decision helpers
# =============================================================================

# All options are 3-tuples: (datetime, display_description, source_category)
# source_category values: "folder" | "metadata" | "filename" | "date_modified" | "manual"

def drop_redundant_whatsapp_filename_option(
    options: list[tuple[datetime, str, str]]
) -> list[tuple[datetime, str, str]]:
    """
    A WhatsApp filename date carries no real time info (fixed noon
    placeholder), so it's only useful when it's the sole source of a date.
    If any other option already covers the same calendar date, the
    WhatsApp filename option is dropped rather than shown as a choice.
    """
    if not options:
        return options

    def is_wa_filename(entry: tuple[datetime, str, str]) -> bool:
        _, source, category = entry
        return category == "filename" and "WhatsApp" in source

    other_dates = {
        date.date() for date, source, category in options
        if not is_wa_filename((date, source, category))
    }

    return [
        entry for entry in options
        if not (is_wa_filename(entry) and entry[0].date() in other_dates)
    ]


def merge_duplicate_options(
    options: list[tuple[datetime, str, str]]
) -> list[tuple[datetime, str, str]]:
    """
    Merge options that agree on date, hour, and minute - seconds alone
    are usually just rounding noise between sources, not a real
    disagreement. Source descriptions are combined into a numbered list.
    When a merged group includes a metadata-sourced entry, that entry's
    exact datetime (with real seconds) is kept as the representative
    value, since metadata is the most authoritative source.
    """
    if not options:
        return options

    groups: dict[tuple, list[tuple[datetime, str, str]]] = {}
    order:  list[tuple] = []
    for date, source, category in options:
        key = (date.year, date.month, date.day, date.hour, date.minute)
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append((date, source, category))

    merged: list[tuple[datetime, str, str]] = []
    for key in order:
        entries = groups[key]
        if len(entries) == 1:
            merged.append(entries[0])
            continue

        metadata_entry = next((e for e in entries if e[2] == "metadata"), None)
        rep_date     = metadata_entry[0] if metadata_entry else entries[0][0]
        rep_category = metadata_entry[2] if metadata_entry else entries[0][2]
        combined     = ", ".join(f"{i+1}. {s}" for i, (_, s, _) in enumerate(entries))
        merged.append((rep_date, combined, rep_category))

    return merged


def ask_user_for_date_choice(
    filepath: Path,
    options:  list[tuple[datetime, str, str]]
) -> tuple[datetime, str, str]:
    """
    Present numbered options to the user.
    Returns (datetime, description, source_category).
    Manual entry is always offered as the last option.

    Before prompting: redundant WhatsApp filename options are dropped,
    and options agreeing to the minute are merged. If that leaves only
    one option, it's used automatically without prompting.
    """
    options = drop_redundant_whatsapp_filename_option(options)
    options = merge_duplicate_options(options)

    if len(options) == 1:
        date, source, category = options[0]
        print(f"  -> All remaining sources agree on this date/time - using {source}")
        return date, source, category

    print(f"\n? Multiple date options for: {filepath.name}")
    print("  Please choose which date to use:")

    for i, (date, source, _) in enumerate(options, 1):
        range_note = "" if is_date_valid(date) else " - OUT OF VALID RANGE"
        print(f"  {i}. {date.strftime('%Y-%m-%d %H:%M:%S')} (from {source}){range_note}")

    manual_num = len(options) + 1
    print(f"  {manual_num}. Enter date manually")

    while True:
        choice = input(f"  Choose (1-{manual_num}): ").strip()
        try:
            n = int(choice)
            if 1 <= n <= len(options):
                chosen_date, chosen_source, chosen_category = options[n - 1]
                if not is_date_valid(chosen_date):
                    print(f"  Selected date {chosen_date} is out of valid range!")
                    if input("  Use it anyway? (y/n): ").strip().lower() != 'y':
                        continue
                return chosen_date, chosen_source, chosen_category
            elif n == manual_num:
                manual_date = ask_for_manual_date(filepath)
                return manual_date, "manual entry", "manual"
        except ValueError:
            pass
        print(f"  Invalid choice. Please enter 1-{manual_num}.")


def ask_for_manual_date(filepath: Path) -> datetime:
    """Prompt for free-form date input; loops until parseable."""
    print("  Enter date manually")
    print("  Formats: DD-MM-YYYY, DD/MM/YYYY, DD-MM-YYYY HH:MM:SS, DD Month YYYY")
    print("  Examples: 25-12-2023   25/12/2023 14:30:00   25 December 2023")
    print(f"  Valid range: {MIN_DATE.date()} to {MAX_DATE.date()}")

    while True:
        parsed = parse_user_date_input(input("  Date: ").strip())
        if parsed:
            if not is_date_valid(parsed):
                print(f"  Date {parsed} is out of valid range!")
                if input("  Use it anyway? (y/n): ").strip().lower() != 'y':
                    continue
            return parsed
        print("  Invalid date format. Please try again.")


def parse_user_date_input(date_str: str) -> datetime | None:
    date_str = date_str.strip()
    for fmt in ('%d-%m-%Y %H:%M:%S', '%d/%m/%Y %H:%M:%S',
                '%d-%m-%Y %H:%M',    '%d/%m/%Y %H:%M',
                '%d-%m-%Y %I:%M %p', '%d/%m/%Y %I:%M %p',
                '%d-%m-%Y %I:%M:%S %p', '%d/%m/%Y %I:%M:%S %p',
                '%d-%m-%Y',          '%d/%m/%Y',
                '%d %B %Y',          '%d %b %Y',
                '%d %B %Y %H:%M:%S', '%d %B %Y %H:%M',
                '%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
        try:
            parsed = datetime.strptime(date_str, fmt)
            if parsed.hour == 0 and parsed.minute == 0 and parsed.second == 0:
                if '%H' not in fmt and '%I' not in fmt:
                    parsed = parsed.replace(hour=12)
            return parsed
        except ValueError:
            continue
    return None


def should_auto_select_filename(
    folder_date:     datetime | None,
    metadata_date:   datetime | None,
    filename_date:   datetime | None,
    metadata_source: str | None = None,
) -> tuple[datetime, str, str] | None:
    """
    Determine whether a date can be auto-selected without user interaction.

    Preconditions (all must hold):
      - All three sources exist
      - All share the same calendar date
      - metadata and filename agree on hour and minute

    Seconds logic:
      - Both 00                         -> use metadata
      - metadata 00,  filename nonzero  -> use filename
      - metadata nonzero, filename 00   -> use metadata
      - both same nonzero               -> use metadata
      - both different nonzero          -> use metadata

    Returns (datetime, description, source_category) or None.
    """
    if not (folder_date and metadata_date and filename_date):
        return None
    if not (folder_date.date() == metadata_date.date() == filename_date.date()):
        return None
    if metadata_date.hour != filename_date.hour:
        return None
    if metadata_date.minute != filename_date.minute:
        return None

    meta_sec = metadata_date.second
    file_sec = filename_date.second
    src      = metadata_source or "Metadata"

    if meta_sec == 0 and file_sec == 0:
        return (metadata_date,
                "Metadata (auto-selected: all sources match)",
                "metadata")
    elif meta_sec == 0 and file_sec != 0:
        return (filename_date,
                "Filename (auto-selected: more precise seconds than metadata)",
                "filename")
    elif meta_sec != 0 and file_sec == 0:
        return (metadata_date,
                f"{src} (auto-selected: more precise seconds than filename)",
                "metadata")
    elif meta_sec == file_sec:
        return (metadata_date,
                "Metadata (auto-selected: all sources match including seconds)",
                "metadata")
    else:
        return (metadata_date,
                "Metadata (auto-selected)",
                "metadata")


# =============================================================================
# Main processing loop
# =============================================================================

def process_folder(folder_path: str) -> None:
    folder         = Path(folder_path)
    all_extensions = IMAGE_EXTENSIONS | RAW_EXTENSIONS | VIDEO_EXTENSIONS

    # Per-run counters
    processed_folder        = 0
    processed_metadata      = 0
    processed_filename      = 0
    processed_date_modified = 0
    processed_manual        = 0
    metadata_updated        = 0
    windows_dt_updated      = 0
    skipped_files           = 0
    error_files: list[tuple[str, str]] = []

    print(f"Processing media files in: {folder_path}\n")

    if not EXIFTOOL_AVAILABLE:
        print("WARNING: exiftool not installed — image/RAW metadata cannot be read or updated")
        print("         Install from: https://exiftool.org/\n")
    if not FFMPEG_AVAILABLE or not FFPROBE_AVAILABLE:
        print("WARNING: ffmpeg/ffprobe not installed — video metadata cannot be read or updated\n")
    if sys.platform == 'win32' and not has_win32:
        print("WARNING: pywin32 not installed — Windows Date Taken cannot be accessed")
        print("         Install with: pip install pywin32\n")

    print(f"Force metadata update: {'ENABLED' if FORCE_METADATA_UPDATE else 'DISABLED'}\n")

    media_files = sorted(
        [f for f in folder.rglob('*')
         if f.is_file() and f.suffix.lower() in all_extensions],
        key=lambda x: (x.parent, x.name.lower())
    )

    print(f"Found {len(media_files)} media files to process\n")
    print("=" * 60 + "\n")

    for i, filepath in enumerate(media_files, 1):
        file_type = get_file_type_name(filepath)
        print(f"[{i}/{len(media_files)}] Processing ({file_type}): {filepath.name}")
        print(f"  Path: {filepath.parent}")

        # ── File-type mismatch check ──────────────────────────────────────────
        if is_image_file(filepath) or is_raw_file(filepath):
            fixed = fix_file_extension_mismatch(filepath)
            if fixed is None:
                print("  -> Skipping file (user choice)")
                skipped_files += 1
                print()
                continue
            elif fixed != filepath:
                filepath  = fixed
                file_type = get_file_type_name(filepath)
                print(f"  -> Now processing as ({file_type}): {filepath.name}")

        # ── Per-file state ────────────────────────────────────────────────────
        date_found               = None
        source                   = None
        source_category          = None
        update_metadata          = False
        metadata_needs_overwrite = False

        # Step 1: Date Modified
        date_modified = get_file_date_modified(filepath)
        if date_modified:
            print(f"  -> Date Modified: {date_modified}")

        # Step 2: Folder date
        folder_date = extract_date_from_folder(filepath.parent)
        if folder_date:
            print(f"  -> Folder date: {folder_date.date()}")

        # Step 3: Metadata date
        metadata_date     = None
        metadata_unparsed = None
        metadata_source   = None

        if is_raw_file(filepath):
            metadata_date, metadata_unparsed, metadata_source = get_raw_datetime(filepath)
        elif is_image_file(filepath):
            metadata_date, metadata_unparsed, metadata_source = get_image_datetime(filepath)
        elif is_video_file(filepath):
            metadata_date, metadata_unparsed, metadata_source = get_video_datetime(filepath)

        if metadata_date:
            print(f"  -> Metadata datetime: {metadata_date} (from {metadata_source})")
        elif metadata_unparsed:
            print(f"  -> Metadata found but unparsed: '{metadata_unparsed}' (from {metadata_source})")
        else:
            print("  -> No metadata datetime found")

        # Step 4: Filename date
        is_whatsapp    = is_whatsapp_file(filepath)
        filename_dates = extract_date_from_filename(filepath, include_whatsapp=True)
        filename_date     = None
        filename_has_time = False
        filename_match    = None

        if filename_dates:
            filename_date, filename_format, filename_match, filename_has_time = filename_dates[0]
            time_info = "with time" if filename_has_time else "no time"
            if is_whatsapp:
                print(f"  -> Filename datetime: {filename_date} ({time_info}, WhatsApp - unreliable)")
            else:
                print(f"  -> Filename datetime: {filename_date} ({time_info}, pattern: {filename_format})")

        # ── Helpers ───────────────────────────────────────────────────────────
        def get_filename_label(prefix: str = "Filename") -> str:
            if is_whatsapp:
                return f"{prefix}: {filename_match} (WhatsApp - unreliable)"
            return f"{prefix}: {filename_match}"

        def add_date_modified_options(
            opts:         list[tuple[datetime, str, str]],
            folder_date:  datetime | None,
            date_modified: datetime | None
        ) -> None:
            if date_modified:
                opts.append((date_modified, "Date Modified", "date_modified"))
                if folder_date and folder_date.date() != date_modified.date():
                    combined = folder_date.replace(
                        hour=date_modified.hour,
                        minute=date_modified.minute,
                        second=date_modified.second
                    )
                    opts.append((
                        combined,
                        f"Folder date ({folder_date.date()}) + "
                        f"Date Modified time ({date_modified.strftime('%H:%M:%S')})",
                        "folder"
                    ))

        # Step 5: Handle unparseable metadata
        if metadata_unparsed and not metadata_date:
            print("  -> Metadata exists but could not be parsed automatically")
            print(f"  -> Please enter the date manually based on: '{metadata_unparsed}'")
            metadata_date            = ask_for_manual_date(filepath)
            metadata_source          = "manual interpretation of metadata"
            metadata_needs_overwrite = True

        # Step 6: Decision tree
        date_from_user_choice  = False
        original_metadata_date = metadata_date

        # Pre-check: auto-select when all sources agree
        auto_result = should_auto_select_filename(
            folder_date, metadata_date, filename_date, metadata_source)

        if auto_result is not None:
            date_found, source, source_category = auto_result
            print(f"  -> {source}")

        # Case A: Folder + Metadata both exist
        elif folder_date and metadata_date:
            if folder_date.date() == metadata_date.date():
                date_found      = metadata_date
                source          = f"{metadata_source} (matches folder date)"
                source_category = "metadata"
                print("  -> Same date in folder and metadata - using metadata datetime")
            else:
                print(f"  -> Different dates: folder={folder_date.date()}, "
                      f"metadata={metadata_date.date()}")
                combined = folder_date.replace(
                    hour=metadata_date.hour,
                    minute=metadata_date.minute,
                    second=metadata_date.second
                )
                opts = [
                    (combined,
                     f"Folder date ({folder_date.date()}) + "
                     f"metadata time ({metadata_date.strftime('%H:%M:%S')})",
                     "folder"),
                    (metadata_date, f"Metadata: {metadata_source}", "metadata"),
                ]
                if filename_date:
                    opts.append((filename_date, get_filename_label(), "filename"))
                add_date_modified_options(opts, folder_date, date_modified)
                date_found, source, source_category = ask_user_for_date_choice(filepath, opts)
                date_from_user_choice = True

        # Case B: Folder exists, Metadata absent
        elif folder_date and not metadata_date:
            opts = []
            if filename_date and filename_has_time:
                combined = folder_date.replace(
                    hour=filename_date.hour,
                    minute=filename_date.minute,
                    second=filename_date.second
                )
                wa_note = " (WhatsApp - unreliable)" if is_whatsapp else ""
                opts.append((
                    combined,
                    f"Folder date ({folder_date.date()}) + "
                    f"filename time ({filename_date.strftime('%H:%M:%S')}){wa_note}",
                    "folder"
                ))
            if filename_date:
                opts.append((filename_date, get_filename_label("Filename datetime"), "filename"))
            opts.append((folder_date,
                         f"Folder date only ({folder_date.date()}, 12:00:00)",
                         "folder"))
            add_date_modified_options(opts, folder_date, date_modified)
            date_found, source, source_category = ask_user_for_date_choice(filepath, opts)
            date_from_user_choice = True

        # Case C: No folder, Metadata + Filename both exist
        elif not folder_date and metadata_date and filename_date:
            if metadata_date == filename_date:
                date_found      = metadata_date
                source          = f"{metadata_source} (matches filename date and time)"
                source_category = "metadata"
                print("  -> Same date and time in metadata and filename - using metadata")
            elif metadata_date.date() == filename_date.date():
                print(f"  -> Same date but different time: "
                      f"metadata={metadata_date.strftime('%H:%M:%S')}, "
                      f"filename={filename_date.strftime('%H:%M:%S')}")
                opts = [
                    (metadata_date, f"Metadata: {metadata_source}", "metadata"),
                    (filename_date, get_filename_label(), "filename"),
                ]
                if date_modified:
                    opts.append((date_modified, "Date Modified", "date_modified"))
                date_found, source, source_category = ask_user_for_date_choice(filepath, opts)
                date_from_user_choice = True
            else:
                print(f"  -> Different dates: metadata={metadata_date.date()}, "
                      f"filename={filename_date.date()}")
                opts = [
                    (metadata_date, f"Metadata: {metadata_source}", "metadata"),
                    (filename_date, get_filename_label(), "filename"),
                ]
                if date_modified:
                    opts.append((date_modified, "Date Modified", "date_modified"))
                date_found, source, source_category = ask_user_for_date_choice(filepath, opts)
                date_from_user_choice = True

        # Case D: No folder, Metadata only
        elif not folder_date and metadata_date and not filename_date:
            print("  -> Only metadata datetime found - asking for confirmation")
            opts = [(metadata_date,
                     f"Metadata: {metadata_source} (cannot verify, no other source)",
                     "metadata")]
            if date_modified:
                opts.append((date_modified, "Date Modified", "date_modified"))
            date_found, source, source_category = ask_user_for_date_choice(filepath, opts)
            date_from_user_choice = True

        # Case E: No folder, Filename only
        elif not folder_date and not metadata_date and filename_date:
            print("  -> Only filename datetime found - asking for confirmation")
            opts = [(filename_date, get_filename_label(), "filename")]
            if date_modified:
                opts.append((date_modified, "Date Modified", "date_modified"))
            date_found, source, source_category = ask_user_for_date_choice(filepath, opts)
            date_from_user_choice = True

        # Case F: Nothing found
        else:
            print("  -> No date found anywhere - showing available options")
            opts = []
            if date_modified:
                opts.append((date_modified, "Date Modified", "date_modified"))
            if opts:
                date_found, source, source_category = ask_user_for_date_choice(filepath, opts)
                date_from_user_choice = True
            else:
                print("  -> No Date Modified available - manual entry required")
                date_found      = ask_for_manual_date(filepath)
                source          = "manual entry"
                source_category = "manual"
                date_from_user_choice = True

        # Step 7: Validate auto-selected dates
        if date_found and not is_date_valid(date_found) and not date_from_user_choice:
            print(f"  Date {date_found} is out of valid range "
                  f"({MIN_DATE.date()} to {MAX_DATE.date()})!")
            if input("  Use this date anyway? (y/n): ").strip().lower() != 'y':
                opts = []
                if folder_date:
                    if metadata_date:
                        combined = folder_date.replace(
                            hour=metadata_date.hour,
                            minute=metadata_date.minute,
                            second=metadata_date.second
                        )
                        opts.append((combined, "Folder date + metadata time", "folder"))
                    elif filename_date and filename_has_time:
                        combined = folder_date.replace(
                            hour=filename_date.hour,
                            minute=filename_date.minute,
                            second=filename_date.second
                        )
                        opts.append((combined, "Folder date + filename time", "folder"))
                    else:
                        opts.append((folder_date,
                                     f"Folder date: {folder_date.date()}",
                                     "folder"))
                if metadata_date:
                    opts.append((metadata_date,
                                 f"Metadata: {metadata_source}",
                                 "metadata"))
                if filename_date:
                    opts.append((filename_date, get_filename_label(), "filename"))
                add_date_modified_options(opts, folder_date, date_modified)
                if opts:
                    date_found, source, source_category = ask_user_for_date_choice(filepath, opts)
                else:
                    date_found      = ask_for_manual_date(filepath)
                    source          = "manual entry (after out-of-range)"
                    source_category = "manual"
                date_from_user_choice = True

        # Step 8: Decide whether metadata write is needed
        if date_found:
            if original_metadata_date is None:
                update_metadata = True
                print("  -> No existing metadata - will create")
            elif date_found != original_metadata_date:
                update_metadata = True
                print(f"  -> Date differs from metadata "
                      f"({original_metadata_date} -> {date_found}) - will update")
            elif metadata_needs_overwrite:
                update_metadata = True
                print("  -> Metadata contained unparseable raw string - will overwrite")
            elif FORCE_METADATA_UPDATE:
                update_metadata = True
                print("  -> Force update enabled - will update metadata")
            else:
                update_metadata = False
                print("  -> Metadata already correct - skipping metadata update")

        # Step 9: Write metadata
        if date_found and update_metadata:
            print(f"  -> Updating metadata to: {date_found}")
            success = False
            if is_raw_file(filepath):
                success = update_raw_metadata(filepath, date_found)
            elif is_image_file(filepath):
                success = update_image_metadata(filepath, date_found)
            elif is_video_file(filepath):
                success = update_video_metadata(filepath, date_found)
            else:
                print("  ! Unknown file type - cannot update metadata")
            if success:
                metadata_updated += 1
            else:
                error_files.append((str(filepath), "Failed to update metadata"))

        # Step 10: Update Windows Date Taken
        if date_found and has_win32:
            if set_windows_date_taken(filepath, date_found):
                windows_dt_updated += 1

        # Step 11: Update file system Date Modified
        if date_found:
            if set_file_system_dates(filepath, date_found):
                print(f"  File date updated to {date_found} (from {source})")
                # Increment counter using explicit category — no string parsing
                if source_category == "folder":
                    processed_folder += 1
                elif source_category == "metadata":
                    processed_metadata += 1
                elif source_category == "filename":
                    processed_filename += 1
                elif source_category == "date_modified":
                    processed_date_modified += 1
                elif source_category == "manual":
                    processed_manual += 1
            else:
                print("  ! Failed to set file system date")
                error_files.append((str(filepath), "Failed to set file system date"))

        print()

    # ── Console summary ───────────────────────────────────────────────────────
    total = (processed_folder + processed_metadata + processed_filename
             + processed_date_modified + processed_manual)

    print("\n" + "=" * 60)
    print("SUMMARY:")
    print("=" * 60)
    print(f"Files updated from folder name:    {processed_folder}")
    print(f"Files updated from metadata:       {processed_metadata}")
    print(f"Files updated from filename:       {processed_filename}")
    print(f"Files updated from Date Modified:  {processed_date_modified}")
    print(f"Files updated from manual entry:   {processed_manual}")
    print(f"Metadata written into files:       {metadata_updated}")
    print(f"Windows Date Taken updated:        {windows_dt_updated}")
    print(f"Files skipped:                     {skipped_files}")
    print(f"Total files processed:             {total}")
    print(f"Total files with errors:           {len(error_files)}")

    # ── Report file ───────────────────────────────────────────────────────────
    # Saved next to fix.py itself (not the processed folder, not the venv),
    # regardless of what the current working directory happens to be.
    script_dir  = Path(__file__).resolve().parent
    report_path = script_dir / "media_date_update_report.txt"

    # Preserve any existing report with a timestamped name
    if report_path.exists():
        try:
            existing_ts   = datetime.fromtimestamp(report_path.stat().st_mtime)
            ts_str        = existing_ts.strftime("%Y%m%d_%H%M%S")
            backup_path   = report_path.parent / f"media_date_update_report_{ts_str}.txt"
            report_path.rename(backup_path)
            print(f"\nPrevious report preserved as: {backup_path.name}")
        except Exception as e:
            print(f"\nCould not rename existing report: {e}")

    with open(report_path, 'w', encoding='utf-8') as f:
        f.write("Media Date Update Report\n")
        f.write("=" * 60 + "\n")
        f.write(f"Generated on:        {datetime.now()}\n")
        f.write(f"Valid date range:     {MIN_DATE.date()} to {MAX_DATE.date()}\n")
        f.write(f"Force metadata update: {'ENABLED' if FORCE_METADATA_UPDATE else 'DISABLED'}\n")
        f.write(f"Folder processed:    {folder_path}\n")
        f.write("=" * 60 + "\n\n")

        f.write("DEPENDENCIES:\n")
        f.write(f"- exiftool: {'Available' if EXIFTOOL_AVAILABLE else 'NOT INSTALLED'}\n")
        f.write(f"- ffmpeg:   {'Available' if FFMPEG_AVAILABLE   else 'NOT INSTALLED'}\n")
        f.write(f"- ffprobe:  {'Available' if FFPROBE_AVAILABLE  else 'NOT INSTALLED'}\n")
        f.write(f"- pywin32:  {'Available' if has_win32          else 'NOT INSTALLED'}\n\n")

        f.write("EXTRACTION PRIORITY:\n")
        f.write("Images:\n")
        f.write("  1. DateTimeOriginal (Exif.Photo.DateTimeOriginal 0x9003)\n")
        f.write("  2. DateTime (Exif.Image.DateTime 0x0132)\n")
        f.write("  3. DateTimeDigitized (Exif.Photo.DateTimeDigitized 0x9004)\n")
        f.write("  4. Windows Date Taken (last resort)\n\n")
        f.write("Videos:\n")
        f.write("  1. creation_time (ffprobe format tags; stream tags as fallback)\n")
        f.write("  2. date (ffprobe)\n")
        f.write("  3. date_recorded (ffprobe)\n")
        f.write("  4. Windows Media Created (last resort)\n")
        f.write("  5. Windows Date Taken (last resort)\n\n")
        f.write("RAW:\n")
        f.write("  1. DateTimeOriginal (exiftool)\n")
        f.write("  2. CreateDate (exiftool)\n")
        f.write("  3. Windows Date Taken (last resort)\n\n")

        f.write("FILE TYPES:\n")
        f.write(f"Images: {', '.join(sorted(IMAGE_EXTENSIONS))}\n")
        f.write(f"RAW:    {', '.join(sorted(RAW_EXTENSIONS))}\n")
        f.write(f"Videos: {', '.join(sorted(VIDEO_EXTENSIONS))}\n\n")

        f.write("PROPERTIES UPDATED:\n")
        f.write("Images: DateTimeOriginal, CreateDate, ModifyDate (via exiftool)\n")
        f.write("RAW:    DateTimeOriginal, CreateDate (via exiftool)\n")
        f.write("Videos: creation_time, date, date_recorded (via ffmpeg, stored as UTC)\n")
        f.write("All:    Windows Date Taken (via propsys API, locale-independent)\n")
        f.write("All:    File system Date Modified (via os.utime)\n\n")

        f.write("VIDEO QUALITY & INTEGRITY:\n")
        f.write("- -map 0          : all streams copied (video, audio, subtitles, attachments)\n")
        f.write("- -c copy         : lossless, no re-encoding\n")
        f.write("- -map_metadata 0 : all existing metadata tags preserved\n")
        f.write("- Only creation_time, date, date_recorded are overridden\n")
        f.write("- Atomic replace  : original replaced only after fully successful write\n\n")

        f.write("DECISION RULES:\n")
        f.write("Pre-check (all sources agree on date/hour/minute):\n")
        f.write("  Both seconds = 00              -> use metadata\n")
        f.write("  metadata=00, filename nonzero  -> use filename (more precise)\n")
        f.write("  metadata nonzero, filename=00  -> use metadata (more precise)\n")
        f.write("  both same nonzero seconds      -> use metadata\n")
        f.write("  both different nonzero seconds -> use metadata\n")
        f.write("A. Folder + Metadata (same date)         -> use metadata\n")
        f.write("A. Folder + Metadata (different date)    -> ask user\n")
        f.write("B. Folder only                           -> compare with filename, ask user\n")
        f.write("C. Metadata + Filename (same date+time)  -> use metadata\n")
        f.write("C. Metadata + Filename (same date, diff time) -> ask user\n")
        f.write("C. Metadata + Filename (different date)  -> ask user\n")
        f.write("D. Metadata only                         -> ask user for confirmation\n")
        f.write("E. Filename only                         -> ask user for confirmation\n")
        f.write("F. Nothing found                         -> Date Modified, then manual\n")
        f.write("WhatsApp files: filename shown with warning, never auto-selected\n\n")

        f.write("METADATA UPDATE RULES:\n")
        f.write("- Update when no metadata exists\n")
        f.write("- Update when chosen date differs from existing metadata\n")
        f.write("- Update when raw metadata string was unparseable (always overwrite)\n")
        f.write(f"- Force update: {'ENABLED' if FORCE_METADATA_UPDATE else 'DISABLED'}\n\n")

        f.write("REPORT BEHAVIOUR:\n")
        f.write("- Each run writes media_date_update_report.txt next to this script\n")
        f.write("- Any existing report is renamed with its own last-modified timestamp\n")
        f.write("  e.g. media_date_update_report_20260519_143022.txt\n\n")

        f.write("SUMMARY:\n")
        f.write(f"- Files updated from folder name:   {processed_folder}\n")
        f.write(f"- Files updated from metadata:      {processed_metadata}\n")
        f.write(f"- Files updated from filename:      {processed_filename}\n")
        f.write(f"- Files updated from Date Modified: {processed_date_modified}\n")
        f.write(f"- Files updated from manual entry:  {processed_manual}\n")
        f.write(f"- Metadata written into files:      {metadata_updated}\n")
        f.write(f"- Windows Date Taken updated:       {windows_dt_updated}\n")
        f.write(f"- Files skipped:                    {skipped_files}\n")
        f.write(f"- Total files processed:            {total}\n")
        f.write(f"- Total files with errors:          {len(error_files)}\n")

        if error_files:
            f.write("\n" + "=" * 60 + "\n")
            f.write(f"FILES WITH ERRORS ({len(error_files)}):\n")
            f.write("=" * 60 + "\n")
            for fp, err in sorted(error_files):
                f.write(f"{fp}: {err}\n")

    print(f"\nReport saved to: {report_path}")


# =============================================================================
# Entry point
# =============================================================================

def main():
    if len(sys.argv) < 2:
        print("Usage: python fix_photo_dates.py <folder_path>")
        print('Example: python fix_photo_dates.py "C:\\Users\\YourName\\Pictures"')
        sys.exit(1)

    folder_path = sys.argv[1]
    if not Path(folder_path).exists():
        print(f"Error: Folder not found: {folder_path}")
        sys.exit(1)

    print("=" * 60)
    print("MEDIA DATE FIXER")
    print("=" * 60)
    print(f"\nSystem Information:")
    print(f"- Current date/time: {CURRENT_DATE}")
    print(f"- Platform:          {sys.platform}")
    print(f"- Python version:    {sys.version.split()[0]}")
    print(f"- Valid date range:  {MIN_DATE.date()} to {MAX_DATE.date()}")
    print()

    # Check all dependencies once — results stored in module-level globals
    print("Checking dependencies...")

    global EXIFTOOL_AVAILABLE, FFMPEG_AVAILABLE, FFPROBE_AVAILABLE

    EXIFTOOL_AVAILABLE = check_exiftool_available()
    if EXIFTOOL_AVAILABLE:
        print("  exiftool found - image/RAW metadata support enabled")
    else:
        print("  exiftool NOT FOUND - image/RAW metadata cannot be read/updated")
        print("    Install from: https://exiftool.org/")
        print("    Windows: download exe, rename to exiftool.exe, add to PATH")
        print("    Mac:     brew install exiftool")
        print("    Linux:   sudo apt install libimage-exiftool-perl")

    FFMPEG_AVAILABLE = check_ffmpeg_available()
    if FFMPEG_AVAILABLE:
        print("  ffmpeg found - video metadata updates enabled")
    else:
        print("  ffmpeg NOT FOUND - video metadata cannot be updated")
        print("    Install from: https://ffmpeg.org/download.html")

    FFPROBE_AVAILABLE = check_ffprobe_available()
    if FFPROBE_AVAILABLE:
        print("  ffprobe found - video metadata reading enabled")
    else:
        print("  ffprobe NOT FOUND - video metadata cannot be read")

    if sys.platform == 'win32':
        if has_win32:
            print("  pywin32 found - Windows Date Taken property access enabled")
            print("    Note: pywin32 build 306+ required for Python 3.13 compatibility")
        else:
            print("  pywin32 NOT FOUND - Windows Date Taken cannot be accessed")
            print("    Install with: pip install pywin32")

    print()
    print("Supported file types:")
    print(f"  Images: {', '.join(sorted(e.upper().lstrip('.') for e in IMAGE_EXTENSIONS))}")
    print(f"  RAW:    {', '.join(sorted(e.upper().lstrip('.') for e in RAW_EXTENSIONS))}")
    print(f"  Videos: {', '.join(sorted(e.upper().lstrip('.') for e in VIDEO_EXTENSIONS))}")
    print()

    print("Metadata Extraction Priority:")
    print("  Images:")
    print("    1. DateTimeOriginal (Exif.Photo.DateTimeOriginal 0x9003)")
    print("    2. DateTime (Exif.Image.DateTime 0x0132)")
    print("    3. DateTimeDigitized (Exif.Photo.DateTimeDigitized 0x9004)")
    print("    4. Windows Date Taken (last resort)")
    print("  Videos:")
    print("    1. creation_time (ffprobe format tags; stream tags as fallback)")
    print("    2. date (ffprobe)")
    print("    3. date_recorded (ffprobe)")
    print("    4. Windows Media Created (last resort)")
    print("    5. Windows Date Taken (last resort)")
    print("  RAW:")
    print("    1. DateTimeOriginal (exiftool)")
    print("    2. CreateDate (exiftool)")
    print("    3. Windows Date Taken (last resort)")
    print()

    print("Properties Updated:")
    print("  Images: DateTimeOriginal, CreateDate, ModifyDate (via exiftool)")
    print("  RAW:    DateTimeOriginal, CreateDate (via exiftool)")
    print("  Videos: creation_time, date, date_recorded (via ffmpeg, stored as UTC)")
    print("  All:    Windows Date Taken (via propsys API, locale-independent)")
    print("  All:    File system Date Modified (via os.utime)")
    print()

    print("Video Quality & Integrity:")
    print("  -map 0          : all streams copied without exception")
    print("  -c copy         : lossless, no re-encoding")
    print("  -map_metadata 0 : all existing metadata tags preserved")
    print("  Only creation_time, date, date_recorded are overridden")
    print("  Atomic replace  : original replaced only after fully successful write")
    print()

    print("Timezone Handling:")
    print("  UTC timestamps (ending with Z) are converted to local time")
    print("  Timezone offsets (e.g. +05:30, -08:00) are converted to local time")
    print("  Timestamps without timezone info assumed to be local time already")
    print("  Video dates stored back as UTC (industry standard for containers)")
    print()

    print("Decision Rules:")
    print("  Pre-check (all sources agree on date/hour/minute):")
    print("    Both seconds = 00              -> use metadata")
    print("    metadata=00, filename nonzero  -> use filename (more precise)")
    print("    metadata nonzero, filename=00  -> use metadata (more precise)")
    print("    both same nonzero seconds      -> use metadata")
    print("    both different nonzero seconds -> use metadata")
    print("  A. Folder + Metadata (same date)              -> use metadata")
    print("  A. Folder + Metadata (different date)         -> ask user")
    print("  B. Folder only                                -> compare with filename, ask user")
    print("  C. Metadata + Filename (same date+time)       -> use metadata")
    print("  C. Metadata + Filename (same date, diff time) -> ask user")
    print("  C. Metadata + Filename (different date)       -> ask user")
    print("  D. Metadata only                              -> ask user for confirmation")
    print("  E. Filename only                              -> ask user for confirmation")
    print("  F. Nothing found                              -> Date Modified, then manual")
    print("  WhatsApp files: filename shown with warning, never auto-selected")
    print()

    print("Metadata Update Rules:")
    print("  Update when no metadata exists")
    print("  Update when chosen date differs from existing metadata")
    print("  Update when raw metadata string was unparseable (always overwrite)")
    print(f"  Force update: {'ENABLED' if FORCE_METADATA_UPDATE else 'DISABLED'}")
    print("  (Set FORCE_METADATA_UPDATE = True at top of script to always update)")
    print()

    print("Report Behaviour:")
    print("  Each run writes media_date_update_report.txt next to fix.py")
    print("  Any existing report is first renamed with its own timestamp")
    print("  e.g. media_date_update_report_20260519_143022.txt")
    print()

    print("File-Type Mismatch Detection:")
    print("  Reads magic bytes to detect actual file format")
    print("  Offers rename / keep / skip if extension does not match content")
    print("  Common case: HEIC files incorrectly named .jpg")
    print()

    process_folder(folder_path)


if __name__ == "__main__":
    main()

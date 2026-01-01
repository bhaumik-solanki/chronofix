import os
import sys
from datetime import datetime
from pathlib import Path
import re
import subprocess
import json

# Check Python version
if sys.version_info < (3, 10):
    print("This script requires Python 3.10 or higher")
    sys.exit(1)

# Optional imports for Windows
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
MIN_DATE = datetime(2022, 1, 1)
MAX_DATE = datetime.now().replace(hour=23, minute=59, second=59)
CURRENT_DATE = datetime.now()

# Set to True to always update metadata even when it already exists and matches
FORCE_METADATA_UPDATE = False

# File extension categories
IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif', 
                    '.heic', '.heif', '.webp'}
RAW_EXTENSIONS = {'.cr2', '.nef', '.arw', '.dng', '.raw', '.orf', '.sr2', '.pef', 
                  '.raf', '.rw2', '.rwl', '.srw', '.x3f'}
VIDEO_EXTENSIONS = {'.mp4', '.mov', '.avi', '.mkv', '.wmv', '.flv', '.webm', 
                    '.m4v', '.mpg', '.mpeg', '.3gp', '.mts', '.m2ts'}

print(f"Current system date/time: {CURRENT_DATE}")
print(f"Valid date range: {MIN_DATE.date()} to {MAX_DATE.date()}\n")


def is_date_valid(date: datetime) -> bool:
    """Check if date is within valid range"""
    return MIN_DATE <= date <= MAX_DATE


def is_whatsapp_file(filepath: Path) -> bool:
    """Check if filename indicates it's a WhatsApp file"""
    filename = filepath.name.lower()
    return 'whatsapp' in filename or '-wa' in filename or '_wa' in filename


def is_raw_file(filepath: Path) -> bool:
    """Check if file is a RAW image file"""
    return filepath.suffix.lower() in RAW_EXTENSIONS


def is_image_file(filepath: Path) -> bool:
    """Check if file is a regular image (non-RAW)"""
    return filepath.suffix.lower() in IMAGE_EXTENSIONS


def is_video_file(filepath: Path) -> bool:
    """Check if file is a video file"""
    return filepath.suffix.lower() in VIDEO_EXTENSIONS


def get_file_type_name(filepath: Path) -> str:
    """Get human-readable file type name"""
    if is_raw_file(filepath):
        return "RAW"
    elif is_image_file(filepath):
        return "Image"
    elif is_video_file(filepath):
        return "Video"
    return "Unknown"


def detect_actual_file_type(filepath: Path) -> str | None:
    """
    Detect the actual file type by reading file header (magic bytes).
    Returns the correct extension or None if unknown.
    """
    try:
        with open(filepath, 'rb') as f:
            header = f.read(32)
        
        # JPEG: starts with FF D8 FF
        if header[:3] == b'\xff\xd8\xff':
            return '.jpg'
        
        # PNG: starts with 89 50 4E 47 0D 0A 1A 0A
        if header[:8] == b'\x89PNG\r\n\x1a\n':
            return '.png'
        
        # HEIC/HEIF: contains 'ftyp' followed by 'heic', 'heix', 'hevc', 'mif1'
        if b'ftyp' in header[:12]:
            if b'heic' in header or b'heix' in header or b'mif1' in header or b'hevc' in header:
                return '.heic'
            if b'avif' in header:
                return '.avif'
            # MP4/MOV: ftyp followed by various types
            if b'mp4' in header or b'isom' in header or b'M4V' in header:
                return '.mp4'
            if b'qt' in header:
                return '.mov'
        
        # GIF: starts with GIF87a or GIF89a
        if header[:6] in (b'GIF87a', b'GIF89a'):
            return '.gif'
        
        # BMP: starts with BM
        if header[:2] == b'BM':
            return '.bmp'
        
        # TIFF: starts with II (little-endian) or MM (big-endian)
        if header[:4] in (b'II\x2a\x00', b'MM\x00\x2a'):
            return '.tiff'
        
        # WebP: starts with RIFF followed by WEBP
        if header[:4] == b'RIFF' and header[8:12] == b'WEBP':
            return '.webp'
        
        # RAW formats - CR2 (Canon)
        if header[:4] in (b'II\x2a\x00', b'MM\x00\x2a'):
            if b'CR' in header:
                return '.cr2'
        
        return None
        
    except Exception as e:
        print(f"    ! Error detecting file type: {e}")
        return None


def fix_file_extension_mismatch(filepath: Path) -> Path | None:
    """
    Check if file extension matches actual file type.
    If mismatch, offer to rename the file.
    Returns the new filepath if renamed, original filepath if no change needed, None if user skipped.
    """
    actual_type = detect_actual_file_type(filepath)
    current_ext = filepath.suffix.lower()
    
    if actual_type is None:
        return filepath  # Can't detect, proceed with original
    
    # Normalize extensions for comparison
    ext_mapping = {
        '.jpeg': '.jpg',
        '.tif': '.tiff',
    }
    normalized_current = ext_mapping.get(current_ext, current_ext)
    normalized_actual = ext_mapping.get(actual_type, actual_type)
    
    if normalized_current == normalized_actual:
        return filepath  # No mismatch
    
    print(f"\n  ⚠️ FILE TYPE MISMATCH DETECTED!")
    print(f"     Current extension: {current_ext}")
    print(f"     Actual file type: {actual_type}")
    print(f"     Options:")
    print(f"     1. Rename file to correct extension ({filepath.stem}{actual_type})")
    print(f"     2. Keep current name and try to process anyway")
    print(f"     3. Skip this file")
    
    while True:
        choice = input(f"     Choose (1-3): ").strip()
        if choice == '1':
            new_path = filepath.parent / f"{filepath.stem}{actual_type}"
            if new_path.exists():
                print(f"     ❌ Cannot rename: {new_path.name} already exists!")
                continue
            try:
                filepath.rename(new_path)
                print(f"     ✓ Renamed to: {new_path.name}")
                return new_path
            except Exception as e:
                print(f"     ❌ Error renaming file: {e}")
                continue
        elif choice == '2':
            return filepath
        elif choice == '3':
            return None
        else:
            print(f"     Invalid choice. Please enter 1, 2, or 3.")


def check_exiftool_available() -> bool:
    """Check if exiftool is available in the system"""
    try:
        result = subprocess.run(['exiftool', '-ver'], capture_output=True, text=True, timeout=10)
        return result.returncode == 0
    except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired):
        return False


def check_ffmpeg_available() -> bool:
    """Check if ffmpeg is available in the system"""
    try:
        subprocess.run(['ffmpeg', '-version'], capture_output=True, check=True, timeout=10)
        return True
    except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired):
        return False


def check_ffprobe_available() -> bool:
    """Check if ffprobe is available in the system"""
    try:
        subprocess.run(['ffprobe', '-version'], capture_output=True, check=True, timeout=10)
        return True
    except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired):
        return False


def extract_date_from_folder(folder_path: Path) -> datetime | None:
    """Extract single date from folder name"""
    folder_name = folder_path.name
    
    months = {
        'january': 1, 'february': 2, 'march': 3, 'april': 4,
        'may': 5, 'june': 6, 'july': 7, 'august': 8,
        'september': 9, 'october': 10, 'november': 11, 'december': 12
    }
    
    # Pattern: Single date (12 September 2025)
    single_date_pattern = r'(\d{1,2})\s+([a-zA-Z]+)\s+(\d{4})'
    match = re.search(single_date_pattern, folder_name)
    
    if match:
        day = int(match.group(1))
        month_name = match.group(2).lower()
        year = int(match.group(3))
        
        if month_name in months:
            try:
                return datetime(year, months[month_name], day, 12, 0, 0)
            except ValueError:
                pass
    
    return None


def extract_date_from_filename(filepath: Path, include_whatsapp: bool = False) -> list[tuple[datetime, str, str, bool]]:
    """
    Extract date from filename
    Args:
        filepath: Path to file
        include_whatsapp: If True, extract from WhatsApp files too (default: False)
    Returns list of tuples: (datetime, format_description, matched_text, has_time)
    """
    if not include_whatsapp and is_whatsapp_file(filepath):
        return []
    
    filename = filepath.stem
    found_dates = []
    matched_positions = set()
    
    # Patterns ordered by specificity (most specific first)
    # 1. First all patterns WITH TIME
    # 2. Then all patterns WITHOUT TIME
    # Within each group, more specific patterns come first
    
    patterns = [
        # ============================================================
        # PATTERNS WITH TIME (ordered by specificity)
        # ============================================================
        
        # WhatsApp Image 2023-04-09 at 6.26.53 PM.jpeg (very specific prefix + format)
        (r'WhatsApp\s+(?:Image|Video)\s+(\d{4})-(\d{2})-(\d{2})\s+at\s+(\d{1,2})\.(\d{2})\.(\d{2})\s+(AM|PM)',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)),
                           (int(m.group(4)) % 12) + (12 if m.group(7).upper() == 'PM' else 0),
                           int(m.group(5)), int(m.group(6))),
         "WhatsApp YYYY-MM-DD at H.MM.SS AM/PM", True),
        
        # IMG_20250730_002711_483 or IMG_20250730_002711 (with time)
        (r'IMG[-_](\d{4})(\d{2})(\d{2})[-_](\d{2})(\d{2})(\d{2})', 
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), 
                           int(m.group(4)), int(m.group(5)), int(m.group(6))),
         "IMG_YYYYMMDD_HHMMSS", True),
        
        # VID_20240329_143022 (with time)
        (r'VID[-_](\d{4})(\d{2})(\d{2})[-_](\d{2})(\d{2})(\d{2})', 
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), 
                           int(m.group(4)), int(m.group(5)), int(m.group(6))),
         "VID_YYYYMMDD_HHMMSS", True),
        
        # IMG20230902153011 (no separators, with time)
        (r'IMG(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})', 
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), 
                           int(m.group(4)), int(m.group(5)), int(m.group(6))),
         "IMGYYYYMMDDHHMMSS", True),
        
        # VID20230902153011 (no separators, with time)
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
        
        # 20230902_160253 or 20230902-160253 (date with time)
        (r'(\d{4})(\d{2})(\d{2})[-_](\d{2})(\d{2})(\d{2})', 
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), 
                           int(m.group(4)), int(m.group(5)), int(m.group(6))),
         "YYYYMMDD_HHMMSS", True),
        
        # 20230815143022 (full timestamp without separators, 14 digits)
        (r'(?<![A-Za-z])(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})(?!\d)', 
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), 
                           int(m.group(4)), int(m.group(5)), int(m.group(6))),
         "YYYYMMDDHHMMSS", True),
        
        # 2023-08-15_14-30-22
        (r'(\d{4})[-_](\d{2})[-_](\d{2})[-_](\d{2})[-_](\d{2})[-_](\d{2})', 
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), 
                           int(m.group(4)), int(m.group(5)), int(m.group(6))),
         "YYYY-MM-DD-HH-MM-SS", True),
        
        # ============================================================
        # PATTERNS WITHOUT TIME (ordered by specificity)
        # ============================================================
        
        # IMG-20230409-WA0016.jpg (WhatsApp - more specific, has -WA suffix)
        (r'IMG-(\d{4})(\d{2})(\d{2})-WA\d+',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), 12, 0, 0),
         "IMG-YYYYMMDD-WA (WhatsApp, no time)", False),
        
        # VID-20230409-WA0010.mp4 (WhatsApp - more specific, has -WA suffix)
        (r'VID-(\d{4})(\d{2})(\d{2})-WA\d+',
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), 12, 0, 0),
         "VID-YYYYMMDD-WA (WhatsApp, no time)", False),
        
        # IMG-20250522 or IMG_20250522 (date only, NO TIME)
        (r'IMG[-_](\d{4})(\d{2})(\d{2})(?![-_]?\d{2}[-_]?\d{2}[-_]?\d{2})', 
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), 12, 0, 0),
         "IMG_YYYYMMDD (no time)", False),
        
        # VID-20240329 (date only, no time)
        (r'VID[-_](\d{4})(\d{2})(\d{2})(?![-_]?\d{2}[-_]?\d{2}[-_]?\d{2})', 
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), 12, 0, 0),
         "VID_YYYYMMDD (no time)", False),
        
        # 2023-08-15 or 2023_08_15 (date only)
        (r'(\d{4})[-_](\d{2})[-_](\d{2})(?![-_]\d)', 
         lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), 12, 0, 0),
         "YYYY-MM-DD (no time)", False),
        
        # 20230815 (date only, 8 digits)
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
    
    # Remove duplicates, prefer versions with time
    unique_dates = []
    seen_dates = set()
    
    found_dates_sorted = sorted(found_dates, key=lambda x: (not x[3], x[0]))
    
    for date_info in found_dates_sorted:
        date_key = date_info[0].strftime("%Y-%m-%d")
        if date_key in seen_dates:
            continue
        seen_dates.add(date_key)
        unique_dates.append(date_info)
    
    return unique_dates


def parse_datetime_with_timezone(date_str: str) -> datetime | None:
    """
    Parse datetime string, handling UTC timezone conversion to local time.
    Works for both EXIF dates and video metadata dates.
    Converts UTC times (ending with 'Z' or '+00:00') to local time.
    """
    if not date_str or not date_str.strip():
        return None
    
    date_str = date_str.strip()
    
    try:
        from datetime import timezone as tz, timedelta
        
        # Check if it's a UTC timestamp
        is_utc = (
            date_str.endswith('Z') or 
            '+00:00' in date_str or 
            date_str.endswith('+0000') or
            '-00:00' in date_str or
            date_str.endswith('-0000')
        )
        
        # Check for other timezone offsets (e.g., +05:30, -08:00)
        has_tz_offset = False
        tz_offset_hours = 0
        tz_offset_minutes = 0
        clean_str = date_str
        
        # Remove milliseconds and 'Z' suffix
        clean_str = clean_str.replace('Z', '')
        if '.' in clean_str:
            clean_str = clean_str.split('.')[0]
        
        # Check for timezone offset at the end (e.g., +05:30, -08:00, +0530)
        tz_match = re.search(r'([+-])(\d{2}):?(\d{2})$', clean_str)
        if tz_match:
            has_tz_offset = True
            sign = 1 if tz_match.group(1) == '+' else -1
            tz_offset_hours = int(tz_match.group(2)) * sign
            tz_offset_minutes = int(tz_match.group(3)) * sign
            clean_str = clean_str[:tz_match.start()]
        
        # Try different parsing formats
        parsed = None
        
        # ISO format: 2023-09-02T14:35:43
        try:
            parsed = datetime.fromisoformat(clean_str)
        except ValueError:
            pass
        
        # EXIF format: 2023:09:02 14:35:43
        if not parsed:
            try:
                parsed = datetime.strptime(clean_str, '%Y:%m:%d %H:%M:%S')
            except ValueError:
                pass
        
        # Other formats
        if not parsed:
            formats = [
                '%Y-%m-%d %H:%M:%S',
                '%Y/%m/%d %H:%M:%S',
                '%Y-%m-%dT%H:%M:%S',
                '%Y:%m:%d %H:%M',
                '%Y-%m-%d %H:%M',
                '%Y:%m:%d',
                '%Y-%m-%d',
            ]
            for fmt in formats:
                try:
                    parsed = datetime.strptime(clean_str, fmt)
                    break
                except ValueError:
                    continue
        
        if not parsed:
            return None
        
        # Handle timezone conversion
        if is_utc:
            # UTC time - convert to local
            utc_dt = parsed.replace(tzinfo=tz.utc)
            local_dt = utc_dt.astimezone()
            print(f"    ✓ Converted UTC {parsed.strftime('%Y-%m-%d %H:%M:%S')} to local {local_dt.strftime('%Y-%m-%d %H:%M:%S')}")
            return local_dt.replace(tzinfo=None)
        elif has_tz_offset:
            # Has timezone offset - convert to local
            offset = timedelta(hours=tz_offset_hours, minutes=tz_offset_minutes)
            source_tz = tz(offset)
            source_dt = parsed.replace(tzinfo=source_tz)
            local_dt = source_dt.astimezone()
            if abs(offset.total_seconds()) > 0:  # Only print if not already local
                local_tz = local_dt.tzinfo
                if source_tz != local_tz:
                    print(f"    ✓ Converted {parsed.strftime('%Y-%m-%d %H:%M:%S')} (UTC{tz_match.group(0)}) to local {local_dt.strftime('%Y-%m-%d %H:%M:%S')}")
            return local_dt.replace(tzinfo=None)
        else:
            # No timezone info - assume already local time
            return parsed
        
    except Exception as e:
        print(f"    ! Error parsing datetime '{date_str}': {e}")
        return None


def parse_exif_date(date_str: str) -> datetime | None:
    """Parse EXIF date string, handling timezone conversion if present."""
    if not date_str or not date_str.strip():
        return None
    
    date_str = date_str.strip()
    
    # Check if timezone info is present - use timezone-aware parser
    has_tz = 'Z' in date_str or (len(date_str) > 10 and '+' in date_str[10:]) or (date_str.count('-') > 2 and len(date_str) > 19)
    if has_tz:
        result = parse_datetime_with_timezone(date_str)
        if result:
            return result
    
    # Standard EXIF format (no timezone - assume local time)
    try:
        return datetime.strptime(date_str, '%Y:%m:%d %H:%M:%S')
    except ValueError:
        pass
    
    # Try other formats
    formats = [
        '%Y-%m-%d %H:%M:%S',
        '%Y/%m/%d %H:%M:%S',
        '%Y-%m-%dT%H:%M:%S',
        '%Y:%m:%d %H:%M',
        '%Y-%m-%d %H:%M',
        '%Y:%m:%d',
        '%Y-%m-%d',
    ]
    
    for fmt in formats:
        try:
            parsed = datetime.strptime(date_str, fmt)
            # If no time component in format, default to noon
            if parsed.hour == 0 and parsed.minute == 0 and parsed.second == 0:
                if '%H' not in fmt and '%I' not in fmt:
                    parsed = parsed.replace(hour=12)
            return parsed
        except ValueError:
            continue
    
    # Last resort - try timezone-aware parser
    return parse_datetime_with_timezone(date_str)


def get_windows_date_taken(filepath: Path) -> datetime | None:
    """Get Windows 'Date Taken' property using Shell"""
    if not has_win32:
        return None
    
    try:
        sh = win32com.client.Dispatch("Shell.Application")
        folder = sh.Namespace(str(filepath.parent))
        item = folder.ParseName(filepath.name)
        
        if not item:
            return None
        
        # Property index 12 = Date Taken
        date_taken = folder.GetDetailsOf(item, 12)
        if date_taken and date_taken.strip():
            # Clean the string
            date_str = date_taken.strip()
            date_str = date_str.encode('ascii', 'ignore').decode('ascii')
            date_str = ' '.join(date_str.split())
            
            # Try various formats
            formats = [
                '%d/%m/%Y %H:%M',
                '%d-%m-%Y %H:%M',
                '%m/%d/%Y %H:%M',
                '%m-%d-%Y %H:%M',
                '%Y-%m-%d %H:%M',
                '%d/%m/%Y %I:%M %p',
                '%d-%m-%Y %I:%M %p',
                '%m/%d/%Y %I:%M %p',
                '%d/%m/%Y',
                '%m/%d/%Y',
                '%Y-%m-%d',
            ]
            
            for fmt in formats:
                try:
                    parsed = datetime.strptime(date_str, fmt)
                    print(f"    ✓ Windows Date Taken parsed: {parsed}")
                    return parsed
                except ValueError:
                    continue
            
            print(f"    ✗ Could not parse Windows Date Taken: '{date_str}'")
        
    except Exception as e:
        print(f"  ! Error reading Windows Date Taken: {e}")
    
    return None


def get_windows_media_created(filepath: Path) -> datetime | None:
    """Get Windows 'Media Created' property using Shell"""
    if not has_win32:
        return None
    
    try:
        sh = win32com.client.Dispatch("Shell.Application")
        folder = sh.Namespace(str(filepath.parent))
        item = folder.ParseName(filepath.name)
        
        if not item:
            return None
        
        # Property index 208 = Media Created
        media_created = folder.GetDetailsOf(item, 208)
        if media_created and media_created.strip():
            date_str = media_created.strip()
            date_str = date_str.encode('ascii', 'ignore').decode('ascii')
            date_str = ' '.join(date_str.split())
            
            formats = [
                '%d/%m/%Y %H:%M',
                '%d-%m-%Y %H:%M',
                '%m/%d/%Y %H:%M',
                '%m-%d-%Y %H:%M',
                '%Y-%m-%d %H:%M',
                '%d/%m/%Y %I:%M %p',
                '%m/%d/%Y %I:%M %p',
                '%d/%m/%Y',
                '%m/%d/%Y',
                '%Y-%m-%d',
            ]
            
            for fmt in formats:
                try:
                    parsed = datetime.strptime(date_str, fmt)
                    print(f"    ✓ Windows Media Created parsed: {parsed}")
                    return parsed
                except ValueError:
                    continue
            
            print(f"    ✗ Could not parse Windows Media Created: '{date_str}'")
        
    except Exception as e:
        print(f"  ! Error reading Windows Media Created: {e}")
    
    return None


def get_image_datetime(filepath: Path) -> tuple[datetime | None, str | None, str | None]:
    """
    Extract datetime from regular image files (non-RAW)
    Priority:
        1. Windows Date Taken
        2. DateTimeOriginal (Exif.Photo.DateTimeOriginal 0x9003)
        3. DateTime (Exif.Image.DateTime 0x0132)
        4. DateTimeDigitized (Exif.Photo.DateTimeDigitized 0x9004)
    Returns: (datetime, unparsed_string, source_name)
    """
    # Priority 1: Windows Date Taken
    print(f"  → Trying Windows Date Taken...")
    win_date = get_windows_date_taken(filepath)
    if win_date:
        return win_date, None, "Windows Date Taken"
    
    if not check_exiftool_available():
        print(f"  → exiftool not available, skipping EXIF extraction")
        return None, None, None
    
    try:
        # Note: exiftool tag names:
        # - DateTimeOriginal = Exif.Photo.DateTimeOriginal (0x9003)
        # - ModifyDate = Exif.Image.DateTime (0x0132)
        # - CreateDate = Exif.Photo.DateTimeDigitized (0x9004)
        # Also get timezone offset tags if available
        cmd = [
            'exiftool', '-json', '-m',
            '-DateTimeOriginal', '-ModifyDate', '-CreateDate',
            '-OffsetTimeOriginal', '-OffsetTime', '-OffsetTimeDigitized',
            str(filepath)
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        
        if result.returncode == 0:
            data = json.loads(result.stdout)
            if data and len(data) > 0:
                tags = data[0]
                
                # Get timezone offsets if available
                offset_original = tags.get('OffsetTimeOriginal', '')
                offset_modify = tags.get('OffsetTime', '')
                offset_digitized = tags.get('OffsetTimeDigitized', '')
                
                # Priority 2: DateTimeOriginal (Exif.Photo.DateTimeOriginal 0x9003)
                if 'DateTimeOriginal' in tags and tags['DateTimeOriginal']:
                    value = str(tags['DateTimeOriginal'])
                    # Append timezone offset if available
                    if offset_original:
                        value = value + offset_original
                    print(f"  → exiftool DateTimeOriginal (0x9003): '{value}'")
                    parsed = parse_exif_date(value)
                    if parsed:
                        return parsed, None, "DateTimeOriginal (0x9003)"
                    else:
                        return None, value, "DateTimeOriginal (0x9003) (unparsed)"
                
                # Priority 3: ModifyDate (Exif.Image.DateTime 0x0132)
                if 'ModifyDate' in tags and tags['ModifyDate']:
                    value = str(tags['ModifyDate'])
                    if offset_modify:
                        value = value + offset_modify
                    print(f"  → exiftool DateTime (0x0132): '{value}'")
                    parsed = parse_exif_date(value)
                    if parsed:
                        return parsed, None, "DateTime (0x0132)"
                    else:
                        return None, value, "DateTime (0x0132) (unparsed)"
                
                # Priority 4: CreateDate (Exif.Photo.DateTimeDigitized 0x9004)
                if 'CreateDate' in tags and tags['CreateDate']:
                    value = str(tags['CreateDate'])
                    if offset_digitized:
                        value = value + offset_digitized
                    print(f"  → exiftool DateTimeDigitized (0x9004): '{value}'")
                    parsed = parse_exif_date(value)
                    if parsed:
                        return parsed, None, "DateTimeDigitized (0x9004)"
                    else:
                        return None, value, "DateTimeDigitized (0x9004) (unparsed)"
    
    except subprocess.TimeoutExpired:
        print(f"  ! exiftool timed out")
    except json.JSONDecodeError as e:
        print(f"  ! exiftool JSON error: {e}")
    except Exception as e:
        print(f"  ! exiftool error: {e}")
    
    return None, None, None


def get_raw_datetime(filepath: Path) -> tuple[datetime | None, str | None, str | None]:
    """
    Extract datetime from RAW image files
    Priority:
        1. Windows Date Taken
        2. DateTimeOriginal (exiftool)
        3. CreateDate (exiftool)
    Returns: (datetime, unparsed_string, source_name)
    """
    # Priority 1: Windows Date Taken
    print(f"  → Trying Windows Date Taken...")
    win_date = get_windows_date_taken(filepath)
    if win_date:
        return win_date, None, "Windows Date Taken"
    
    if not check_exiftool_available():
        print(f"  → exiftool not available, skipping EXIF extraction")
        return None, None, None
    
    try:
        # Also get timezone offset tags if available
        cmd = [
            'exiftool', '-json', '-m',
            '-DateTimeOriginal', '-CreateDate',
            '-OffsetTimeOriginal', '-OffsetTimeDigitized',
            str(filepath)
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        
        if result.returncode == 0:
            data = json.loads(result.stdout)
            if data and len(data) > 0:
                tags = data[0]
                
                # Get timezone offsets if available
                offset_original = tags.get('OffsetTimeOriginal', '')
                offset_digitized = tags.get('OffsetTimeDigitized', '')
                
                # Priority 2: DateTimeOriginal
                if 'DateTimeOriginal' in tags and tags['DateTimeOriginal']:
                    value = str(tags['DateTimeOriginal'])
                    if offset_original:
                        value = value + offset_original
                    print(f"  → exiftool DateTimeOriginal: '{value}'")
                    parsed = parse_exif_date(value)
                    if parsed:
                        return parsed, None, "DateTimeOriginal"
                    else:
                        return None, value, "DateTimeOriginal (unparsed)"
                
                # Priority 3: CreateDate
                if 'CreateDate' in tags and tags['CreateDate']:
                    value = str(tags['CreateDate'])
                    if offset_digitized:
                        value = value + offset_digitized
                    print(f"  → exiftool CreateDate: '{value}'")
                    parsed = parse_exif_date(value)
                    if parsed:
                        return parsed, None, "CreateDate"
                    else:
                        return None, value, "CreateDate (unparsed)"
    
    except subprocess.TimeoutExpired:
        print(f"  ! exiftool timed out")
    except json.JSONDecodeError as e:
        print(f"  ! exiftool JSON error: {e}")
    except Exception as e:
        print(f"  ! exiftool error: {e}")
    
    return None, None, None


def get_video_datetime(filepath: Path) -> tuple[datetime | None, str | None, str | None]:
    """
    Extract datetime from video file
    Priority:
        1. Windows Media Created
        2. Windows Date Taken
        3. creation_time (ffprobe)
        4. date / date_recorded (ffprobe)
    Returns: (datetime, unparsed_string, source_name)
    """
    # Priority 1: Windows Media Created
    print(f"  → Trying Windows Media Created...")
    win_date = get_windows_media_created(filepath)
    if win_date:
        return win_date, None, "Windows Media Created"
    
    # Priority 2: Windows Date Taken
    print(f"  → Trying Windows Date Taken...")
    win_date = get_windows_date_taken(filepath)
    if win_date:
        return win_date, None, "Windows Date Taken"
    
    # Try ffprobe for priorities 3 and 4
    if check_ffprobe_available():
        try:
            cmd = [
                'ffprobe', '-v', 'quiet', '-print_format', 'json',
                '-show_entries', 'format_tags=creation_time,date,date_recorded',
                str(filepath)
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0:
                data = json.loads(result.stdout)
                if 'format' in data and 'tags' in data['format']:
                    tags = data['format']['tags']
                    
                    # Priority 3: creation_time
                    if 'creation_time' in tags and tags['creation_time']:
                        value = tags['creation_time']
                        print(f"  → ffprobe creation_time: '{value}'")
                        parsed = parse_datetime_with_timezone(value)
                        if parsed:
                            return parsed, None, "creation_time"
                        else:
                            return None, value, "creation_time (unparsed)"
                    
                    # Priority 4: date
                    if 'date' in tags and tags['date']:
                        value = tags['date']
                        print(f"  → ffprobe date: '{value}'")
                        parsed = parse_datetime_with_timezone(value)
                        if parsed:
                            return parsed, None, "date"
                        else:
                            return None, value, "date (unparsed)"
                    
                    # Priority 4 (alternate): date_recorded
                    if 'date_recorded' in tags and tags['date_recorded']:
                        value = tags['date_recorded']
                        print(f"  → ffprobe date_recorded: '{value}'")
                        parsed = parse_datetime_with_timezone(value)
                        if parsed:
                            return parsed, None, "date_recorded"
                        else:
                            return None, value, "date_recorded (unparsed)"
        
        except subprocess.TimeoutExpired:
            print(f"  ! ffprobe timed out")
        except Exception as e:
            print(f"  ! ffprobe error: {e}")
    
    return None, None, None


def ask_user_for_date_choice(filepath: Path, options: list[tuple[datetime, str]]) -> tuple[datetime, str]:
    """Ask user to choose from multiple date options - no skip option"""
    print(f"\n❓ Multiple date options for: {filepath.name}")
    print("  Please choose which date to use:")
    
    for i, (date, source) in enumerate(options, 1):
        if is_date_valid(date):
            print(f"  {i}. {date.strftime('%Y-%m-%d %H:%M:%S')} (from {source})")
        else:
            print(f"  {i}. {date.strftime('%Y-%m-%d %H:%M:%S')} (from {source}) - ⚠️ OUT OF VALID RANGE")
    
    print(f"  {len(options) + 1}. Enter date manually")
    
    while True:
        choice = input(f"  Choose (1-{len(options) + 1}): ").strip()
        try:
            choice_num = int(choice)
            if 1 <= choice_num <= len(options):
                chosen_date, chosen_source = options[choice_num - 1]
                if not is_date_valid(chosen_date):
                    print(f"  ⚠️ Selected date {chosen_date} is out of valid range!")
                    confirm = input("  Do you want to use this date anyway? (y/n): ").strip().lower()
                    if confirm != 'y':
                        continue
                return chosen_date, chosen_source
            elif choice_num == len(options) + 1:
                manual_date = ask_for_manual_date(filepath)
                return manual_date, "manual entry"
        except ValueError:
            pass
        print(f"  Invalid choice. Please enter 1-{len(options) + 1}.")


def ask_for_manual_date(filepath: Path) -> datetime:
    """Ask user to manually enter a date for a file - keeps asking until valid"""
    print("  Enter date manually")
    print("  Formats: DD-MM-YYYY, DD/MM/YYYY, DD-MM-YYYY HH:MM:SS, DD Month YYYY")
    print("  Examples: 25-12-2023, 25/12/2023 14:30:00, 25 December 2023")
    print(f"  Valid range: {MIN_DATE.date()} to {MAX_DATE.date()}")
    
    while True:
        user_input = input("  Date: ").strip()
        
        parsed_date = parse_user_date_input(user_input)
        if parsed_date:
            if not is_date_valid(parsed_date):
                print(f"  ⚠️ Date {parsed_date} is out of valid range!")
                confirm = input("  Use it anyway? (y/n): ").strip().lower()
                if confirm != 'y':
                    continue
            return parsed_date
        else:
            print("  ❌ Invalid date format. Please try again.")


def parse_user_date_input(date_str: str) -> datetime | None:
    """Parse user-entered date string in various formats"""
    date_str = date_str.strip()
    
    formats = [
        '%d-%m-%Y %H:%M:%S',
        '%d/%m/%Y %H:%M:%S',
        '%d-%m-%Y %H:%M',
        '%d/%m/%Y %H:%M',
        '%d-%m-%Y %I:%M %p',
        '%d/%m/%Y %I:%M %p',
        '%d-%m-%Y %I:%M:%S %p',
        '%d/%m/%Y %I:%M:%S %p',
        '%d-%m-%Y',
        '%d/%m/%Y',
        '%d %B %Y',
        '%d %b %Y',
        '%d %B %Y %H:%M:%S',
        '%d %B %Y %H:%M',
        '%Y-%m-%d %H:%M:%S',
        '%Y-%m-%d',
    ]
    
    for fmt in formats:
        try:
            parsed_date = datetime.strptime(date_str, fmt)
            if parsed_date.hour == 0 and parsed_date.minute == 0 and parsed_date.second == 0:
                if '%H' not in fmt and '%I' not in fmt:
                    parsed_date = parsed_date.replace(hour=12)
            return parsed_date
        except ValueError:
            continue
    
    return None


def update_image_metadata(filepath: Path, new_date: datetime) -> bool:
    """
    Update metadata in image files using exiftool
    Updates: DateTimeOriginal, CreateDate, ModifyDate
    Works with: JPEG, PNG, HEIC, TIFF, and most image formats
    """
    if not check_exiftool_available():
        print("  ⚠ exiftool not available - cannot update image metadata")
        return False
    
    try:
        exif_date = new_date.strftime("%Y:%m:%d %H:%M:%S")
        
        cmd = [
            'exiftool',
            '-m',  # Ignore minor errors
            '-overwrite_original',
            f'-DateTimeOriginal={exif_date}',
            f'-CreateDate={exif_date}',
            f'-ModifyDate={exif_date}',
            str(filepath)
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        
        if result.returncode == 0:
            # Check if actually updated
            if "1 image files updated" in result.stdout:
                print(f"  → Updated image metadata: DateTimeOriginal, CreateDate, ModifyDate = {new_date}")
                return True
            elif "0 image files updated" in result.stdout:
                print(f"  ⚠ exiftool reported 0 files updated - metadata may not have been written")
                return False
            else:
                print(f"  → Updated image metadata: DateTimeOriginal, CreateDate, ModifyDate = {new_date}")
                return True
        else:
            error_msg = result.stderr.strip() if result.stderr else result.stdout.strip()
            
            # Check for specific errors
            if "Not a valid" in error_msg:
                print(f"  ⚠ File format mismatch detected: {error_msg[:100]}")
                print(f"  → Try option to fix file extension mismatch")
            else:
                print(f"  ⚠ exiftool failed: {error_msg[:100]}")
            return False
            
    except subprocess.TimeoutExpired:
        print(f"  ⚠ exiftool timed out")
        return False
    except Exception as e:
        print(f"  ⚠ Error updating image metadata: {e}")
        return False


def update_raw_metadata(filepath: Path, new_date: datetime) -> bool:
    """
    Update metadata in RAW image files using exiftool
    Updates: DateTimeOriginal, CreateDate
    """
    if not check_exiftool_available():
        print("  ⚠ exiftool not available - cannot update RAW metadata")
        return False
    
    try:
        exif_date = new_date.strftime("%Y:%m:%d %H:%M:%S")
        
        cmd = [
            'exiftool',
            '-m',  # Ignore minor errors
            '-overwrite_original',
            f'-DateTimeOriginal={exif_date}',
            f'-CreateDate={exif_date}',
            str(filepath)
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        
        if result.returncode == 0:
            if "0 image files updated" in result.stdout:
                print(f"  ⚠ exiftool reported 0 files updated - metadata may not have been written")
                return False
            print(f"  → Updated RAW metadata: DateTimeOriginal, CreateDate = {new_date}")
            return True
        else:
            error_msg = result.stderr.strip() if result.stderr else result.stdout.strip()
            if "Not a valid" in error_msg:
                print(f"  ⚠ File format mismatch detected: {error_msg[:100]}")
            else:
                print(f"  ⚠ exiftool failed: {error_msg[:100]}")
            return False
            
    except subprocess.TimeoutExpired:
        print(f"  ⚠ exiftool timed out")
        return False
    except Exception as e:
        print(f"  ⚠ Error updating RAW metadata: {e}")
        return False


def update_video_metadata(filepath: Path, new_date: datetime) -> bool:
    """
    Update metadata in video files using ffmpeg
    Updates: creation_time, date, date_recorded
    Note: Converts local time to UTC for storage (industry standard)
    """
    if not check_ffmpeg_available():
        print("  ⚠ ffmpeg not available - cannot update video metadata")
        return False
    
    try:
        temp_file = filepath.parent / f"{filepath.stem}_temp{filepath.suffix}"
        
        # Convert local time to UTC for storage
        # This is the industry standard - devices expect UTC and convert to local for display
        from datetime import timezone as tz
        
        # Treat input as local time and convert to UTC
        local_tz = datetime.now().astimezone().tzinfo
        local_dt = new_date.replace(tzinfo=local_tz)
        utc_dt = local_dt.astimezone(tz.utc)
        
        # Format as UTC with Z suffix
        iso_date = utc_dt.strftime("%Y-%m-%dT%H:%M:%S.000000Z")
        
        print(f"  → Converting local {new_date.strftime('%Y-%m-%d %H:%M:%S')} to UTC {utc_dt.strftime('%Y-%m-%d %H:%M:%S')}Z for storage")
        
        cmd = [
            'ffmpeg',
            '-i', str(filepath),
            '-map', '0',
            '-c', 'copy',
            '-metadata', f'creation_time={iso_date}',
            '-metadata', f'date={iso_date}',
            '-metadata', f'date_recorded={iso_date}',
            '-y',
            '-loglevel', 'error',
            str(temp_file)
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        
        if result.returncode == 0 and temp_file.exists() and temp_file.stat().st_size > 0:
            try:
                filepath.unlink()
                temp_file.rename(filepath)
                print(f"  → Updated video metadata: creation_time, date, date_recorded = {new_date} (stored as UTC)")
                return True
            except Exception as e:
                print(f"  ⚠ Error replacing file: {e}")
                if temp_file.exists():
                    temp_file.unlink()
                return False
        else:
            if temp_file.exists():
                temp_file.unlink()
            if result.stderr:
                error_msg = result.stderr.strip()[:100]
                print(f"  ⚠ ffmpeg failed: {error_msg}")
            return False
            
    except subprocess.TimeoutExpired:
        print(f"  ⚠ ffmpeg timed out")
        if 'temp_file' in locals() and temp_file.exists():
            temp_file.unlink()
        return False
    except Exception as e:
        if 'temp_file' in locals() and temp_file.exists():
            temp_file.unlink()
        print(f"  ⚠ Error updating video metadata: {e}")
        return False


def set_windows_date_taken(filepath: Path, new_date: datetime) -> bool:
    """
    Set Windows 'Date Taken' property using Property System
    This updates the property that Windows Explorer shows as 'Date Taken'
    """
    if not has_win32:
        return False
    
    try:
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Get property store for the file
        ps = propsys.SHGetPropertyStoreFromParsingName(
            str(filepath), 
            None, 
            shellcon.GPS_READWRITE, 
            propsys.IID_IPropertyStore
        )
        
        # System.Photo.DateTaken property key
        PKEY_Photo_DateTaken = propsys.PSGetPropertyKeyFromName("System.Photo.DateTaken")
        
        # Convert datetime to Windows FILETIME format via pywintypes
        import pywintypes
        pytime = pywintypes.Time(new_date)
        
        # Create PROPVARIANT with the date
        pv = propsys.PROPVARIANTType(pytime)
        
        # Set the property
        ps.SetValue(PKEY_Photo_DateTaken, pv)
        ps.Commit()
        
        print(f"  → Updated Windows Date Taken = {new_date}")
        return True
        
    except Exception as e:
        # This is expected to fail for some file types that don't support Date Taken
        # Silently continue - the metadata update should have handled it
        pass
    
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass
    
    return False


def set_file_system_dates(filepath: Path, new_date: datetime) -> bool:
    """Set file system modification date (Windows Date Modified)"""
    try:
        timestamp = new_date.timestamp()
        os.utime(filepath, (timestamp, timestamp))
        print(f"  → Updated file system Date Modified = {new_date}")
        return True
    except Exception as e:
        print(f"  ✗ Error setting file system date: {e}")
        return False


def process_folder(folder_path: str) -> None:
    """Process all media files in folder and subfolders"""
    folder = Path(folder_path)
    
    all_extensions = IMAGE_EXTENSIONS | RAW_EXTENSIONS | VIDEO_EXTENSIONS
    
    processed_folder = 0
    processed_metadata = 0
    processed_filename = 0
    processed_manual = 0
    metadata_updated = 0
    windows_date_taken_updated = 0
    skipped_files = 0
    error_files = []
    
    print(f"Processing media files in: {folder_path}\n")
    
    # Check dependencies
    exiftool_available = check_exiftool_available()
    ffmpeg_available = check_ffmpeg_available()
    ffprobe_available = check_ffprobe_available()
    
    if not exiftool_available:
        print("⚠️  WARNING: exiftool not installed!")
        print("    Image/RAW metadata cannot be read or updated.")
        print("    Install from: https://exiftool.org/")
        print()
    
    if not ffmpeg_available or not ffprobe_available:
        print("⚠️  WARNING: ffmpeg/ffprobe not installed!")
        print("    Video metadata cannot be read or updated.")
        print()
    
    if sys.platform == 'win32' and not has_win32:
        print("⚠️  WARNING: pywin32 not installed!")
        print("    Windows Date Taken property cannot be read or updated.")
        print("    Install with: pip install pywin32")
        print()
    
    print(f"Force metadata update: {'ENABLED' if FORCE_METADATA_UPDATE else 'DISABLED'}")
    print()
    
    all_files = list(folder.rglob('*'))
    media_files = sorted(
        [f for f in all_files if f.is_file() and f.suffix.lower() in all_extensions],
        key=lambda x: (x.parent, x.name.lower())
    )
    
    print(f"Found {len(media_files)} media files to process\n")
    print(f"{'='*60}\n")
    
    for i, filepath in enumerate(media_files, 1):
        file_type = get_file_type_name(filepath)
        print(f"[{i}/{len(media_files)}] Processing ({file_type}): {filepath.name}")
        print(f"  Path: {filepath.parent}")
        
        # Check for file type mismatch (images and RAW only)
        if is_image_file(filepath) or is_raw_file(filepath):
            fixed_path = fix_file_extension_mismatch(filepath)
            if fixed_path is None:
                print(f"  → Skipping file (user choice)")
                skipped_files += 1
                print()
                continue
            elif fixed_path != filepath:
                filepath = fixed_path
                file_type = get_file_type_name(filepath)
                print(f"  → Now processing as ({file_type}): {filepath.name}")
        
        date_found = None
        source = None
        update_metadata = False
        
        # Step 1: Extract folder date
        folder_date = extract_date_from_folder(filepath.parent)
        if folder_date:
            print(f"  → Folder date: {folder_date.date()}")
        
        # Step 2: Extract metadata datetime based on file type
        metadata_date = None
        metadata_unparsed = None
        metadata_source = None
        
        if is_raw_file(filepath):
            metadata_date, metadata_unparsed, metadata_source = get_raw_datetime(filepath)
        elif is_image_file(filepath):
            metadata_date, metadata_unparsed, metadata_source = get_image_datetime(filepath)
        elif is_video_file(filepath):
            metadata_date, metadata_unparsed, metadata_source = get_video_datetime(filepath)
        
        if metadata_date:
            print(f"  → Metadata datetime: {metadata_date} (from {metadata_source})")
        elif metadata_unparsed:
            print(f"  → Metadata found but unparsed: '{metadata_unparsed}' (from {metadata_source})")
        else:
            print(f"  → No metadata datetime found")
        
        # Step 3: Extract filename datetime
        # For WhatsApp files, extract but mark as unreliable
        is_whatsapp = is_whatsapp_file(filepath)
        filename_dates = extract_date_from_filename(filepath, include_whatsapp=True)
        filename_date = None
        filename_has_time = False
        filename_match = None
        
        if filename_dates:
            filename_date, filename_format, filename_match, filename_has_time = filename_dates[0]
            time_info = "with time" if filename_has_time else "no time"
            if is_whatsapp:
                print(f"  → Filename datetime: {filename_date} ({time_info}, WhatsApp ⚠️ unreliable)")
            else:
                print(f"  → Filename datetime: {filename_date} ({time_info}, pattern: {filename_format})")
        
        # Helper function to get filename label with WhatsApp warning if needed
        def get_filename_label(prefix: str = "Filename") -> str:
            if is_whatsapp:
                return f"{prefix}: {filename_match} ⚠️ WhatsApp (unreliable)"
            return f"{prefix}: {filename_match}"
        
        # Step 4: Handle unparsed metadata - ask user to interpret
        if metadata_unparsed and not metadata_date:
            print(f"  → Metadata exists but couldn't be parsed automatically")
            print(f"  → Please enter the date manually based on: '{metadata_unparsed}'")
            metadata_date = ask_for_manual_date(filepath)
            metadata_source = "manual interpretation of metadata"
            update_metadata = True
        
        # Step 5: Decision logic
        date_from_user_choice = False
        original_metadata_date = metadata_date
        
        # Case A: Folder date EXISTS and Metadata datetime EXISTS
        if folder_date and metadata_date:
            if folder_date.date() == metadata_date.date():
                date_found = metadata_date
                source = f"{metadata_source} (matches folder date)"
                print(f"  → Same date in folder and metadata - using metadata datetime")
            else:
                print(f"  → Different dates: folder={folder_date.date()}, metadata={metadata_date.date()}")
                options = []
                
                combined = folder_date.replace(
                    hour=metadata_date.hour,
                    minute=metadata_date.minute,
                    second=metadata_date.second
                )
                options.append((combined, f"Folder date ({folder_date.date()}) + metadata time ({metadata_date.strftime('%H:%M:%S')})"))
                options.append((metadata_date, f"Metadata: {metadata_source}"))
                
                if filename_date:
                    options.append((filename_date, get_filename_label()))
                
                date_found, source = ask_user_for_date_choice(filepath, options)
                date_from_user_choice = True
        
        # Case B: Folder date EXISTS but Metadata DOES NOT EXIST
        elif folder_date and not metadata_date:
            options = []
            
            if filename_date and filename_has_time:
                combined = folder_date.replace(
                    hour=filename_date.hour,
                    minute=filename_date.minute,
                    second=filename_date.second
                )
                warn = " ⚠️ WhatsApp" if is_whatsapp else ""
                options.append((combined, f"Folder date ({folder_date.date()}) + filename time ({filename_date.strftime('%H:%M:%S')}){warn}"))
            
            if filename_date:
                options.append((filename_date, get_filename_label("Filename datetime")))
            
            options.append((folder_date, f"Folder date only ({folder_date.date()}, 12:00:00)"))
            
            date_found, source = ask_user_for_date_choice(filepath, options)
            date_from_user_choice = True
        
        # Case C: No folder date, but Metadata AND Filename BOTH exist
        elif not folder_date and metadata_date and filename_date:
            if metadata_date.date() == filename_date.date():
                date_found = metadata_date
                source = f"{metadata_source} (matches filename date)"
                print(f"  → Same date in metadata and filename - using metadata datetime")
            else:
                print(f"  → Different dates: metadata={metadata_date.date()}, filename={filename_date.date()}")
                options = [
                    (metadata_date, f"Metadata: {metadata_source}"),
                    (filename_date, get_filename_label())
                ]
                date_found, source = ask_user_for_date_choice(filepath, options)
                date_from_user_choice = True
        
        # Case D: No folder date, Metadata EXISTS, but NO filename date
        elif not folder_date and metadata_date and not filename_date:
            print(f"  → Only metadata datetime found - asking for confirmation")
            options = [
                (metadata_date, f"Metadata: {metadata_source} (⚠️ cannot verify, no other source)")
            ]
            date_found, source = ask_user_for_date_choice(filepath, options)
            date_from_user_choice = True
        
        # Case E: ONLY Filename datetime exists (no folder, no metadata)
        elif not folder_date and not metadata_date and filename_date:
            print(f"  → Only filename datetime found - asking for confirmation")
            options = [
                (filename_date, get_filename_label())
            ]
            date_found, source = ask_user_for_date_choice(filepath, options)
            date_from_user_choice = True
        
        # Case F: NOTHING found anywhere
        else:
            print(f"  → No date found anywhere - manual entry required")
            date_found = ask_for_manual_date(filepath)
            date_from_user_choice = True
            source = "manual entry"
        
        # Step 6: Validate date ONLY for auto-selected dates (not user choices)
        if date_found and not is_date_valid(date_found) and not date_from_user_choice:
            print(f"  ⚠️ Date {date_found} is out of valid range ({MIN_DATE.date()} to {MAX_DATE.date()})!")
            confirm = input("  Use this date anyway? (y/n): ").strip().lower()
            if confirm != 'y':
                options = []
                if folder_date:
                    if metadata_date:
                        combined = folder_date.replace(
                            hour=metadata_date.hour,
                            minute=metadata_date.minute,
                            second=metadata_date.second
                        )
                        options.append((combined, f"Folder date + metadata time"))
                    elif filename_date and filename_has_time:
                        combined = folder_date.replace(
                            hour=filename_date.hour,
                            minute=filename_date.minute,
                            second=filename_date.second
                        )
                        options.append((combined, f"Folder date + filename time"))
                    else:
                        options.append((folder_date, f"Folder date: {folder_date.date()}"))
                
                if metadata_date:
                    options.append((metadata_date, f"Metadata: {metadata_source}"))
                
                if filename_date:
                    options.append((filename_date, get_filename_label()))
                
                if options:
                    date_found, source = ask_user_for_date_choice(filepath, options)
                else:
                    date_found = ask_for_manual_date(filepath)
                    source = "manual entry (after out-of-range)"
                date_from_user_choice = True
        
        # Step 7: Determine if metadata update is needed
        if date_found:
            if original_metadata_date is None:
                update_metadata = True
                print(f"  → No existing metadata - will create metadata")
            elif date_found != original_metadata_date:
                update_metadata = True
                print(f"  → Date differs from metadata ({original_metadata_date} → {date_found}) - will update")
            elif FORCE_METADATA_UPDATE:
                update_metadata = True
                print(f"  → Force update enabled - will update metadata")
            else:
                update_metadata = False
                print(f"  → Metadata already correct - skipping metadata update")
        
        # Step 8: Update metadata if needed
        if date_found and update_metadata:
            print(f"  → Updating metadata to: {date_found}")
            update_success = False
            
            if is_raw_file(filepath):
                update_success = update_raw_metadata(filepath, date_found)
            elif is_image_file(filepath):
                update_success = update_image_metadata(filepath, date_found)
            elif is_video_file(filepath):
                update_success = update_video_metadata(filepath, date_found)
            else:
                print(f"  ⚠ Unknown file type - cannot update metadata")
            
            if update_success:
                metadata_updated += 1
            else:
                error_files.append((str(filepath), "Failed to update metadata"))
        
        # Step 9: Update Windows Date Taken property
        if date_found and has_win32:
            if set_windows_date_taken(filepath, date_found):
                windows_date_taken_updated += 1
        
        # Step 10: Update file system Date Modified
        if date_found:
            if set_file_system_dates(filepath, date_found):
                print(f"  ✓ File date updated to {date_found} (from {source})")
                
                if "Folder" in source:
                    processed_folder += 1
                elif "Metadata" in source or "metadata" in source:
                    processed_metadata += 1
                elif "Filename" in source or "filename" in source:
                    processed_filename += 1
                elif "manual" in source:
                    processed_manual += 1
            else:
                print(f"  ✗ Failed to set file system date")
                error_files.append((str(filepath), f"Failed to set file system date"))
        
        print()
    
    # Summary
    print(f"\n{'='*60}")
    print(f"SUMMARY:")
    print(f"{'='*60}")
    print(f"Files updated from folder name: {processed_folder}")
    print(f"Files with existing metadata: {processed_metadata}")
    print(f"Files updated from filename: {processed_filename}")
    print(f"Files updated from manual entry: {processed_manual}")
    print(f"Metadata updated in files: {metadata_updated}")
    print(f"Windows Date Taken updated: {windows_date_taken_updated}")
    print(f"Files skipped: {skipped_files}")
    print(f"Total files processed: {processed_folder + processed_metadata + processed_filename + processed_manual}")
    print(f"Total files with errors: {len(error_files)}")
    
    # Save report
    report_file = Path(folder_path) / "media_date_update_report.txt"
    with open(report_file, 'w', encoding='utf-8') as f:
        f.write(f"Media Date Update Report\n")
        f.write(f"{'='*60}\n")
        f.write(f"Generated on: {datetime.now()}\n")
        f.write(f"Valid date range: {MIN_DATE.date()} to {MAX_DATE.date()}\n")
        f.write(f"Force metadata update: {'ENABLED' if FORCE_METADATA_UPDATE else 'DISABLED'}\n")
        f.write(f"Folder processed: {folder_path}\n")
        f.write(f"{'='*60}\n\n")
        
        f.write(f"DEPENDENCIES:\n")
        f.write(f"- exiftool: {'Available' if exiftool_available else 'NOT INSTALLED'}\n")
        f.write(f"- ffmpeg: {'Available' if ffmpeg_available else 'NOT INSTALLED'}\n")
        f.write(f"- ffprobe: {'Available' if ffprobe_available else 'NOT INSTALLED'}\n")
        f.write(f"- pywin32: {'Available' if has_win32 else 'NOT INSTALLED'}\n\n")
        
        f.write(f"EXTRACTION PRIORITY:\n")
        f.write(f"Images:\n")
        f.write(f"  1. Windows Date Taken\n")
        f.write(f"  2. DateTimeOriginal (Exif.Photo.DateTimeOriginal 0x9003)\n")
        f.write(f"  3. DateTime (Exif.Image.DateTime 0x0132)\n")
        f.write(f"  4. DateTimeDigitized (Exif.Photo.DateTimeDigitized 0x9004)\n\n")
        f.write(f"Videos:\n")
        f.write(f"  1. Windows Media Created\n")
        f.write(f"  2. Windows Date Taken\n")
        f.write(f"  3. creation_time (ffprobe)\n")
        f.write(f"  4. date / date_recorded (ffprobe)\n\n")
        f.write(f"RAW:\n")
        f.write(f"  1. Windows Date Taken\n")
        f.write(f"  2. DateTimeOriginal (exiftool)\n")
        f.write(f"  3. CreateDate (exiftool)\n\n")
        
        f.write(f"FILE TYPES:\n")
        f.write(f"Images: {', '.join(sorted(IMAGE_EXTENSIONS))}\n")
        f.write(f"RAW: {', '.join(sorted(RAW_EXTENSIONS))}\n")
        f.write(f"Videos: {', '.join(sorted(VIDEO_EXTENSIONS))}\n\n")
        
        f.write(f"PROPERTIES UPDATED:\n")
        f.write(f"Images: DateTimeOriginal, CreateDate, ModifyDate (via exiftool)\n")
        f.write(f"RAW: DateTimeOriginal, CreateDate (via exiftool)\n")
        f.write(f"Videos: creation_time, date, date_recorded (via ffmpeg)\n")
        f.write(f"All files: Windows Date Taken property\n")
        f.write(f"All files: File system Date Modified\n\n")
        
        f.write(f"DECISION RULES:\n")
        f.write(f"A. Folder + Metadata: same=use metadata, different=ask user\n")
        f.write(f"B. Folder only: compare with filename, ask user\n")
        f.write(f"C. Metadata + Filename: same=use metadata, different=ask user\n")
        f.write(f"D. Metadata only: ask user for confirmation\n")
        f.write(f"E. Filename only: ask user for confirmation\n")
        f.write(f"F. Nothing found: ask user for manual entry\n")
        f.write(f"WhatsApp files: filename shown as option with warning\n\n")
        
        f.write(f"METADATA UPDATE RULES:\n")
        f.write(f"- Update when no metadata exists\n")
        f.write(f"- Update when chosen date differs from existing metadata\n")
        f.write(f"- Update when FORCE_METADATA_UPDATE is enabled\n")
        f.write(f"- Skip when metadata exists and matches chosen date (unless forced)\n\n")
        
        f.write(f"SUMMARY:\n")
        f.write(f"- Files updated from folder name: {processed_folder}\n")
        f.write(f"- Files with existing metadata: {processed_metadata}\n")
        f.write(f"- Files updated from filename: {processed_filename}\n")
        f.write(f"- Files updated from manual entry: {processed_manual}\n")
        f.write(f"- Metadata updated in files: {metadata_updated}\n")
        f.write(f"- Windows Date Taken updated: {windows_date_taken_updated}\n")
        f.write(f"- Files skipped: {skipped_files}\n")
        f.write(f"- Total files processed: {processed_folder + processed_metadata + processed_filename + processed_manual}\n")
        f.write(f"- Total files with errors: {len(error_files)}\n")
        
        if error_files:
            f.write(f"\n{'='*60}\n")
            f.write(f"FILES WITH ERRORS ({len(error_files)} files):\n")
            f.write(f"{'='*60}\n")
            for file, error in sorted(error_files):
                f.write(f"{file}: {error}\n")
    
    print(f"\n📄 Full report saved to: {report_file}")


def main():
    if len(sys.argv) < 2:
        print("Usage: python fix_photo_dates.py <folder_path>")
        print("\nExample: python fix_photo_dates.py \"C:\\Users\\YourName\\Pictures\"")
        sys.exit(1)
    
    folder_path = sys.argv[1]
    
    if not Path(folder_path).exists():
        print(f"Error: Folder not found: {folder_path}")
        sys.exit(1)
    
    print(f"{'='*60}")
    print(f"MEDIA DATE FIXER")
    print(f"{'='*60}")
    print(f"\nSystem Information:")
    print(f"- Current date/time: {CURRENT_DATE}")
    print(f"- Platform: {sys.platform}")
    print(f"- Python version: {sys.version.split()[0]}")
    print(f"- Valid date range: {MIN_DATE.date()} to {MAX_DATE.date()}")
    print()
    
    print("Checking dependencies...")
    
    exiftool_available = check_exiftool_available()
    if exiftool_available:
        print("✓ exiftool found - image/RAW metadata support enabled")
    else:
        print("✗ exiftool NOT FOUND - image/RAW metadata cannot be read/updated")
        print("  Install from: https://exiftool.org/")
        print("  Windows: Download exe, rename to exiftool.exe, add to PATH")
        print("  Mac: brew install exiftool")
        print("  Linux: sudo apt install libimage-exiftool-perl")
    
    ffmpeg_available = check_ffmpeg_available()
    if ffmpeg_available:
        print("✓ ffmpeg found - video metadata updates enabled")
    else:
        print("✗ ffmpeg NOT FOUND - video metadata cannot be updated")
        print("  Install from: https://ffmpeg.org/download.html")
    
    ffprobe_available = check_ffprobe_available()
    if ffprobe_available:
        print("✓ ffprobe found - video metadata reading enabled")
    else:
        print("✗ ffprobe NOT FOUND - video metadata cannot be read")
    
    if sys.platform == 'win32':
        if has_win32:
            print("✓ pywin32 found - Windows Date Taken property access enabled")
        else:
            print("✗ pywin32 NOT FOUND - Windows Date Taken property cannot be accessed")
            print("  Install with: pip install pywin32")
    
    print()
    print("Supported file types:")
    print(f"- Images: {', '.join(sorted(ext.upper().lstrip('.') for ext in IMAGE_EXTENSIONS))}")
    print(f"- RAW: {', '.join(sorted(ext.upper().lstrip('.') for ext in RAW_EXTENSIONS))}")
    print(f"- Videos: {', '.join(sorted(ext.upper().lstrip('.') for ext in VIDEO_EXTENSIONS))}")
    print()
    
    print("Metadata Extraction Priority:")
    print("• Images:")
    print("    1. Windows Date Taken")
    print("    2. DateTimeOriginal (Exif.Photo.DateTimeOriginal 0x9003)")
    print("    3. DateTime (Exif.Image.DateTime 0x0132)")
    print("    4. DateTimeDigitized (Exif.Photo.DateTimeDigitized 0x9004)")
    print("• Videos:")
    print("    1. Windows Media Created")
    print("    2. Windows Date Taken")
    print("    3. creation_time (ffprobe)")
    print("    4. date / date_recorded (ffprobe)")
    print("• RAW:")
    print("    1. Windows Date Taken")
    print("    2. DateTimeOriginal (exiftool)")
    print("    3. CreateDate (exiftool)")
    print()
    
    print("Properties Updated:")
    print("• Images: DateTimeOriginal, CreateDate, ModifyDate (via exiftool)")
    print("• RAW: DateTimeOriginal, CreateDate (via exiftool)")
    print("• Videos: creation_time, date, date_recorded (via ffmpeg)")
    print("• All files: Windows Date Taken property")
    print("• All files: File system Date Modified")
    print()
    
    print("Timezone Handling:")
    print("• UTC timestamps (ending with 'Z') are converted to local time")
    print("• Timezone offsets (e.g., +05:30, -08:00) are converted to local time")
    print("• Timestamps without timezone info are assumed to be local time")
    print()
    
    print("Decision Rules:")
    print("• Folder date + Metadata exist (same date) → Use metadata")
    print("• Folder date + Metadata exist (different) → Ask user")
    print("• Folder date only → Compare with filename, ask user")
    print("• Metadata + Filename only (same) → Use metadata")
    print("• Metadata + Filename only (different) → Ask user")
    print("• Metadata only → Ask user for confirmation")
    print("• Filename only → Ask user for confirmation")
    print("• Nothing found → Ask user for manual entry")
    print("• WhatsApp files → Filename date extraction disabled")
    print()
    
    print("Metadata Update Rules:")
    print("• Update when no metadata exists")
    print("• Update when chosen date differs from existing metadata")
    print(f"• Force update: {'ENABLED' if FORCE_METADATA_UPDATE else 'DISABLED'}")
    print("  (Set FORCE_METADATA_UPDATE = True at top of script to always update)")
    print()
    
    print("Quality Preservation:")
    print("• No re-encoding of images or videos")
    print("• Only metadata is modified")
    print("• Original quality preserved")
    print()
    
    print("File Type Mismatch Detection:")
    print("• Detects when file extension doesn't match actual file content")
    print("• Offers to rename files to correct extension")
    print("• Common issue: HEIC files that are actually JPEG")
    print()
    
    process_folder(folder_path)


if __name__ == "__main__":
    main()
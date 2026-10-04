import os
import json
import uuid
from datetime import datetime
from typing import List, Dict, Any, Optional

HISTORY_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data', 'history.json')
MAX_HISTORY_ENTRIES = 200

# Hindi month names for friendly display
HINDI_MONTHS = {
    1: 'जनवरी', 2: 'फ़रवरी', 3: 'मार्च', 4: 'अप्रैल', 5: 'मई', 6: 'जून',
    7: 'जुलाई', 8: 'अगस्त', 9: 'सितम्बर', 10: 'अक्टूबर', 11: 'नवम्बर', 12: 'दिसम्बर'
}

def _ensure_dir():
    d = os.path.dirname(HISTORY_FILE)
    if not os.path.exists(d):
        os.makedirs(d, exist_ok=True)

def load_history() -> List[Dict[str, Any]]:
    """Loads all history records sorted newest first."""
    _ensure_dir()
    if not os.path.exists(HISTORY_FILE):
        return []
    try:
        with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
            if isinstance(data, list):
                return data
            return []
    except Exception:
        return []

def save_history(records: List[Dict[str, Any]]) -> None:
    """Saves history records to JSON."""
    _ensure_dir()
    try:
        with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
            json.dump(records[:MAX_HISTORY_ENTRIES], f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Error saving history: {e}")

def format_hindi_datetime(dt: datetime) -> Dict[str, str]:
    """Formats datetime into clean readable Hindi time string."""
    day = dt.day
    month = HINDI_MONTHS.get(dt.month, str(dt.month))
    year = dt.year
    hour = dt.strftime('%I').lstrip('0') or '12'
    minute = dt.strftime('%M')
    ampm = 'AM' if dt.hour < 12 else 'PM'
    
    date_str = f"{day:02d} {month} {year}"
    time_str = f"{hour}:{minute} {ampm}"
    short_time = dt.strftime('%d/%m/%Y, %I:%M %p')
    
    return {
        'display_date': date_str,
        'display_time': time_str,
        'full_str': f"{date_str} को {time_str}",
        'short_str': short_time
    }

def log_activity(
    feature_id: str,
    feature_name: str,
    file_names: List[str] or str,
    action: str,
    details: str = "",
    status: str = "सफल"
) -> Dict[str, Any]:
    """
    Logs an activity entry. Stores ONLY the file name(s), never full PDF binaries or paths.
    """
    if isinstance(file_names, str):
        file_names = [file_names]

    # Clean file names to ensure only base names are preserved (no paths, no binaries)
    clean_names = []
    for f in file_names:
        if f:
            base = os.path.basename(f)
            # Remove any prepended UUID prefixes if present
            base = os.sub(r'^[a-f0-9\-]{36}_', '', base) if hasattr(os, 'sub') else base
            import re
            base = re.sub(r'^[a-f0-9\-]{36}_', '', base)
            clean_names.append(base)

    if not clean_names:
        clean_names = ["बिना नाम फ़ाइल"]

    now = datetime.now()
    time_info = format_hindi_datetime(now)

    record = {
        'id': str(uuid.uuid4())[:8],
        'timestamp': now.isoformat(),
        'display_date': time_info['display_date'],
        'display_time': time_info['display_time'],
        'full_time': time_info['full_str'],
        'short_time': time_info['short_str'],
        'feature_id': feature_id,
        'feature_name': feature_name,
        'file_names': clean_names,
        'primary_file': clean_names[0],
        'file_count': len(clean_names),
        'action': action,
        'details': details,
        'status': status
    }

    records = load_history()
    records.insert(0, record)
    save_history(records)
    return record

def get_recent_history(limit: int = 50, feature_id: Optional[str] = None) -> List[Dict[str, Any]]:
    """Retrieves recent history, optionally filtered by feature."""
    records = load_history()
    if feature_id:
        records = [r for r in records if r.get('feature_id') == feature_id]
    return records[:limit]

def clear_all_history() -> bool:
    """Clears all history."""
    save_history([])
    return True

def delete_history_item(record_id: str) -> bool:
    """Deletes a single history record."""
    records = load_history()
    updated = [r for r in records if r.get('id') != record_id]
    save_history(updated)
    return len(updated) < len(records)

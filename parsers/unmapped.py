"""
Shared helper: identifies top-level JSON keys present in a source file that
the parser doesn't actively extract, so nothing sits silently unused without
at least being visible somewhere. Used to populate the workbook's
"Unmapped Data" sheet.
"""
import json


def find_unmapped_keys(raw_dict, known_keys):
    """Returns a list of {key, preview} for every top-level key in raw_dict
    not in known_keys. 'preview' shows the actual content (compact JSON,
    truncated) so real figures are visible at a glance -- e.g. if an unmapped
    field turns out to hold real tax amounts, you'll see the number directly
    instead of just 'list with 1 item(s)' and having to ask what's inside."""
    unmapped = []
    for key, value in raw_dict.items():
        if key in known_keys:
            continue
        unmapped.append({"key": key, "preview": _preview(value)})
    return unmapped


def _preview(value, max_len=300):
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        s = str(value)
        return s if len(s) <= max_len else s[:max_len] + "..."
    if isinstance(value, (list, dict)):
        try:
            s = json.dumps(value, separators=(",", ":"))
        except (TypeError, ValueError):
            return f"{'list' if isinstance(value, list) else 'dict'} (not serializable)"
        if len(s) <= max_len:
            return s
        return s[:max_len] + f"... [truncated, {len(s)} chars total]"
    return str(type(value))

"""
Official GST state/UT code table (first 2 digits of any GSTIN).
Source: CBIC-notified state code list.
"""
GST_STATE_CODES = {
    "01": "Jammu and Kashmir", "02": "Himachal Pradesh", "03": "Punjab",
    "04": "Chandigarh", "05": "Uttarakhand", "06": "Haryana", "07": "Delhi",
    "08": "Rajasthan", "09": "Uttar Pradesh", "10": "Bihar", "11": "Sikkim",
    "12": "Arunachal Pradesh", "13": "Nagaland", "14": "Manipur", "15": "Mizoram",
    "16": "Tripura", "17": "Meghalaya", "18": "Assam", "19": "West Bengal",
    "20": "Jharkhand", "21": "Odisha", "22": "Chhattisgarh", "23": "Madhya Pradesh",
    "24": "Gujarat", "25": "Daman and Diu", "26": "Dadra and Nagar Haveli",
    "27": "Maharashtra", "28": "Andhra Pradesh (Old)", "29": "Karnataka",
    "30": "Goa", "31": "Lakshadweep", "32": "Kerala", "33": "Tamil Nadu",
    "34": "Puducherry", "35": "Andaman and Nicobar Islands", "36": "Telangana",
    "37": "Andhra Pradesh", "38": "Ladakh", "97": "Other Territory",
    "99": "Centre Jurisdiction",
}


def state_name_from_gstin(gstin):
    """Returns the state/UT name for a GSTIN, using its first 2 digits.
    Falls back to the raw code if it's not in the table (new codes get added
    occasionally, so this shouldn't hard-fail)."""
    if not gstin or len(gstin) < 2:
        return "Unknown"
    code = gstin[:2]
    return GST_STATE_CODES.get(code, f"State-{code}")


def safe_filename_component(text):
    """Strips characters that are invalid in Windows/Mac/Linux filenames."""
    invalid = '<>:"/\\|?*'
    return "".join(c for c in text if c not in invalid).strip()

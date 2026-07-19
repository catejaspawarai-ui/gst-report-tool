"""
Parser for GSTR-2B Summary JSON (portal export, the '_Summary.json' file).
GSTR-2B ships pre-aggregated ITC totals under data.itcsumm, broken into:
  itcavl      -> ITC available (imports, reverse-charge supplies, ISD, others, non-reverse-charge)
  itcunavl    -> ITC not available (e.g. blocked under Sec 17(5), supplier not filed, etc.)
  itcRejected -> ITC in documents you've rejected in IMS
We roll these up into IGST/CGST/SGST/Cess totals for the summary report.

Taxable value note: each category's txval is NOT a top-level field -- it's
nested inside each category's sub-keys (b2b, cdnr, cdnra, cdnrrev, isd, etc.),
and different categories have different sub-keys. We sum txval across all
nested dicts within a category to get its taxable value. The 'imports'
category has no txval anywhere in this JSON at all (Bill of Entry customs
value isn't carried in this statement) -- that's a genuine source-data gap,
not something we can compute from here.
"""
import json


def _sum_nested_txval(cat):
    """Sums txval across every nested sub-dict in a category (b2b, cdnr, isd, etc.)."""
    total = 0.0
    for key, val in cat.items():
        if isinstance(val, dict):
            total += val.get("txval", 0) or 0
    return total


def _category_totals(cat):
    return {
        "txval": _sum_nested_txval(cat),
        "igst": cat.get("igst", 0) or 0,
        "cgst": cat.get("cgst", 0) or 0,
        "sgst": cat.get("sgst", 0) or 0,
        "cess": cat.get("cess", 0) or 0,
    }


def _rollup(section):
    """Sum igst/cgst/sgst/cess/txval across all categories in an itcsumm section."""
    out = {"txval": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
    for cat_name, cat in section.items():
        if not isinstance(cat, dict):
            continue
        totals = _category_totals(cat)
        for k in out:
            out[k] += totals[k]
    return out


def parse_r2b_summary(filepath):
    with open(filepath, "r", encoding="utf-8") as f:
        d = json.load(f)

    data = d.get("data", {})
    itcsumm = data.get("itcsumm", {})
    itcavl_raw = itcsumm.get("itcavl", {})
    itcunavl_raw = itcsumm.get("itcunavl", {})
    itcrej_raw = itcsumm.get("itcRejected", {})

    itc_available = _rollup(itcavl_raw)
    itc_unavailable = _rollup(itcunavl_raw)
    itc_rejected = _rollup(itcrej_raw)

    # Per-category breakdown -- lets the comparator match "like against like"
    # (e.g. R3B's IMPG line against R2B's "imports" category only), instead of
    # comparing grand totals where a reverse-charge/RCM gap gets misread as a
    # general ITC over-claim.
    by_category = {cat_name: _category_totals(cat) for cat_name, cat in itcavl_raw.items() if isinstance(cat, dict)}

    return {
        "gstin": data.get("gstin"),
        "period": data.get("rtnprd"),
        "generated_on": data.get("gendt"),
        "itc_available": itc_available,
        "itc_available_by_category": by_category,  # imports / revsup / isdsup / othersup / nonrevsup
        "itc_unavailable": itc_unavailable,  # {txval, igst, cgst, sgst, cess}
        "itc_rejected": itc_rejected,  # {txval, igst, cgst, sgst, cess}
        "itc_available_total": itc_available["igst"] + itc_available["cgst"] + itc_available["sgst"] + itc_available["cess"],
    }


if __name__ == "__main__":
    import sys
    print(json.dumps(parse_r2b_summary(sys.argv[1]), indent=2))

"""
Builds the two comparison views:
  - GSTR-3B vs GSTR-1 : outward liability (taxable value & tax) match
  - GSTR-3B vs GSTR-2B: ITC claimed vs ITC available match, done CATEGORY BY
    CATEGORY (Import of Goods, Reverse Charge/RCM, ISD, Other/Regular ITC)
    rather than as one grand total. This matters: RCM (ISRC) credit is
    self-declared by the taxpayer against tax they paid, not sourced from a
    supplier's GSTR-1 filing, so it will never show up under GSTR-2B's
    "reverse charge supplies" category in the same way -- comparing grand
    totals made this look like a ~1.65 Cr over-claim in testing, when it was
    actually a normal, expected gap. Category matching avoids that false flag.
"""

TOLERANCE = 100  # rupees; variances below this are rounding noise, not flagged

# Maps 3B's itc_by_type keys to the matching 2B itc_available_by_category keys.
# OTH (3B "all other ITC") nets against BOTH 2B's regular B2B credit (nonrevsup)
# and CDNR-driven adjustments (othersup), since 3B's Table 4A.5 doesn't split them.
CATEGORY_MAP = {
    "IMPG": {"label": "Import of Goods", "r2b_keys": ["imports"]},
    "ISRC": {"label": "Reverse Charge (RCM)", "r2b_keys": ["revsup"]},
    "ISD": {"label": "ISD Credit", "r2b_keys": ["isdsup"]},
    "OTH": {"label": "All Other ITC (B2B + CDNR adj.)", "r2b_keys": ["nonrevsup", "othersup"]},
}


def compare_3b_vs_1(r3b, r1):
    liability_txval_3b = r3b["total_liability_txval"]
    liability_tax_3b = r3b["total_liability_tax"]

    txval_1 = r1["total_taxable_value"]
    tax_1 = r1["total_tax"]

    txval_variance = round(liability_txval_3b - txval_1, 2)
    tax_variance = round(liability_tax_3b - tax_1, 2)

    return {
        "gstin": r3b["gstin"],
        "period": r3b["period"],
        "taxable_value_3b": liability_txval_3b,
        "taxable_value_1": txval_1,
        "taxable_value_variance": txval_variance,
        "tax_3b": liability_tax_3b,
        "tax_1": tax_1,
        "tax_variance": tax_variance,
        "flag": "REVIEW" if abs(tax_variance) > TOLERANCE or abs(txval_variance) > TOLERANCE else "OK",
    }


def compare_3b_vs_2b(r3b, r2b):
    itc_3b_by_type = r3b.get("itc_by_type", {})
    itc_2b_by_cat = r2b.get("itc_available_by_category", {})

    rows = []
    for r3b_key, mapping in CATEGORY_MAP.items():
        claimed = itc_3b_by_type.get(r3b_key, {"igst": 0, "cgst": 0, "sgst": 0, "cess": 0})
        available = {"igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
        for r2b_key in mapping["r2b_keys"]:
            cat = itc_2b_by_cat.get(r2b_key, {})
            for head in available:
                available[head] += cat.get(head, 0) or 0

        for head in ("igst", "cgst", "sgst", "cess"):
            c = claimed.get(head, 0) or 0
            a = available.get(head, 0) or 0
            variance = round(c - a, 2)
            flag = "OK"
            if variance > TOLERANCE:
                flag = "CLAIMED > AVAILABLE"
            elif abs(variance) > TOLERANCE:
                flag = "REVIEW"
            rows.append({
                "category": mapping["label"],
                "tax_head": head.upper(),
                "itc_claimed_3b": c,
                "itc_available_2b": a,
                "variance": variance,
                "flag": flag,
            })

    total_variance = round(sum(r["variance"] for r in rows), 2)
    return {
        "gstin": r3b["gstin"],
        "period": r3b["period"],
        "rows": rows,
        "total_variance": total_variance,
        "any_over_claimed": any(r["flag"] == "CLAIMED > AVAILABLE" for r in rows),
    }

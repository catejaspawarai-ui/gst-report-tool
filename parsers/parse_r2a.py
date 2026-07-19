"""
Parser for GSTR-2A JSON (portal export, inside the R2A zip).
Unlike 2B, 2A ships no pre-built summary -- we aggregate invoice-level b2b,
cdn (credit/debit notes), impg (imports), and isd (ISD credit) ourselves.
"""
import json


def _sum_itm(itm_list):
    out = {"txval": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
    for itm in itm_list:
        det = itm.get("itm_det", itm)
        out["txval"] += det.get("txval", 0) or 0
        out["igst"] += det.get("iamt", 0) or 0
        out["cgst"] += det.get("camt", 0) or 0
        out["sgst"] += det.get("samt", 0) or 0
        out["cess"] += det.get("csamt", 0) or 0
    return out


def _sum_b2b(block):
    out = {"txval": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
    doc_count = 0
    for entry in block:
        for inv in entry.get("inv", []):
            doc_count += 1
            t = _sum_itm(inv.get("itms", []))
            for k in out:
                out[k] += t[k]
    out["doc_count"] = doc_count
    return out


def _sum_cdn(block):
    """cdn: credit notes ('C') reduce ITC, debit notes ('D') increase it."""
    out = {"txval": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
    doc_count = 0
    for entry in block:
        for nt in entry.get("nt", []):
            doc_count += 1
            t = _sum_itm(nt.get("itms", []))
            sign = -1 if nt.get("ntty", "C") == "C" else 1
            for k in out:
                out[k] += sign * t[k]
    out["doc_count"] = doc_count
    return out


def _sum_impg(block):
    out = {"txval": 0.0, "igst": 0.0, "cess": 0.0}
    for row in block:
        out["txval"] += row.get("txval", 0) or 0
        out["igst"] += row.get("iamt", 0) or 0
        out["cess"] += row.get("csamt", 0) or 0
    out["doc_count"] = len(block)
    return out


def _sum_isd(block):
    out = {"igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
    doc_count = 0
    for entry in block:
        for doc in entry.get("doclist", []):
            doc_count += 1
            out["igst"] += doc.get("iamt", 0) or 0
            out["cgst"] += doc.get("camt", 0) or 0
            out["sgst"] += doc.get("samt", 0) or 0
            out["cess"] += doc.get("cess", 0) or 0
    out["doc_count"] = doc_count
    return out


def parse_r2a(filepath):
    with open(filepath, "r", encoding="utf-8") as f:
        d = json.load(f)

    b2b = _sum_b2b(d.get("b2b", []))
    cdn = _sum_cdn(d.get("cdn", []))
    impg = _sum_impg(d.get("impg", []))
    isd = _sum_isd(d.get("isd", []))

    itc_total = {
        "igst": b2b["igst"] + cdn["igst"] + impg["igst"] + isd["igst"],
        "cgst": b2b["cgst"] + cdn["cgst"] + isd["cgst"],
        "sgst": b2b["sgst"] + cdn["sgst"] + isd["sgst"],
        "cess": b2b["cess"] + cdn["cess"] + impg["cess"] + isd["cess"],
    }

    return {
        "gstin": d.get("gstin"),
        "period": d.get("fp"),
        "b2b": b2b,
        "cdn": cdn,
        "impg": impg,
        "isd": isd,
        "itc_available_total_by_head": itc_total,
        "itc_available_total": sum(itc_total.values()),
    }


if __name__ == "__main__":
    import sys
    print(json.dumps(parse_r2a(sys.argv[1]), indent=2))

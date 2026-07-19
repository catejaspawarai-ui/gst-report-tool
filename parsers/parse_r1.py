"""
Parser for GSTR-1 JSON (portal export, inside the R1 zip).
Extracts summary fields matching the GSTR-1 Summary report card.

Note on tax fields: interstate invoice lines (itm_det) carry only 'iamt' (IGST).
Intrastate lines carry 'camt'/'samt' (CGST/SGST) instead of 'iamt'. We sum
whichever keys are present, so this works for both without double counting.
"""
import json


def _sum_itm(itm_list):
    out = {"txval": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
    for itm in itm_list:
        det = itm.get("itm_det", itm)  # hsn rows don't nest itm_det
        out["txval"] += det.get("txval", 0) or 0
        out["igst"] += det.get("iamt", 0) or 0
        out["cgst"] += det.get("camt", 0) or 0
        out["sgst"] += det.get("samt", 0) or 0
        out["cess"] += det.get("csamt", 0) or 0
    return out


def _sum_invoice_block(block, doc_key="inv", note_filter=None):
    """
    b2b / cdnr / exp share the same nested {ctin, inv[...]} shape (doc_key='inv'
    for b2b/exp, 'nt' for cdnr notes).
    note_filter, if given, is a predicate on the note's 'ntty' ('C'/'D') --
    used to split cdnr into separate credit-note and debit-note totals.
    """
    out = {"txval": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
    doc_count = 0
    for entry in block:
        invoices = entry.get(doc_key, [])
        for inv in invoices:
            if note_filter and not note_filter(inv.get("ntty", "C")):
                continue
            doc_count += 1
            itm_totals = _sum_itm(inv.get("itms", []))
            for k in out:
                out[k] += itm_totals[k]
    out["doc_count"] = doc_count
    return out


def _sum_exp_by_type(exp_block):
    """Splits exports into WPAY (with payment of tax) and WOPAY (LUT/bond, no tax)."""
    result = {}
    for typ in ("WPAY", "WOPAY"):
        matching = [e for e in exp_block if e.get("exp_typ") == typ]
        totals = {"txval": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
        doc_count = 0
        for entry in matching:
            for inv in entry.get("inv", []):
                doc_count += 1
                t = _sum_itm(inv.get("itms", []))
                for k in totals:
                    totals[k] += t[k]
        totals["doc_count"] = doc_count
        result[typ] = totals
    return result


def _sum_notes_flat(block, note_filter=None, type_filter=None):
    """cdnur: each array entry IS the note (itms directly on it, no ctin/inv wrapper).
    type_filter, if given, filters by the note's 'typ' field ('EXPWP'/'EXPWOP' for
    export-related unregistered notes, vs domestic types like 'B2CL') -- these need
    to land in the matching outward-supply section, not be lumped together."""
    out = {"txval": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
    doc_count = 0
    for note in block:
        ntty = note.get("ntty", "C")
        if note_filter and not note_filter(ntty):
            continue
        if type_filter and not type_filter(note.get("typ", "")):
            continue
        itm_totals = _sum_itm(note.get("itms", []))
        doc_count += 1
        for k in out:
            out[k] += itm_totals[k]
    out["doc_count"] = doc_count
    return out


def _sum_b2cs(b2cs_list):
    out = {"txval": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
    for row in b2cs_list:
        out["txval"] += row.get("txval", 0) or 0
        out["igst"] += row.get("iamt", 0) or 0
        out["cgst"] += row.get("camt", 0) or 0
        out["sgst"] += row.get("samt", 0) or 0
        out["cess"] += row.get("csamt", 0) or 0
    out["doc_count"] = len(b2cs_list)
    return out


def parse_r1(filepath):
    with open(filepath, "r", encoding="utf-8") as f:
        d = json.load(f)

    gstin = d.get("gstin")
    period = d.get("fp")  # MMYYYY

    b2b = _sum_invoice_block(d.get("b2b", []))
    cdnr_credit = _sum_invoice_block(d.get("cdnr", []), doc_key="nt", note_filter=lambda t: t == "C")
    cdnr_debit = _sum_invoice_block(d.get("cdnr", []), doc_key="nt", note_filter=lambda t: t == "D")
    is_export_typ = lambda typ: typ in ("EXPWP", "EXPWOP")
    cdnur_dom_credit = _sum_notes_flat(d.get("cdnur", []), note_filter=lambda t: t == "C", type_filter=lambda t: not is_export_typ(t))
    cdnur_dom_debit = _sum_notes_flat(d.get("cdnur", []), note_filter=lambda t: t == "D", type_filter=lambda t: not is_export_typ(t))
    cdnur_exp_credit = _sum_notes_flat(d.get("cdnur", []), note_filter=lambda t: t == "C", type_filter=is_export_typ)
    cdnur_exp_debit = _sum_notes_flat(d.get("cdnur", []), note_filter=lambda t: t == "D", type_filter=is_export_typ)
    exp_by_type = _sum_exp_by_type(d.get("exp", []))
    b2cs = _sum_b2cs(d.get("b2cs", []))

    def _signed(block, sign):
        return {k: sign * v for k, v in block.items() if k != "doc_count"}

    # Net effect: credit notes reduce outward liability, debit notes increase it.
    total_txval = (
        b2b["txval"]
        - cdnr_credit["txval"] + cdnr_debit["txval"]
        - cdnur_dom_credit["txval"] + cdnur_dom_debit["txval"]
        - cdnur_exp_credit["txval"] + cdnur_exp_debit["txval"]
        + exp_by_type["WPAY"]["txval"] + exp_by_type["WOPAY"]["txval"]
        + b2cs["txval"]
    )
    total_tax = 0.0
    for head in ("igst", "cgst", "sgst", "cess"):
        total_tax += (
            b2b[head]
            - cdnr_credit[head] + cdnr_debit[head]
            - cdnur_dom_credit[head] + cdnur_dom_debit[head]
            - cdnur_exp_credit[head] + cdnur_exp_debit[head]
            + exp_by_type["WPAY"][head] + exp_by_type["WOPAY"][head]
            + b2cs[head]
        )

    return {
        "gstin": gstin,
        "period": period,
        "b2b": b2b,
        "cdnr_credit": cdnr_credit,
        "cdnr_debit": cdnr_debit,
        "cdnur_credit": cdnur_dom_credit,   # domestic unregistered notes only
        "cdnur_debit": cdnur_dom_debit,
        "cdnur_exp_credit": cdnur_exp_credit,  # export-related unregistered notes
        "cdnur_exp_debit": cdnur_exp_debit,
        "exp_wpay": exp_by_type["WPAY"],
        "exp_wopay": exp_by_type["WOPAY"],
        "b2cs": b2cs,
        "grand_total_declared": d.get("cur_gt", d.get("gt", 0)),
        "total_taxable_value": total_txval,
        "total_tax": total_tax,
    }


if __name__ == "__main__":
    import sys
    print(json.dumps(parse_r1(sys.argv[1]), indent=2))

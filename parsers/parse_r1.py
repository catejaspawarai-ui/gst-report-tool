"""
Parser for GSTR-1 JSON (portal export, inside the R1 zip).
Extracts summary fields matching the GSTR-1 Summary report card.

Note on tax fields: interstate invoice lines (itm_det) carry only 'iamt' (IGST).
Intrastate lines carry 'camt'/'samt' (CGST/SGST) instead of 'iamt'. We sum
whichever keys are present, so this works for both without double counting.
"""
import json
from parsers.unmapped import find_unmapped_keys


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


def _sum_invoice_block(block, doc_key="inv", note_filter=None, type_filter=None):
    """
    b2b / cdnr / exp share the same nested {ctin, inv[...]} shape (doc_key='inv'
    for b2b/exp, 'nt' for cdnr notes).
    note_filter, if given, is a predicate on the note's 'ntty' ('C'/'D') --
    used to split cdnr into separate credit-note and debit-note totals.
    type_filter, if given, is a predicate on 'inv_typ' (R/SEWP/SEWOP/DE) --
    used to separate Regular+Deemed Export (taxed normally, belongs in
    GSTR-3B 3.1.A) from SEZ supplies (zero-rated, belongs in 3.1.B). Confirmed
    against a real GSTN-generated file: SEWP carries real tax, SEWOP doesn't,
    DE carries real tax like Regular.
    """
    out = {"txval": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
    doc_count = 0
    for entry in block:
        invoices = entry.get(doc_key, [])
        for inv in invoices:
            if note_filter and not note_filter(inv.get("ntty", "C")):
                continue
            if type_filter and not type_filter(inv.get("inv_typ", "R")):
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


def _parse_hsn_summary(hsn_dict):
    """Combines HSN entries, aggregated by (HSN code, description, rate) --
    matching official GSTR-1 Table 12 granularity. Same HSN code can
    legitimately appear at multiple tax rates, so we keep rate as part of the key.

    Handles two JSON schema variations seen in real portal exports:
      - hsn_b2b / hsn_b2c as separate arrays (seen in some periods)
      - a single combined 'data' array with no B2B/B2C split (seen in others)
    Checking both rather than assuming one, since the portal's export schema
    isn't consistent across all periods/versions.

    Returns (rows, schema_variant) -- schema_variant is surfaced in the report
    so anyone reviewing it knows the source format differed across periods,
    rather than silently blending two schemas without comment."""
    combined = {}
    schema_variant = "none"

    def _add_entries(entries):
        for entry in entries:
            key = (entry.get("hsn_sc", ""), entry.get("desc", ""), entry.get("rt", 0))
            row = combined.setdefault(key, {"qty": 0.0, "txval": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0, "uqc": entry.get("uqc", "")})
            row["qty"] += entry.get("qty", 0) or 0
            row["txval"] += entry.get("txval", 0) or 0
            row["igst"] += entry.get("iamt", 0) or 0
            row["cgst"] += entry.get("camt", 0) or 0
            row["sgst"] += entry.get("samt", 0) or 0
            row["cess"] += entry.get("csamt", 0) or 0

    if "hsn_b2b" in hsn_dict or "hsn_b2c" in hsn_dict:
        schema_variant = "b2b_b2c_split"
        _add_entries(hsn_dict.get("hsn_b2b", []))
        _add_entries(hsn_dict.get("hsn_b2c", []))
    elif "data" in hsn_dict and isinstance(hsn_dict["data"], list):
        schema_variant = "combined_data"
        _add_entries(hsn_dict["data"])

    rows = [
        {"hsn_code": k[0], "description": k[1], "rate": k[2], **v}
        for k, v in combined.items()
    ]
    return rows, schema_variant


def _parse_nil(nil_dict):
    """Nil-rated/Exempted/Non-GST supplies, broken down by supply type
    (Inter/Intra-state, B2B/B2C). Verified against a real government-generated
    export: sply_ty values seen are INTRB2B, INTRAB2B, INTRB2C, INTRAB2C."""
    out = {"exempted": 0.0, "nil_rated": 0.0, "non_gst": 0.0}
    for entry in nil_dict.get("inv", []):
        out["exempted"] += entry.get("expt_amt", 0) or 0
        out["nil_rated"] += entry.get("nil_amt", 0) or 0
        out["non_gst"] += entry.get("ngsup_amt", 0) or 0
    return out


def _parse_advances(entries):
    """Shared shape for AT (Advances Received, Table 11A), ATA (its amendment),
    and TXPD (Advance Adjustments, Table 11B -- confirmed via a real file that
    the actual JSON key is 'txpd', not 'atadj' as the Excel template's sheet
    name suggested). Each entry has itms[] with rt/ad_amt/tax amounts."""
    out = {"ad_amt": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
    for entry in entries:
        for itm in entry.get("itms", []):
            out["ad_amt"] += itm.get("ad_amt", 0) or 0
            out["igst"] += itm.get("iamt", 0) or 0
            out["cgst"] += itm.get("camt", 0) or 0
            out["sgst"] += itm.get("samt", 0) or 0
            out["cess"] += itm.get("csamt", 0) or 0
    return out


def _parse_b2cs_like(entries):
    """B2CSA (B2CS amendments) shape: entry has itms[] with rt/txval/tax --
    slightly different from base b2cs (which has txval directly on the entry)."""
    out = {"txval": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
    for entry in entries:
        for itm in entry.get("itms", []):
            out["txval"] += itm.get("txval", 0) or 0
            out["igst"] += itm.get("iamt", 0) or 0
            out["cgst"] += itm.get("camt", 0) or 0
            out["sgst"] += itm.get("samt", 0) or 0
            out["cess"] += itm.get("csamt", 0) or 0
    return out


# Official GSTN document type codes (GSTR-1 Table 13)
DOC_TYPE_LABELS = {
    1: "Invoice for Outward Supply",
    2: "Invoice for Inward Supply from Unregistered Person",
    3: "Revised Invoice",
    4: "Debit Note",
    5: "Credit Note",
    6: "Receipt Voucher",
    7: "Payment Voucher",
    8: "Refund Voucher",
    9: "Delivery Challan (Job Work)",
    10: "Delivery Challan (Supply on Approval)",
    11: "Delivery Challan (Liquid Gas)",
    12: "Delivery Challan (Other than Supply)",
}


def _parse_doc_issue_summary(doc_issue_dict):
    """Each doc_num is an official GSTN document type code; each has one or more
    numbered series (docs[]) with from/to ranges and cancellation counts.
    Newer exports include the human-readable label directly as 'doc_typ' --
    we prefer that when present (more robust/future-proof than our own
    hardcoded mapping), falling back to DOC_TYPE_LABELS for older files that
    don't include it (confirmed: some real files have it, some don't)."""
    rows = []
    for entry in doc_issue_dict.get("doc_det", []):
        doc_num = entry.get("doc_num")
        label = entry.get("doc_typ") or DOC_TYPE_LABELS.get(doc_num, f"Document Type {doc_num}")
        total_issued = 0
        total_cancelled = 0
        total_net = 0
        for doc in entry.get("docs", []):
            total_issued += doc.get("totnum", 0) or 0
            total_cancelled += doc.get("cancel", 0) or 0
            total_net += doc.get("net_issue", 0) or 0
        rows.append({
            "doc_type": label,
            "series_count": len(entry.get("docs", [])),
            "total_issued": total_issued,
            "total_cancelled": total_cancelled,
            "net_issued": total_net,
        })
    return rows


def parse_r1(filepath):
    with open(filepath, "r", encoding="utf-8") as f:
        d = json.load(f)

    gstin = d.get("gstin")
    period = d.get("fp")  # MMYYYY

    is_sez_type = lambda t: t in ("SEWP", "SEWOP")
    is_taxable_type = lambda t: t not in ("SEWP", "SEWOP")  # Regular + Deemed Export

    b2b = _sum_invoice_block(d.get("b2b", []), type_filter=is_taxable_type)
    b2b_sez_wp = _sum_invoice_block(d.get("b2b", []), type_filter=lambda t: t == "SEWP")
    b2b_sez_wop = _sum_invoice_block(d.get("b2b", []), type_filter=lambda t: t == "SEWOP")
    b2b_deemed_exp = _sum_invoice_block(d.get("b2b", []), type_filter=lambda t: t == "DE")

    cdnr_credit = _sum_invoice_block(d.get("cdnr", []), doc_key="nt", note_filter=lambda t: t == "C", type_filter=is_taxable_type)
    cdnr_debit = _sum_invoice_block(d.get("cdnr", []), doc_key="nt", note_filter=lambda t: t == "D", type_filter=is_taxable_type)
    cdnr_sez_credit = _sum_invoice_block(d.get("cdnr", []), doc_key="nt", note_filter=lambda t: t == "C", type_filter=is_sez_type)
    cdnr_sez_debit = _sum_invoice_block(d.get("cdnr", []), doc_key="nt", note_filter=lambda t: t == "D", type_filter=is_sez_type)
    is_export_typ = lambda typ: typ in ("EXPWP", "EXPWOP")
    cdnur_dom_credit = _sum_notes_flat(d.get("cdnur", []), note_filter=lambda t: t == "C", type_filter=lambda t: not is_export_typ(t))
    cdnur_dom_debit = _sum_notes_flat(d.get("cdnur", []), note_filter=lambda t: t == "D", type_filter=lambda t: not is_export_typ(t))
    cdnur_exp_credit = _sum_notes_flat(d.get("cdnur", []), note_filter=lambda t: t == "C", type_filter=is_export_typ)
    cdnur_exp_debit = _sum_notes_flat(d.get("cdnur", []), note_filter=lambda t: t == "D", type_filter=is_export_typ)
    # CDNUR amendments -- combined (not split domestic/export) for simplicity
    cdnura_credit = _sum_notes_flat(d.get("cdnura", []), note_filter=lambda t: t == "C")
    cdnura_debit = _sum_notes_flat(d.get("cdnura", []), note_filter=lambda t: t == "D")
    exp_by_type = _sum_exp_by_type(d.get("exp", []))
    b2cs = _sum_b2cs(d.get("b2cs", []))

    # Amendments -- reported in the period the correction is MADE, not the
    # period of the original invoice, so they genuinely add to (or reduce)
    # THIS period's liability. Omitting them would create a false variance
    # against GSTR-3B for any period with amendments.
    b2b_amendments = _sum_invoice_block(d.get("b2ba", []))
    exp_amendments_by_type = _sum_exp_by_type(d.get("expa", []))

    # B2C Large (inter-state, high-value consumer supplies) -- same shape as b2b
    b2cl = _sum_invoice_block(d.get("b2cl", []))
    b2cl_amendments = _sum_invoice_block(d.get("b2cla", []))

    # B2CS amendments -- itms[]-nested shape, different from base b2cs
    b2cs_amendments = _parse_b2cs_like(d.get("b2csa", []))

    # Nil-rated/Exempted/Non-GST -- zero tax by definition, tracked separately,
    # not part of the taxable-value/tax totals used for 3B comparison
    nil_supplies = _parse_nil(d.get("nil", {}))

    # Advances: Table 11A (received) adds to this period's liability; Table
    # 11B (adjustment, real JSON key 'txpd') subtracts the portion now
    # invoiced, to avoid double-counting -- same netting logic as credit notes.
    advances_received = _parse_advances(d.get("at", []))
    advances_received_amendments = _parse_advances(d.get("ata", []))
    advances_adjusted = _parse_advances(d.get("txpd", []))
    advances_adjusted_amendments = _parse_advances(d.get("txpda", []))

    # CDNR amendments (registered) -- confirmed real structure: grouped by
    # ctin like cdnr itself, NOT flat like cdnura (unregistered notes
    # amendments). These are two genuinely different fields.
    cdnra_credit = _sum_invoice_block(d.get("cdnra", []), doc_key="nt", note_filter=lambda t: t == "C")
    cdnra_debit = _sum_invoice_block(d.get("cdnra", []), doc_key="nt", note_filter=lambda t: t == "D")

    # E-commerce supplies U/s 9(5) -- confirmed real structure: ecom.b2b (has
    # stin=ECO GSTIN, rtin=recipient GSTIN, nested inv[]), ecom.b2c and
    # ecom.urp2c (flat, direct tax fields like b2cs).
    ecom = d.get("ecom", {})
    ecom_b2b = _sum_invoice_block(ecom.get("b2b", []))
    ecom_b2c = {"txval": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
    for entry in ecom.get("b2c", []):
        ecom_b2c["txval"] += entry.get("txval", 0) or 0
        ecom_b2c["igst"] += entry.get("iamt", 0) or 0
        ecom_b2c["cgst"] += entry.get("camt", 0) or 0
        ecom_b2c["sgst"] += entry.get("samt", 0) or 0
        ecom_b2c["cess"] += entry.get("csamt", 0) or 0
    ecom_urp2c = {"txval": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
    for entry in ecom.get("urp2c", []):
        ecom_urp2c["txval"] += entry.get("txval", 0) or 0
        ecom_urp2c["igst"] += entry.get("iamt", 0) or 0
        ecom_urp2c["cgst"] += entry.get("camt", 0) or 0
        ecom_urp2c["sgst"] += entry.get("samt", 0) or 0
        ecom_urp2c["cess"] += entry.get("csamt", 0) or 0

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
        + b2b_amendments["txval"]
        + exp_amendments_by_type["WPAY"]["txval"] + exp_amendments_by_type["WOPAY"]["txval"]
        + b2cl["txval"] + b2cl_amendments["txval"]
        + b2cs_amendments["txval"]
        - cdnura_credit["txval"] + cdnura_debit["txval"]
        + b2b_sez_wp["txval"] + b2b_sez_wop["txval"]
        - cdnr_sez_credit["txval"] + cdnr_sez_debit["txval"]
        - cdnra_credit["txval"] + cdnra_debit["txval"]
        + ecom_b2b["txval"] + ecom_b2c["txval"] + ecom_urp2c["txval"]
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
            + b2b_amendments[head]
            + exp_amendments_by_type["WPAY"][head] + exp_amendments_by_type["WOPAY"][head]
            + b2cl[head] + b2cl_amendments[head]
            + b2cs_amendments[head]
            - cdnura_credit[head] + cdnura_debit[head]
            + b2b_sez_wp[head] + b2b_sez_wop[head]
            - cdnr_sez_credit[head] + cdnr_sez_debit[head]
            - cdnra_credit[head] + cdnra_debit[head]
            + ecom_b2b[head] + ecom_b2c[head] + ecom_urp2c[head]
        )
    # Advances: received (11A) adds liability; adjusted/'txpd' (11B) subtracts
    # the portion now invoiced elsewhere, to avoid double-counting.
    for head in ("igst", "cgst", "sgst", "cess"):
        total_tax += (
            advances_received[head] + advances_received_amendments[head]
            - advances_adjusted[head] - advances_adjusted_amendments[head]
        )

    hsn_rows, hsn_schema_variant = _parse_hsn_summary(d.get("hsn", {}))

    known_keys = {"gstin", "fp", "b2b", "cdnr", "exp", "cdnur", "b2cs", "hsn",
                  "doc_issue", "gt", "cur_gt", "filing_typ", "fil_dt", "b2ba", "expa",
                  "b2cl", "b2cla", "b2csa", "nil", "at", "ata", "txpd", "txpda", "cdnura",
                  "cdnra", "ecom", "ecoma", "version", "hash"}
    unmapped = find_unmapped_keys(d, known_keys)

    return {
        "gstin": gstin,
        "period": period,
        "b2b": b2b,
        "b2b_sez_wp": b2b_sez_wp,
        "b2b_sez_wop": b2b_sez_wop,
        "b2b_deemed_exp": b2b_deemed_exp,
        "cdnr_sez_credit": cdnr_sez_credit,
        "cdnr_sez_debit": cdnr_sez_debit,
        "cdnra_credit": cdnra_credit,
        "cdnra_debit": cdnra_debit,
        "ecom_b2b": ecom_b2b,
        "ecom_b2c": ecom_b2c,
        "ecom_urp2c": ecom_urp2c,
        "b2b_amendments": b2b_amendments,
        "exp_wpay_amendments": exp_amendments_by_type["WPAY"],
        "exp_wopay_amendments": exp_amendments_by_type["WOPAY"],
        "b2cl": b2cl,
        "b2cl_amendments": b2cl_amendments,
        "b2cs_amendments": b2cs_amendments,
        "nil_supplies": nil_supplies,
        "advances_received": advances_received,
        "advances_received_amendments": advances_received_amendments,
        "advances_adjusted": advances_adjusted,
        "advances_adjusted_amendments": advances_adjusted_amendments,
        "cdnr_credit": cdnr_credit,
        "cdnr_debit": cdnr_debit,
        "cdnur_credit": cdnur_dom_credit,   # domestic unregistered notes only
        "cdnur_debit": cdnur_dom_debit,
        "cdnur_exp_credit": cdnur_exp_credit,  # export-related unregistered notes
        "cdnura_credit": cdnura_credit,
        "cdnura_debit": cdnura_debit,
        "cdnur_exp_debit": cdnur_exp_debit,
        "exp_wpay": exp_by_type["WPAY"],
        "exp_wopay": exp_by_type["WOPAY"],
        "b2cs": b2cs,
        "hsn_summary": hsn_rows,
        "hsn_schema_variant": hsn_schema_variant,
        "doc_issue_summary": _parse_doc_issue_summary(d.get("doc_issue", {})),
        "filed_on": d.get("fil_dt"),
        "grand_total_declared": d.get("cur_gt", d.get("gt", 0)),
        "total_taxable_value": total_txval,
        "total_tax": total_tax,
        "unmapped_keys": unmapped,
    }


if __name__ == "__main__":
    import sys
    print(json.dumps(parse_r1(sys.argv[1]), indent=2))

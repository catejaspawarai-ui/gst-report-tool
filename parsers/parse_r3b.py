"""
Parser for GSTR-3B JSON (portal export, inside the R3B zip).
Extracts summary fields matching the GSTR-3B Summary report card.
"""
import json
from parsers.unmapped import find_unmapped_keys


def _sum_head(d):
    """Sum iamt/camt/samt/csamt from a single {iamt,camt,samt,csamt} dict."""
    return {
        "igst": d.get("iamt", 0) or 0,
        "cgst": d.get("camt", 0) or 0,
        "sgst": d.get("samt", 0) or 0,
        "cess": d.get("csamt", 0) or 0,
    }


def _sum_list(items, keys=("iamt", "camt", "samt", "csamt")):
    out = {"igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
    for it in items:
        out["igst"] += it.get("iamt", 0) or 0
        out["cgst"] += it.get("camt", 0) or 0
        out["sgst"] += it.get("samt", 0) or 0
        out["cess"] += it.get("csamt", 0) or 0
    return out


def parse_r3b(filepath):
    with open(filepath, "r", encoding="utf-8") as f:
        d = json.load(f)

    gstin = d.get("gstin")
    period = d.get("ret_period")  # MMYYYY

    sup = d.get("sup_details", {})
    outward_taxable = _sum_head(sup.get("osup_det", {}))
    outward_taxable["txval"] = sup.get("osup_det", {}).get("txval", 0) or 0

    zero_rated = _sum_head(sup.get("osup_zero", {}))
    zero_rated["txval"] = sup.get("osup_zero", {}).get("txval", 0) or 0

    nil_exempt = _sum_head(sup.get("osup_nil_exmp", {}))
    nil_exempt["txval"] = sup.get("osup_nil_exmp", {}).get("txval", 0) or 0

    rcm_inward = _sum_head(sup.get("isup_rev", {}))
    rcm_inward["txval"] = sup.get("isup_rev", {}).get("txval", 0) or 0

    non_gst = sup.get("osup_nongst", {}).get("txval", 0) or 0

    itc = d.get("itc_elg", {})
    itc_by_type = {}
    for row in itc.get("itc_avl", []):
        itc_by_type[row.get("ty")] = _sum_head(row)
    itc_available = _sum_list(itc.get("itc_avl", []))

    itc_rev_by_type = {}
    for row in itc.get("itc_rev", []):
        itc_rev_by_type[row.get("ty")] = _sum_head(row)
    itc_reversed = _sum_list(itc.get("itc_rev", []))

    itc_inelg_by_type = {}
    for row in itc.get("itc_inelg", []):
        itc_inelg_by_type[row.get("ty")] = _sum_head(row)
    itc_ineligible = _sum_list(itc.get("itc_inelg", []))

    itc_net = _sum_head(itc.get("itc_net", {}))

    tt_val = d.get("tt_val", {})

    # Section 9(5) -- supplies notified where the e-commerce operator pays tax
    eco = d.get("eco_dtls", {})
    eco_supplies = _sum_head(eco.get("eco_sup", {}))
    eco_supplies["txval"] = eco.get("eco_sup", {}).get("txval", 0) or 0

    # Interest & late fee due (Table 5.1/6.1) -- usually zero, but wire it in
    # so it surfaces the moment there's a real delay.
    intr_ltfee = d.get("intr_ltfee", {})
    interest_due = _sum_head(intr_ltfee.get("intr_details", {}))
    late_fee_due = _sum_head(intr_ltfee.get("ltfee_details", {}))

    # Payment cross-utilization matrix: which ITC head (or cash) discharged each
    # liability head. Sourced from taxpayble.returnsDbCdredList.tax_paid --
    # pd_by_itc gives the ITC-to-liability offset, pd_by_cash gives the cash leg.
    tax_paid = d.get("taxpayble", {}).get("returnsDbCdredList", {}).get("tax_paid", {})
    pd_by_itc = tax_paid.get("pd_by_itc", [])
    pd_by_cash = tax_paid.get("pd_by_cash", [])

    def _sum_field(entries, field):
        return sum(e.get(field, 0) or 0 for e in entries)

    def _sum_cash_head(entries, head):
        return sum((e.get(head, {}).get("tx", 0) or 0) for e in entries)

    payment_matrix = {
        "IGST": {
            "via_igst_itc": _sum_field(pd_by_itc, "igst_igst_amt"),
            "via_cgst_itc": _sum_field(pd_by_itc, "igst_cgst_amt"),
            "via_sgst_itc": _sum_field(pd_by_itc, "igst_sgst_amt"),
            "via_cash": _sum_cash_head(pd_by_cash, "igst"),
        },
        "CGST": {
            "via_igst_itc": _sum_field(pd_by_itc, "cgst_igst_amt"),
            "via_cgst_itc": _sum_field(pd_by_itc, "cgst_cgst_amt"),
            "via_cash": _sum_cash_head(pd_by_cash, "cgst"),
        },
        "SGST": {
            "via_igst_itc": _sum_field(pd_by_itc, "sgst_igst_amt"),
            "via_sgst_itc": _sum_field(pd_by_itc, "sgst_sgst_amt"),
            "via_cash": _sum_cash_head(pd_by_cash, "sgst"),
        },
        "CESS": {
            "via_cess_itc": _sum_field(pd_by_itc, "cess_cess_amt"),
            "via_cash": _sum_cash_head(pd_by_cash, "cess"),
        },
    }

    return {
        "gstin": gstin,
        "period": period,
        "outward_taxable_value": outward_taxable,
        "zero_rated": zero_rated,
        "nil_exempt": nil_exempt,
        "rcm_inward_liability": rcm_inward,
        "non_gst_outward_txval": non_gst,
        "eco_supplies_9_5": eco_supplies,  # Section 9(5) notified supplies (e-commerce operator pays tax)
        "itc_available": itc_available,
        "itc_by_type": itc_by_type,  # IMPG / IMPS / ISRC / ISD / OTH, each {igst,cgst,sgst,cess}
        "itc_reversed": itc_reversed,
        "itc_reversed_by_type": itc_rev_by_type,  # RUL (as per rules) / OTH (others)
        "itc_ineligible": itc_ineligible,
        "itc_ineligible_by_type": itc_inelg_by_type,  # RUL (Sec 17(5)) / OTH (others)
        "itc_net": itc_net,
        "tax_paid_cash": tt_val.get("tt_csh_pd", 0) or 0,
        "tax_paid_itc": tt_val.get("tt_itc_pd", 0) or 0,
        "interest_due": interest_due,
        "late_fee_due": late_fee_due,
        "payment_matrix": payment_matrix,
        # Total outward taxable value + tax, used for 3B-vs-1 comparison
        "total_liability_txval": outward_taxable["txval"] + zero_rated["txval"],
        "total_liability_tax": (
            outward_taxable["igst"] + outward_taxable["cgst"] + outward_taxable["sgst"] + outward_taxable["cess"]
            + zero_rated["igst"] + zero_rated["cgst"] + zero_rated["sgst"] + zero_rated["cess"]
        ),
        "unmapped_keys": find_unmapped_keys(d, {
            "gstin", "ret_period", "sup_details", "itc_elg", "tt_val", "eco_dtls",
            "taxpayble", "intr_ltfee",
        }),
    }


if __name__ == "__main__":
    import sys
    print(json.dumps(parse_r3b(sys.argv[1]), indent=2))

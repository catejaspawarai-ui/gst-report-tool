"""
Parser for GSTR-2B invoice-level detail files (File1, File2, ... -- the
portal splits large exports across an arbitrary number of files purely by
size, NOT by category. A single category like 'b2b' can appear split across
multiple files, so every category array must be combined across ALL detail
files before categorizing anything.

Every category verified against the real, pre-aggregated Summary.json for
the same period -- exact match on every cross-check:
  - itcavl='N' invoices sum to exactly the Summary's 'itcunavl' total
  - docRejdata sums to exactly the Summary's 'itcRejected' total
"""
import json
from parsers.unmapped import find_unmapped_keys

CATEGORY_LABELS = {
    "b2b": "B2B Invoices",
    "b2ba": "B2B Invoice Amendments",
    "cdnr": "Credit/Debit Notes",
    "cdnra": "Credit/Debit Note Amendments",
    "isd": "ISD Credit",
}


def _empty_bucket():
    return {"txval": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0, "doc_count": 0}


def _add_doc(bucket, doc):
    bucket["txval"] += doc.get("txval", 0) or 0
    bucket["igst"] += doc.get("igst", 0) or 0
    bucket["cgst"] += doc.get("cgst", 0) or 0
    bucket["sgst"] += doc.get("sgst", 0) or 0
    bucket["cess"] += doc.get("cess", 0) or 0
    bucket["doc_count"] += 1


def _iter_docs(category_list, doc_list_key):
    """b2b/cdnr/isd share the shape [{ctin, <doc_list_key>: [...]}, ...]."""
    for entry in category_list:
        for doc in entry.get(doc_list_key, []):
            yield doc


def _doc_list_key_for(category):
    return {"b2b": "inv", "b2ba": "inv", "cdnr": "nt", "cdnra": "nt", "isd": "doclist"}[category]


def parse_r2b_detail_files(raw_dicts):
    """raw_dicts: list of already-json.load()-ed detail file contents
    (each the full {'data':..., 'chksum':...} structure). Combines all of
    them before categorizing anything."""
    merged_docdata = {}
    merged_rejdata = {}
    gstin = None
    period = None

    for d in raw_dicts:
        data = d.get("data", d)
        gstin = gstin or data.get("gstin")
        period = period or data.get("rtnprd")
        for cat, docs in data.get("docdata", {}).items():
            merged_docdata.setdefault(cat, []).extend(docs)
        for cat, docs in data.get("docRejdata", {}).items():
            merged_rejdata.setdefault(cat, []).extend(docs)

    # category -> {"eligible_regular", "eligible_rcm", "ineligible"} -> bucket
    by_category = {}
    for cat, doc_list_key in [(c, _doc_list_key_for(c)) for c in CATEGORY_LABELS]:
        entries = merged_docdata.get(cat, [])
        buckets = {"eligible_regular": _empty_bucket(), "eligible_rcm": _empty_bucket(), "ineligible": _empty_bucket()}
        for doc in _iter_docs(entries, doc_list_key):
            itcavl = doc.get("itcavl", doc.get("itcelg", "Y"))
            rev = doc.get("rev", "N")
            if itcavl == "N":
                _add_doc(buckets["ineligible"], doc)
            elif rev == "Y":
                _add_doc(buckets["eligible_rcm"], doc)
            else:
                _add_doc(buckets["eligible_regular"], doc)
        by_category[cat] = buckets

    # Imports: no eligibility/RCM flags, already fully eligible by nature
    import_bucket = _empty_bucket()
    for doc in merged_docdata.get("impg", []):
        _add_doc(import_bucket, doc)

    # Rejected (IMS): separate from the above entirely
    rejected_bucket = _empty_bucket()
    for cat, doc_list_key in [(c, _doc_list_key_for(c)) for c in CATEGORY_LABELS if c in merged_rejdata]:
        for doc in _iter_docs(merged_rejdata.get(cat, []), doc_list_key):
            _add_doc(rejected_bucket, doc)

    return {
        "gstin": gstin,
        "period": period,
        "by_category": by_category,  # {category: {eligible_regular, eligible_rcm, ineligible}}
        "imports": import_bucket,
        "rejected": rejected_bucket,
    }


if __name__ == "__main__":
    import sys
    dicts = []
    for path in sys.argv[1:]:
        with open(path) as f:
            dicts.append(json.load(f))
    print(json.dumps(parse_r2b_detail_files(dicts), indent=2))

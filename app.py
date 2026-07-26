"""
GST Annual Report Generator -- free web tool.
User drags in ALL their GSTR-1, GSTR-2B Summary, and GSTR-3B files for a
whole financial year at once, and downloads one Excel report with every
month as a column.

Files can be uploaded in any order, in any combination of .json/.zip -- each
file's return type AND period are detected from its own JSON content (not
filename), since public users won't follow any particular naming convention.
"""
import os
import sys
import json as _json
import uuid
import zipfile
import shutil
import subprocess
import hashlib
import tempfile
import traceback
from flask import Flask, request, render_template, send_file, flash, redirect, url_for, make_response

sys.path.insert(0, os.path.dirname(__file__))
from parsers.parse_r1 import parse_r1
from parsers.parse_r2b import parse_r2b_summary
from parsers.parse_r2b_detail import parse_r2b_detail_files
from parsers.parse_r3b import parse_r3b
from comparator import compare_3b_vs_1, compare_3b_vs_2b
from report_builder import build_workbook
from state_codes import state_name_from_gstin, safe_filename_component

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-change-in-production")
app.config["MAX_CONTENT_LENGTH"] = 150 * 1024 * 1024  # 150MB -- a full year across 6+ GSTINs adds up

OUTPUT_DIR = os.path.join(tempfile.gettempdir(), "gst_report_webapp_outputs")
os.makedirs(OUTPUT_DIR, exist_ok=True)

LABELS = {"gstr1": "GSTR-1", "gstr2b": "GSTR-2B Summary", "gstr3b": "GSTR-3B"}


def _recalculate_formulas(xlsx_path):
    """Forces Excel formulas (e.g. the Total column) to have cached values,
    so they display correctly even if the viewer doesn't auto-calculate on
    open (some Excel installs default to Manual calculation; some non-Excel
    viewers never calculate at all). Uses LibreOffice if it's available on
    this machine; silently skips otherwise -- Excel will still calculate the
    formulas itself on open in the vast majority of cases, so this is a
    belt-and-braces step, not a hard requirement."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return  # not available on this machine -- rely on Excel's own auto-calc
    try:
        with tempfile.TemporaryDirectory() as recalc_tmpdir:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "xlsx", "--outdir", recalc_tmpdir, xlsx_path],
                timeout=60, capture_output=True, check=True,
            )
            converted_path = os.path.join(recalc_tmpdir, os.path.basename(xlsx_path))
            if os.path.exists(converted_path):
                shutil.move(converted_path, xlsx_path)
    except Exception:
        pass  # recalculation is a nice-to-have -- never let it break report generation


def _extract_json_paths(uploaded_file, tmpdir):
    """Saves an uploaded file and returns a list of JSON file paths found in it
    (a .json file yields itself; a .zip yields every .json inside it).

    Each upload gets its own unique subdirectory before extraction -- the GST
    portal reuses the same internal filename (e.g.
    'returns_<downloaddate>_R1_<gstin>_offline_others_0.json') across every
    month's export, so extracting multiple zips into one shared folder would
    silently overwrite earlier months with later ones."""
    fname = uploaded_file.filename
    if not fname:
        return []
    upload_dir = os.path.join(tmpdir, uuid.uuid4().hex)
    os.makedirs(upload_dir, exist_ok=True)
    raw_path = os.path.join(upload_dir, fname)
    uploaded_file.save(raw_path)

    if fname.lower().endswith(".zip"):
        try:
            with zipfile.ZipFile(raw_path) as z:
                json_names = [n for n in z.namelist() if n.lower().endswith(".json")]
                paths = []
                for n in json_names:
                    z.extract(n, upload_dir)
                    paths.append(os.path.join(upload_dir, n))
                return paths
        except zipfile.BadZipFile:
            raise ValueError(f"'{fname}' doesn't look like a valid zip file. Please re-download it.")
    elif fname.lower().endswith(".json"):
        return [raw_path]
    else:
        raise ValueError(f"'{fname}' isn't a .json or .zip file -- skipping it.")


def _detect_return_type_and_period(data):
    """Identifies return type AND period from a JSON blob's own content.
    Returns (return_type, gstin, period) or (None, None, None) if unrecognized."""
    if not isinstance(data, dict):
        return None, None, None
    if "sup_details" in data and "itc_elg" in data:
        return "gstr3b", data.get("gstin"), data.get("ret_period")
    # GSTR-1: GSTN's own tool allows any subset of sections to be filled in --
    # a file with only, say, b2cl/at/ecom and no b2b at all is completely valid
    # and real (confirmed against an actual government-generated export). So we
    # detect on ANY recognized GSTR-1 section key being present, not requiring
    # b2b+b2cs+gt/cur_gt all together.
    gstr1_section_keys = {
        "b2b", "b2ba", "b2cl", "b2cla", "b2cs", "b2csa", "cdnr", "cdnra",
        "cdnur", "cdnura", "exp", "expa", "nil", "at", "ata", "atadj", "atadja",
        "txpd", "txpda", "hsn", "doc_issue", "eco", "ecoa", "ecom", "ecoma",
    }
    if "gstin" in data and "fp" in data and gstr1_section_keys & set(data.keys()):
        return "gstr1", data.get("gstin"), data.get("fp")
    if "data" in data and isinstance(data["data"], dict) and "itcsumm" in data["data"]:
        inner = data["data"]
        return "gstr2b_summary", inner.get("gstin"), inner.get("rtnprd")
    if "data" in data and isinstance(data["data"], dict) and "docdata" in data["data"]:
        inner = data["data"]
        return "gstr2b_detail", inner.get("gstin"), inner.get("rtnprd")
    return None, None, None


def _parse_with_friendly_errors(parser_fn, path, label):
    try:
        return parser_fn(path)
    except _json.JSONDecodeError:
        raise ValueError(f"A {label} file doesn't look like valid JSON. "
                          f"Make sure it's downloaded directly from the GST portal.")
    except KeyError as e:
        raise ValueError(f"A {label} file is missing an expected field ({e}). "
                          f"It may not be the return type it looks like -- please check.")


def _dedupe_by_content(paths):
    """Given multiple file paths detected for the same GSTIN+period+return-type,
    checks whether they're byte-identical (e.g. the same file accidentally
    downloaded/uploaded twice -- Windows often names the second copy with a
    ' (1)' suffix, which is exactly this case) versus genuinely different
    content (a real conflict, or possibly an actual large-data split this tool
    doesn't merge). Returns (unique_paths, was_deduped)."""
    hashes = {}
    for p in paths:
        with open(p, "rb") as f:
            h = hashlib.sha256(f.read()).hexdigest()
        hashes.setdefault(h, []).append(p)
    unique_paths = [group[0] for group in hashes.values()]
    return unique_paths, len(unique_paths) < len(paths)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    try:
        uploaded_files = request.files.getlist("files")
        uploaded_files = [f for f in uploaded_files if f and f.filename]
        if not uploaded_files:
            flash("Please add your GSTR-1, GSTR-2B Summary, GSTR-2B Detail (File1/File2/...), and GSTR-3B files before generating.")
            return redirect(url_for("index"))

        with tempfile.TemporaryDirectory() as tmpdir:
            # by_key[(gstin, period, return_type)] -> list of file paths (normally 1;
            # >1 means either an accidental duplicate upload or a large-data split
            # file we don't yet know how to merge -- either way we flag it rather
            # than silently combine or silently drop data)
            by_key = {}
            ignored = []

            for uf in uploaded_files:
                try:
                    json_paths = _extract_json_paths(uf, tmpdir)
                except ValueError as e:
                    flash(str(e))
                    return redirect(url_for("index"))

                for path in json_paths:
                    try:
                        with open(path, "r", encoding="utf-8") as f:
                            data = _json.load(f)
                    except _json.JSONDecodeError:
                        ignored.append(os.path.basename(path))
                        continue

                    rtype, gstin, period = _detect_return_type_and_period(data)
                    if rtype is None or not gstin or not period:
                        ignored.append(os.path.basename(path))
                        continue
                    by_key.setdefault((gstin, period, rtype), []).append(path)

            if not by_key:
                flash("Couldn't recognize any GSTR-1, GSTR-2B Summary, GSTR-2B Detail, or GSTR-3B files in what you uploaded. "
                      "Make sure they're downloaded directly from the GST portal.")
                return redirect(url_for("index"))

            gstins_seen = set(k[0] for k in by_key.keys())
            if len(gstins_seen) > 1:
                flash(f"Found files for {len(gstins_seen)} different GSTINs in one upload. "
                      f"Please generate one report per GSTIN at a time.")
                return redirect(url_for("index"))
            gstin = gstins_seen.pop()

            # Group by period, keep only periods with a complete set of files
            periods = sorted(set(k[1] for k in by_key.keys()))
            notes = []
            gstr3b_records, gstr1_records, gstr2b_records, gstr2b_detail_records = [], [], [], []
            cmp_3b1, cmp_3b2b = [], []

            for period in periods:
                r3b_files, r3b_deduped = _dedupe_by_content(by_key.get((gstin, period, "gstr3b"), []))
                r1_files, r1_deduped = _dedupe_by_content(by_key.get((gstin, period, "gstr1"), []))
                r2b_files, r2b_deduped = _dedupe_by_content(by_key.get((gstin, period, "gstr2b_summary"), []))
                # GSTR-2B detail files (File1/File2/...) legitimately split data by
                # size across an arbitrary number of files -- these get MERGED,
                # not deduped-or-flagged, unlike the other single-file return types.
                r2b_detail_files = by_key.get((gstin, period, "gstr2b_detail"), [])

                deduped_types = [t for t, was_deduped in
                                 [("GSTR-1", r1_deduped), ("GSTR-2B Summary", r2b_deduped), ("GSTR-3B", r3b_deduped)]
                                 if was_deduped]
                if deduped_types:
                    notes.append(f"{period}: {', '.join(deduped_types)} had duplicate uploads of the identical "
                                 f"file (e.g. a re-download) -- used one copy, no data lost.")

                multi_part = [(t, fs) for t, fs in
                              [("GSTR-1", r1_files), ("GSTR-2B Summary", r2b_files), ("GSTR-3B", r3b_files)]
                              if len(fs) > 1]
                if multi_part:
                    parts_desc = ", ".join(f"{t} ({len(fs)} genuinely different files)" for t, fs in multi_part)
                    notes.append(f"{period}: skipped -- multiple DIFFERENT files detected for {parts_desc}. "
                                 f"This looks like a real large-data split file this tool doesn't yet merge -- "
                                 f"only complete, single-file periods are included below.")
                    continue

                missing = [t for t, fs in [("GSTR-1", r1_files), ("GSTR-2B Summary", r2b_files),
                                            ("GSTR-3B", r3b_files), ("GSTR-2B Detail (File1/File2/...)", r2b_detail_files)]
                           if not fs]
                if missing:
                    notes.append(f"{period}: skipped -- missing {', '.join(missing)}.")
                    continue

                r3b = _parse_with_friendly_errors(parse_r3b, r3b_files[0], "GSTR-3B")
                r1 = _parse_with_friendly_errors(parse_r1, r1_files[0], "GSTR-1")
                r2b = _parse_with_friendly_errors(parse_r2b_summary, r2b_files[0], "GSTR-2B Summary")
                r2b_detail_dicts = []
                for p in r2b_detail_files:
                    with open(p, "r", encoding="utf-8") as f:
                        r2b_detail_dicts.append(_json.load(f))
                r2b_detail = parse_r2b_detail_files(r2b_detail_dicts)

                gstr3b_records.append(r3b)
                gstr1_records.append(r1)
                gstr2b_records.append(r2b)
                gstr2b_detail_records.append(r2b_detail)
                cmp_3b1.append(compare_3b_vs_1(r3b, r1))
                cmp_3b2b.append(compare_3b_vs_2b(r3b, r2b))

            if not gstr3b_records:
                flash("None of the periods you uploaded had a complete set of GSTR-1 + GSTR-2B + GSTR-3B. " +
                      (" ".join(notes) if notes else ""))
                return redirect(url_for("index"))

            included_periods = sorted(set(r["period"] for r in gstr3b_records))
            included_labels = [f"{p[:2]}/{p[2:]}" for p in included_periods]
            notes.insert(0, f"Included periods ({len(included_periods)}): {', '.join(included_labels)}")

            hsn_variants = set(r.get("hsn_schema_variant", "none") for r in gstr1_records) - {"none"}
            if len(hsn_variants) > 1:
                notes.append("Note: the GST portal's HSN JSON structure differed across the periods you uploaded "
                             "(some use separate B2B/B2C arrays, others a combined format) -- both are handled "
                             "correctly, this is just so the difference doesn't look like a data gap if you "
                             "compare periods closely.")

            state = state_name_from_gstin(gstin).upper()
            company_name = request.form.get("company_name", "").strip() or None
            financial_year = request.form.get("financial_year", "").strip() or None

            out_filename = safe_filename_component(f"GST_Report_{gstin.upper()}_{state}_{uuid.uuid4().hex[:6]}.xlsx")
            out_path = os.path.join(OUTPUT_DIR, out_filename)

            unmapped_info = []
            for records, label in [(gstr3b_records, "GSTR-3B"), (gstr1_records, "GSTR-1"), (gstr2b_records, "GSTR-2B")]:
                for r in records:
                    for u in r.get("unmapped_keys", []):
                        unmapped_info.append({
                            "gstin": r["gstin"], "period": r["period"], "return_type": label,
                            "key": u["key"], "preview": u["preview"],
                        })

            build_workbook(gstr3b_records, gstr1_records, gstr2b_records, cmp_3b1, cmp_3b2b, out_path,
                            company_name=company_name, financial_year=financial_year, data_notes=notes,
                            unmapped_info=unmapped_info, gstr2b_detail_records=gstr2b_detail_records)
            _recalculate_formulas(out_path)

            download_name = safe_filename_component(f"GST_Report_{gstin.upper()}_{state}.xlsx")
            response = make_response(send_file(out_path, as_attachment=True, download_name=download_name))
            token = request.form.get("download_token", "")
            if token:
                response.set_cookie("download_token", token, max_age=60)
            return response

    except ValueError as e:
        flash(str(e))
        return redirect(url_for("index"))
    except Exception:
        traceback.print_exc()
        flash("Something went wrong reading your files. Double-check they're the right "
              "JSON exports from the GST portal, then try again.")
        return redirect(url_for("index"))


if __name__ == "__main__":
    # SECURITY: debug mode must stay OFF for any public/hosted deployment --
    # Werkzeug's debugger allows remote code execution if left on. Only set
    # FLASK_DEBUG=true in your own local environment while testing.
    debug_mode = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    app.run(debug=debug_mode, use_reloader=False, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))

import io
import re
from decimal import Decimal, InvalidOperation

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple

import streamlit as st

# CONFIG (toujours en premier)
st.set_page_config(
    page_title="Athina Logistics Tool",
    page_icon="logo.png",
    layout="wide"
)

# 👇 ICI tu mets ton logo sidebar
st.sidebar.image("assets/logo.png", width=200)
st.sidebar.markdown("### Athina Logistics")
st.sidebar.caption("Global Access")

# TON APP
st.title("HS Code Analyzer from Invoices")

st.write("Upload your invoices...")

# =========================
# Helpers
# =========================
COUNTRY_CODE_RE = re.compile(r"^[A-Z]{2}$")
_REF_RE = re.compile(
    r"^\s*=\s*(?:(?P<sheet>'[^']+'|[A-Za-z0-9 _.-]+)!)?\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)\s*$"
)

def to_decimal(val):
    if val is None:
        return None
    s = str(val).replace(",", ".").strip()
    if not s:
        return None
    try:
        return Decimal(s)
    except InvalidOperation:
        return None

def sheet_by_name_ci(wb, wanted):
    norm = wanted.strip().lower()
    for name in wb.sheetnames:
        if name.strip().lower() == norm:
            return wb[name]
    return None

def find_sum_row(ws, start_row=1, label_col="B"):
    if ws is None:
        return None
    for r in range(start_row, 10000):
        v = ws[f"{label_col}{r}"].value
        if not v:
            continue
        s = str(v).strip().upper()
        if s == "SUM" or re.match(r"^SUM\b", s):
            return r
    return None

def is_cell_in_merged(ws, row, col):
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return True
    return False

def get_merged_value(ws, cell_ref):
    cell = ws[cell_ref]
    if cell.value is not None:
        return cell.value

    r, c = coordinate_to_tuple(cell_ref)
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col:
            return ws.cell(row=rng.min_row, column=rng.min_col).value
    return None

def get_effective_cell_value(ws, row, col):
    val = ws.cell(row=row, column=col).value
    if val is not None:
        return val

    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return ws.cell(row=rng.min_row, column=rng.min_col).value
    return None

def contains_chinese(text):
    if not isinstance(text, str):
        return False
    return re.search(r'[\u4e00-\u9fff]', text) is not None

def _col_to_idx(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n

def resolve_simple_formula(wbX, ws_current, formula, depth=0):
    if depth > 10:
        return None
    if not isinstance(formula, str) or not formula.strip().startswith("="):
        return None
    m = _REF_RE.match(formula)
    if not m:
        return None
    sheet = m.group("sheet")
    col = m.group("col")
    row = int(m.group("row"))

    if sheet:
        sheet = sheet.strip()
        if sheet.startswith("'") and sheet.endswith("'"):
            sheet = sheet[1:-1]
        if sheet not in wbX.sheetnames:
            return None
        ws = wbX[sheet]
    else:
        ws = ws_current

    c = _col_to_idx(col)
    v = ws.cell(row=row, column=c).value
    if isinstance(v, str) and v.strip().startswith("="):
        return resolve_simple_formula(wbX, ws, v, depth + 1)
    return v

def header_value(wb_formula, ws_dataonly, ws_formula, ref):
    v = ws_dataonly[ref].value
    if v is not None and str(v).strip() != "":
        return v
    if ws_formula is None:
        return v
    vf = ws_formula[ref].value
    if isinstance(vf, str) and vf.strip().startswith("="):
        rv = resolve_simple_formula(wb_formula, ws_formula, vf)
        return rv
    return vf

def q2(x):
    if x is None:
        return None
    try:
        return x.quantize(Decimal("0.01"))
    except Exception:
        return None

def check_file(uploaded_file):
    errors = []
    warnings = []
    info = {}

    raw = uploaded_file.getvalue()
    bio1 = io.BytesIO(raw)
    bio2 = io.BytesIO(raw)

    wb = load_workbook(bio1, data_only=True)
    wb_f = load_workbook(bio2, data_only=False)

    ws_inv = sheet_by_name_ci(wb, "INVOICE")
    ws_pack = sheet_by_name_ci(wb, "PACKING LIST")
    ws_inv_f = sheet_by_name_ci(wb_f, "INVOICE")
    ws_pack_f = sheet_by_name_ci(wb_f, "PACKING LIST")

    fname = uploaded_file.name
    fname_no_ext = fname.rsplit(".", 1)[0]

    if ws_inv is None:
        errors.append("Missing sheet: INVOICE")
        return {
            "file": fname,
            "status": "ERROR",
            "errors": errors,
            "warnings": warnings,
            "info": info,
            "error_count": len(errors),
            "warning_count": len(warnings),
        }

    if ws_pack is None:
        warnings.append("Missing sheet: PACKING LIST (cross-checks limited)")

    def cell(ws, ref):
        return ws[ref].value if ws is not None else None

    # A) Base checks
    inv_a2 = header_value(wb_f, ws_inv, ws_inv_f, "A2")
    inv_c4 = header_value(wb_f, ws_inv, ws_inv_f, "C4")
    if inv_a2 != inv_c4:
        errors.append(f"INVOICE header mismatch: A2='{inv_a2}' vs C4='{inv_c4}'")

    if ws_pack is not None:
        pack_a2 = header_value(wb_f, ws_pack, ws_pack_f, "A2")
        if not (inv_a2 == inv_c4 == pack_a2):
            errors.append(
                f"Header mismatch across sheets: INVOICE A2='{inv_a2}', C4='{inv_c4}', PACKING A2='{pack_a2}'"
            )

    # B) Filename = C5 = J4 = PACK.B4
    inv_c5_val = header_value(wb_f, ws_inv, ws_inv_f, "C5")
    inv_j4_val = header_value(wb_f, ws_inv, ws_inv_f, "J4")
    pack_b4_val = header_value(wb_f, ws_pack, ws_pack_f, "B4") if ws_pack is not None else None

    inv_c5 = str(inv_c5_val).strip() if inv_c5_val else ""
    inv_j4 = str(inv_j4_val).strip() if inv_j4_val else ""
    pack_b4 = str(pack_b4_val).strip() if pack_b4_val else ""

    if ws_pack is not None:
        if not (fname_no_ext == inv_c5 == inv_j4 == pack_b4):
            errors.append(
                f"Filename mismatch: file='{fname_no_ext}', INVOICE C5='{inv_c5}', INVOICE J4='{inv_j4}', PACKING B4='{pack_b4}'"
            )
    else:
        if not (fname_no_ext == inv_c5 == inv_j4):
            errors.append(
                f"Filename mismatch: file='{fname_no_ext}', INVOICE C5='{inv_c5}', INVOICE J4='{inv_j4}'"
            )

    # C) Key fields
    j13 = str(cell(ws_inv, "J13")).strip().upper() if cell(ws_inv, "J13") else ""
    if j13 != "EUR":
        errors.append(f"J13 must be 'EUR' (found: {cell(ws_inv, 'J13')})")

    j14 = str(cell(ws_inv, "J14")).strip().upper() if cell(ws_inv, "J14") else ""
    if j14 != "CIF":
        warnings.append(f"J14 should be 'CIF' (found: {cell(ws_inv, 'J14')})")

    try:
        j16 = Decimal(str(cell(ws_inv, "J16")).strip()) if cell(ws_inv, "J16") else None
    except InvalidOperation:
        j16 = None
    if j16 != Decimal(4200):
        errors.append(f"J16 must be 4200 (found: {cell(ws_inv, 'J16')})")

    for r in range(11, 16):
        v = cell(ws_inv, f"C{r}")
        if v is None or str(v).strip() == "":
            errors.append(f"C{r} is empty")

    for r in (16, 17):
        v = cell(ws_inv, f"C{r}")
        s = str(v).strip().upper() if v is not None else ""
        if not COUNTRY_CODE_RE.match(s):
            errors.append(f"C{r} must be a 2-letter country code (found: '{v}')")

    for r in (11, 13, 15):
        v = cell(ws_inv, f"J{r}")
        if v is None or str(v).strip() == "":
            errors.append(f"J{r} is empty")

    c9 = str(cell(ws_inv, "C9")).strip().upper() if cell(ws_inv, "C9") else ""
    if c9 != "CN":
        errors.append(f"C9 must be 'CN' (found: {cell(ws_inv, 'C9')})")

    # D) SUM rows
    inv_sum_row = find_sum_row(ws_inv, start_row=19, label_col="B")
    pack_sum_row = find_sum_row(ws_pack, start_row=6, label_col="B") if ws_pack else None

    if inv_sum_row is None:
        errors.append("INVOICE SUM row not found in column B")
    if ws_pack is not None and pack_sum_row is None:
        errors.append("PACKING LIST SUM row not found in column B")

    # D-CH) Chinese characters
    if inv_sum_row:
        inv_ch_rows = []
        for row in range(20, inv_sum_row):
            cell_value = ws_inv[f"B{row}"].value
            if contains_chinese(cell_value):
                inv_ch_rows.append(f"B{row}")
        if inv_ch_rows:
            errors.append(f"Chinese characters found in INVOICE descriptions: {', '.join(inv_ch_rows[:20])}" + (" ..." if len(inv_ch_rows) > 20 else ""))

    if ws_pack and pack_sum_row:
        pack_ch_rows = []
        for row in range(6, pack_sum_row):
            cell_value = ws_pack[f"B{row}"].value
            if contains_chinese(cell_value):
                pack_ch_rows.append(f"B{row}")
        if pack_ch_rows:
            errors.append(f"Chinese characters found in PACKING LIST descriptions: {', '.join(pack_ch_rows[:20])}" + (" ..." if len(pack_ch_rows) > 20 else ""))

    # D0) Merged cells forbidden in weight zones
    if inv_sum_row:
        merged_inv = []
        for r in range(20, inv_sum_row):
            if is_cell_in_merged(ws_inv, r, 10):
                merged_inv.append(f"J{r}")
            if is_cell_in_merged(ws_inv, r, 11):
                merged_inv.append(f"K{r}")
        if merged_inv:
            errors.append(
                "Merged cells not allowed in INVOICE J/K area: "
                + ", ".join(merged_inv[:20])
                + (" ..." if len(merged_inv) > 20 else "")
            )

    if ws_pack and pack_sum_row:
        merged_pack = []
        for r in range(6, pack_sum_row):
            if is_cell_in_merged(ws_pack, r, 9):
                merged_pack.append(f"I{r}")
            if is_cell_in_merged(ws_pack, r, 10):
                merged_pack.append(f"J{r}")
        if merged_pack:
            errors.append(
                "Merged cells not allowed in PACKING LIST I/J area: "
                + ", ".join(merged_pack[:20])
                + (" ..." if len(merged_pack) > 20 else "")
            )

    # D2) G length <= 48
    if inv_sum_row:
        for r in range(20, inv_sum_row):
            val = ws_inv[f"G{r}"].value
            if isinstance(val, str) and len(val.strip()) > 48:
                errors.append("INVOICE column G contains a value longer than 48 characters")
                break

    # D3) INVOICE D and I must be numeric, non-empty, non-zero
    if inv_sum_row:
        di_errors = []
        def check_DI_cell(row, col_letter, col_index):
            if is_cell_in_merged(ws_inv, row, col_index):
                return
            val = ws_inv[f"{col_letter}{row}"].value
            if val is None:
                di_errors.append(f"{col_letter}{row} empty")
                return
            s = str(val).strip().replace(",", ".")
            if s in ("0", "0.0", "0.00"):
                di_errors.append(f"{col_letter}{row} = 0")
                return
            try:
                Decimal(s)
            except InvalidOperation:
                di_errors.append(f"{col_letter}{row} non-numeric ('{val}')")

        for r in range(20, inv_sum_row):
            check_DI_cell(r, "D", 4)
            check_DI_cell(r, "I", 9)

        if di_errors:
            errors.append(
                "Invalid goods values in INVOICE D/I: "
                + "; ".join(di_errors[:20])
                + (" ..." if len(di_errors) > 20 else "")
            )

    # E) Line-by-line net < gross
    if inv_sum_row:
        text_rows = []
        bad_rows = []
        for r in range(20, inv_sum_row):
            j_val = ws_inv[f"J{r}"].value
            k_val = ws_inv[f"K{r}"].value
            if isinstance(j_val, str) or isinstance(k_val, str):
                text_rows.append(r)
                continue
            j_dec = to_decimal(j_val)
            k_dec = to_decimal(k_val)
            if j_dec is not None and k_dec is not None and j_dec >= k_dec:
                bad_rows.append(r)

        if text_rows:
            errors.append(f"Text found in INVOICE J/K line weights: rows {text_rows[:20]}" + (" ..." if len(text_rows) > 20 else ""))
        if bad_rows:
            errors.append(f"Net weight >= gross weight in INVOICE: rows {bad_rows[:20]}" + (" ..." if len(bad_rows) > 20 else ""))

    # F) Total comparisons
    if inv_sum_row and ws_pack and pack_sum_row:
        inv_pieces = to_decimal(ws_inv[f"H{inv_sum_row}"].value)
        inv_net = to_decimal(ws_inv[f"J{inv_sum_row}"].value)
        inv_gross = to_decimal(ws_inv[f"K{inv_sum_row}"].value)
        pack_pieces = to_decimal(ws_pack[f"H{pack_sum_row}"].value)
        pack_net = to_decimal(ws_pack[f"I{pack_sum_row}"].value)
        pack_gross = to_decimal(ws_pack[f"J{pack_sum_row}"].value)
        pack_cartons = to_decimal(get_merged_value(ws_pack, f"G{pack_sum_row}"))

        if inv_net is None or inv_gross is None:
            errors.append(f"INVOICE total net/gross not numeric at SUM row")
        if pack_net is None or pack_gross is None:
            errors.append(f"PACKING LIST total net/gross not numeric at SUM row")

        if inv_pieces != pack_pieces:
            errors.append(f"Total pieces mismatch: INVOICE={inv_pieces}, PACKING={pack_pieces}")
        if inv_net != pack_net:
            errors.append(f"Total net weight mismatch: INVOICE={inv_net}, PACKING={pack_net}")
        if inv_gross != pack_gross:
            errors.append(f"Total gross weight mismatch: INVOICE={inv_gross}, PACKING={pack_gross}")

        if (inv_net is not None and inv_gross is not None) and inv_net > inv_gross:
            errors.append(f"INVOICE total net weight > gross weight ({inv_net} > {inv_gross})")

        if pack_cartons is None:
            errors.append("PACKING LIST total cartons missing or non-numeric")
        else:
            info["cartons"] = str(pack_cartons)

        if pack_gross is not None:
            info["gross_weight"] = str(pack_gross)

        inv_b_values = [str(ws_inv[f"B{r}"].value).strip() if ws_inv[f"B{r}"].value else "" for r in range(20, inv_sum_row)]
        pack_b_values = [str(ws_pack[f"B{r}"].value).strip() if ws_pack[f"B{r}"].value else "" for r in range(6, pack_sum_row)]
        if inv_b_values != pack_b_values:
            errors.append(
                f"Description column B mismatch between INVOICE and PACKING LIST (INV lines={len(inv_b_values)}, PACK lines={len(pack_b_values)})"
            )

        # F3) line-by-line comparison
        line_errors = []
        offset = 14
        for inv_row in range(20, inv_sum_row):
            pack_row = inv_row - offset
            if pack_row < 6 or pack_row >= pack_sum_row:
                continue

            inv_pieces_line = to_decimal(ws_inv[f"H{inv_row}"].value)
            inv_net_line = q2(to_decimal(ws_inv[f"J{inv_row}"].value))
            inv_gross_line = q2(to_decimal(ws_inv[f"K{inv_row}"].value))

            pack_pieces_line = to_decimal(ws_pack[f"H{pack_row}"].value)
            pack_net_line = q2(to_decimal(ws_pack[f"I{pack_row}"].value))
            pack_gross_line = q2(to_decimal(ws_pack[f"J{pack_row}"].value))

            if inv_pieces_line != pack_pieces_line:
                line_errors.append(f"Row {inv_row}: pieces mismatch")
            if inv_net_line != pack_net_line:
                line_errors.append(f"Row {inv_row}: net weight mismatch")
            if inv_gross_line != pack_gross_line:
                line_errors.append(f"Row {inv_row}: gross weight mismatch")

        if line_errors:
            errors.append(
                "Line-by-line differences between INVOICE and PACKING LIST: "
                + "; ".join(line_errors[:25])
                + (" ..." if len(line_errors) > 25 else "")
            )

    # G) PACKING LIST G must not be empty / 0 / non-numeric
    if ws_pack and pack_sum_row:
        bad_rows = []
        for r in range(6, pack_sum_row):
            gv_raw = get_effective_cell_value(ws_pack, r, 7)
            gv_dec = to_decimal(gv_raw)

            if gv_raw is None or str(gv_raw).strip() == "":
                bad_rows.append(f"G{r}=empty")
                continue
            if gv_dec is None:
                bad_rows.append(f"G{r}=non-numeric")
                continue
            if gv_dec == 0:
                bad_rows.append(f"G{r}=0")

        if bad_rows:
            errors.append(
                "Invalid carton values in PACKING LIST column G: "
                + "; ".join(bad_rows[:25])
                + (" ..." if len(bad_rows) > 25 else "")
            )

    status = "OK"
    if errors:
        status = "ERROR"
    elif warnings:
        status = "WARNING"

    return {
        "file": fname,
        "status": status,
        "errors": errors,
        "warnings": warnings,
        "info": info,
        "error_count": len(errors),
        "warning_count": len(warnings),
    }

# =========================
# UI
# =========================
st.title("Final Check - Container Invoice Review")
st.caption("Upload all invoice Excel files for one container. The app will check them one by one and show which files have warnings or errors.")

uploaded_files = st.file_uploader(
    "Upload invoice Excel files (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
)

if uploaded_files:
    results = []
    for f in uploaded_files:
        try:
            results.append(check_file(f))
        except Exception as e:
            results.append({
                "file": f.name,
                "status": "ERROR",
                "errors": [f"File could not be read: {e}"],
                "warnings": [],
                "info": {},
                "error_count": 1,
                "warning_count": 0,
            })

    df = pd.DataFrame([
        {
            "File": r["file"],
            "Status": r["status"],
            "Errors": r["error_count"],
            "Warnings": r["warning_count"],
            "Cartons": r["info"].get("cartons", ""),
            "Gross Weight": r["info"].get("gross_weight", ""),
        }
        for r in results
    ])

    total_files = len(results)
    files_with_errors = sum(1 for r in results if r["error_count"] > 0)
    files_with_warnings_only = sum(1 for r in results if r["error_count"] == 0 and r["warning_count"] > 0)
    files_ok = sum(1 for r in results if r["error_count"] == 0 and r["warning_count"] == 0)

    c1, c2, c3 = st.columns(3)
    c1.metric("Files checked", total_files)
    c2.metric("Files with errors", files_with_errors)
    c3.metric("Files with warnings only", files_with_warnings_only)

    if files_ok:
        st.success(f"{files_ok} file(s) passed with no issues.")
    if files_with_warnings_only:
        st.warning(f"{files_with_warnings_only} file(s) have warnings only.")
    if files_with_errors:
        st.error(f"{files_with_errors} file(s) have errors.")

    st.subheader("Summary")
    st.dataframe(df, use_container_width=True, hide_index=True)

    problem_files = [r for r in results if r["error_count"] > 0 or r["warning_count"] > 0]
    st.subheader("Files with issues")

    if not problem_files:
        st.info("No issues found.")
    else:
        for r in problem_files:
            icon = "❌" if r["error_count"] > 0 else "⚠️"
            title = f"{icon} {r['file']}  |  Errors: {r['error_count']}  |  Warnings: {r['warning_count']}"
            with st.expander(title, expanded=False):
                if r["errors"]:
                    st.markdown("**Errors**")
                    for msg in r["errors"]:
                        st.write(f"- {msg}")
                if r["warnings"]:
                    st.markdown("**Warnings**")
                    for msg in r["warnings"]:
                        st.write(f"- {msg}")

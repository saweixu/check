# -*- coding: utf-8 -*-
# app_sgs_streamlit.py
# Streamlit SGS Generator: upload invoices -> Final Check -> generate T1_SGS.xlsx

import io
import re
import tempfile
from pathlib import Path
from decimal import Decimal, InvalidOperation
from datetime import datetime
from collections import defaultdict, Counter
from copy import copy

import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.utils import column_index_from_string
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

APP_DIR = Path(__file__).resolve().parent
TEMPLATE_NAME = "T1_SGS.xlsx"

# =========================
# BASIC TOOLS
# =========================
def natural_key(name: str):
    stem = Path(name).stem
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", stem)]

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

def to_float(value) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s:
        return 0.0
    s = s.replace(" ", "")
    if s.count(",") and s.count("."):
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0

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

def get_effective_cell_value(ws, row, col):
    val = ws.cell(row=row, column=col).value
    if val is not None:
        return val
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return ws.cell(row=rng.min_row, column=rng.min_col).value
    return None

def get_merged_value(ws, cell_ref):
    cell = ws[cell_ref]
    if cell.value is not None:
        return cell.value
    r, c = coordinate_to_tuple(cell_ref)
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col:
            return ws.cell(row=rng.min_row, column=rng.min_col).value
    return None

def contains_chinese(text):
    return isinstance(text, str) and re.search(r"[\u4e00-\u9fff]", text) is not None

COUNTRY_CODE_RE = re.compile(r"^[A-Z]{2}$")

# =========================
# FINAL CHECK FOR STREAMLIT
# =========================
def final_check_file(file_path: Path):
    errors = []
    warnings = []
    cartons = Decimal(0)
    gross = Decimal(0)

    try:
        wb = load_workbook(file_path, data_only=True)
        wb_f = load_workbook(file_path, data_only=False)
    except Exception as e:
        return {"file": file_path.name, "errors": [f"Impossible ouvrir fichier: {e}"], "warnings": [], "cartons": 0, "gross": 0}

    ws_inv = sheet_by_name_ci(wb, "INVOICE")
    ws_inv_f = sheet_by_name_ci(wb_f, "INVOICE")
    ws_pack = sheet_by_name_ci(wb, "PACKING LIST")
    ws_pack_f = sheet_by_name_ci(wb_f, "PACKING LIST")

    if ws_inv is None:
        return {"file": file_path.name, "errors": ["Feuille INVOICE absente"], "warnings": [], "cartons": 0, "gross": 0}

    if ws_pack is None:
        warnings.append("Feuille PACKING LIST absente")

    ref_re = re.compile(r"^\s*=\s*(?:(?P<sheet>'[^']+'|[A-Za-z0-9 _.-]+)!)?\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d+)\s*$")

    def col_to_idx(col):
        n = 0
        for ch in col:
            n = n * 26 + (ord(ch) - 64)
        return n

    def resolve_simple_formula(wbX, ws_current, formula, depth=0):
        if depth > 10 or not isinstance(formula, str) or not formula.strip().startswith("="):
            return None
        m = ref_re.match(formula)
        if not m:
            return None
        sheet = m.group("sheet")
        col = m.group("col")
        row = int(m.group("row"))
        if sheet:
            sheet = sheet.strip().strip("'")
            if sheet not in wbX.sheetnames:
                return None
            ws = wbX[sheet]
        else:
            ws = ws_current
        v = ws.cell(row=row, column=col_to_idx(col)).value
        if isinstance(v, str) and v.strip().startswith("="):
            return resolve_simple_formula(wbX, ws, v, depth + 1)
        return v

    def header_value(ws_data, ws_formula, ref):
        v = ws_data[ref].value if ws_data else None
        if v is not None and str(v).strip() != "":
            return v
        if ws_formula is None:
            return v
        vf = ws_formula[ref].value
        if isinstance(vf, str) and vf.strip().startswith("="):
            return resolve_simple_formula(wb_f, ws_formula, vf)
        return vf

    fname_no_ext = file_path.stem

    # Headers / filename
    inv_a2 = header_value(ws_inv, ws_inv_f, "A2")
    inv_c4 = header_value(ws_inv, ws_inv_f, "C4")
    if inv_a2 != inv_c4:
        errors.append(f"INVOICE A2 != C4 ({inv_a2} / {inv_c4})")

    if ws_pack:
        pack_a2 = header_value(ws_pack, ws_pack_f, "A2")
        if not (inv_a2 == inv_c4 == pack_a2):
            errors.append(f"Entêtes incohérents A2/C4/PACK A2 ({inv_a2} / {inv_c4} / {pack_a2})")

    inv_c5 = str(header_value(ws_inv, ws_inv_f, "C5") or "").strip()
    inv_j4 = str(header_value(ws_inv, ws_inv_f, "J4") or "").strip()
    pack_b4 = str(header_value(ws_pack, ws_pack_f, "B4") or "").strip() if ws_pack else ""
    if not (fname_no_ext == inv_c5 == inv_j4 == pack_b4):
        errors.append(f"Nom fichier ≠ C5/J4/PACK B4 ({fname_no_ext}, {inv_c5}, {inv_j4}, {pack_b4})")

    def cell(ref):
        return ws_inv[ref].value

    j13 = str(cell("J13") or "").strip().upper()
    if j13 != "EUR":
        errors.append(f"J13 ≠ EUR ({cell('J13')})")

    j14 = str(cell("J14") or "").strip().upper()
    if j14 != "CIF":
        warnings.append(f"J14 ≠ CIF ({cell('J14')})")

    j16 = to_decimal(cell("J16"))
    if j16 != Decimal(4200):
        errors.append(f"J16 ≠ 4200 ({cell('J16')})")

    for r in range(11, 16):
        v = cell(f"C{r}")
        if v is None or str(v).strip() == "":
            errors.append(f"C{r} vide")

    for r in (16, 17):
        v = cell(f"C{r}")
        s = str(v or "").strip().upper()
        if not COUNTRY_CODE_RE.match(s):
            errors.append(f"C{r} doit être code pays 2 lettres ({v})")

    for r in (11, 13, 15):
        v = cell(f"J{r}")
        if v is None or str(v).strip() == "":
            errors.append(f"J{r} vide")

    c9 = str(cell("C9") or "").strip().upper()
    if c9 != "CN":
        errors.append(f"C9 doit être CN ({cell('C9')})")

    inv_sum = find_sum_row(ws_inv, 19, "B")
    pack_sum = find_sum_row(ws_pack, 6, "B") if ws_pack else None

    if inv_sum is None:
        errors.append("INVOICE ligne SUM introuvable")
    if ws_pack and pack_sum is None:
        errors.append("PACKING LIST ligne SUM introuvable")

    # Chinese check
    if inv_sum:
        for r in range(20, inv_sum):
            if contains_chinese(ws_inv[f"B{r}"].value):
                errors.append(f"Caractère chinois INVOICE B{r}")
                break
    if ws_pack and pack_sum:
        for r in range(6, pack_sum):
            if contains_chinese(ws_pack[f"B{r}"].value):
                errors.append(f"Caractère chinois PACKING B{r}")
                break

    # Merged forbidden in weight columns
    if inv_sum:
        merged = []
        for r in range(20, inv_sum):
            if is_cell_in_merged(ws_inv, r, 10): merged.append(f"J{r}")
            if is_cell_in_merged(ws_inv, r, 11): merged.append(f"K{r}")
        if merged:
            errors.append("INVOICE cellules fusionnées interdites J/K: " + ", ".join(merged[:20]))

    if ws_pack and pack_sum:
        merged = []
        for r in range(6, pack_sum):
            if is_cell_in_merged(ws_pack, r, 9): merged.append(f"I{r}")
            if is_cell_in_merged(ws_pack, r, 10): merged.append(f"J{r}")
        if merged:
            errors.append("PACKING cellules fusionnées interdites I/J: " + ", ".join(merged[:20]))

    # G length / D and I controls
    if inv_sum:
        if any(isinstance(ws_inv[f"G{r}"].value, str) and len(ws_inv[f"G{r}"].value.strip()) > 48 for r in range(20, inv_sum)):
            errors.append("INVOICE G contient une valeur > 48 caractères")

        bad_di = []
        for r in range(20, inv_sum):
            for col_letter, col_index in (("D", 4), ("I", 9)):
                if is_cell_in_merged(ws_inv, r, col_index):
                    continue
                val = ws_inv[f"{col_letter}{r}"].value
                dec = to_decimal(val)
                if val is None or str(val).strip() == "" or dec is None or dec == 0:
                    bad_di.append(f"{col_letter}{r}")
        if bad_di:
            errors.append("INVOICE D/I vide, 0 ou non numérique: " + ", ".join(bad_di[:30]))

        bad_weight = []
        text_weight = []
        for r in range(20, inv_sum):
            jv = ws_inv[f"J{r}"].value
            kv = ws_inv[f"K{r}"].value
            if isinstance(jv, str) or isinstance(kv, str):
                text_weight.append(r)
                continue
            jd = to_decimal(jv)
            kd = to_decimal(kv)
            if jd is not None and kd is not None and jd >= kd:
                bad_weight.append(r)
        if text_weight:
            errors.append("INVOICE J/K texte lignes: " + ", ".join(map(str, text_weight[:30])))
        if bad_weight:
            errors.append("INVOICE poids net >= brut lignes: " + ", ".join(map(str, bad_weight[:30])))

    if inv_sum and ws_pack and pack_sum:
        inv_pieces = to_decimal(ws_inv[f"H{inv_sum}"].value)
        inv_net = to_decimal(ws_inv[f"J{inv_sum}"].value)
        inv_gross = to_decimal(ws_inv[f"K{inv_sum}"].value)
        pack_pieces = to_decimal(ws_pack[f"H{pack_sum}"].value)
        pack_net = to_decimal(ws_pack[f"I{pack_sum}"].value)
        pack_gross = to_decimal(ws_pack[f"J{pack_sum}"].value)
        pack_cartons = to_decimal(get_merged_value(ws_pack, f"G{pack_sum}"))

        if inv_pieces != pack_pieces: errors.append(f"Total pièces diff INV/PACK ({inv_pieces}/{pack_pieces})")
        if inv_net != pack_net: errors.append(f"Total net diff INV/PACK ({inv_net}/{pack_net})")
        if inv_gross != pack_gross: errors.append(f"Total gross diff INV/PACK ({inv_gross}/{pack_gross})")
        if inv_net is not None and inv_gross is not None and inv_net > inv_gross:
            errors.append(f"Total net > gross ({inv_net}>{inv_gross})")
        if pack_cartons is None:
            errors.append(f"Nombre cartons manquant PACK G{pack_sum}")
        else:
            cartons = pack_cartons
        if pack_gross is not None:
            gross = pack_gross

        inv_b = [str(ws_inv[f"B{r}"].value or "").strip() for r in range(20, inv_sum)]
        pack_b = [str(ws_pack[f"B{r}"].value or "").strip() for r in range(6, pack_sum)]
        if inv_b != pack_b:
            errors.append(f"Descriptions B différentes INV/PACK ({len(inv_b)} lignes / {len(pack_b)} lignes)")

        line_errors = []
        for inv_row in range(20, inv_sum):
            pack_row = inv_row - 14
            if pack_row < 6 or pack_row >= pack_sum:
                continue
            inv_p = to_decimal(ws_inv[f"H{inv_row}"].value)
            inv_n = to_decimal(ws_inv[f"J{inv_row}"].value)
            inv_g = to_decimal(ws_inv[f"K{inv_row}"].value)
            pack_p = to_decimal(ws_pack[f"H{pack_row}"].value)
            pack_n = to_decimal(ws_pack[f"I{pack_row}"].value)
            pack_g = to_decimal(ws_pack[f"J{pack_row}"].value)
            q = Decimal("0.01")
            if inv_p != pack_p or (inv_n and pack_n and inv_n.quantize(q) != pack_n.quantize(q)) or (inv_g and pack_g and inv_g.quantize(q) != pack_g.quantize(q)):
                line_errors.append(inv_row)
        if line_errors:
            errors.append("Différences ligne-par-ligne INV/PACK: " + ", ".join(map(str, line_errors[:30])))

    if ws_pack and pack_sum:
        bad_g = []
        for r in range(6, pack_sum):
            raw = get_effective_cell_value(ws_pack, r, 7)
            dec = to_decimal(raw)
            if raw is None or str(raw).strip() == "" or dec is None or dec == 0:
                bad_g.append(f"G{r}")
        if bad_g:
            errors.append("PACKING G vide/0/non numérique: " + ", ".join(bad_g[:30]))

    return {
        "file": file_path.name,
        "errors": errors,
        "warnings": warnings,
        "cartons": float(cartons),
        "gross": float(gross),
    }

# =========================
# SUMMARY / SGS DATA LOGIC
# =========================
def find_header_row_by_keyword(ws, col, keyword):
    for row in range(1, ws.max_row + 1):
        v = ws.cell(row=row, column=col).value
        if isinstance(v, str) and keyword in v:
            return row
    return None

def _merged_top_left(ws, row, col):
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return rng.min_row, rng.min_col, str(rng)
    return row, col, None

def _get_cell_value_merged(ws, row, col):
    r0, c0, rng = _merged_top_left(ws, row, col)
    return ws.cell(row=r0, column=c0).value, rng

def parse_invoice(ws):
    sub_order_no = str(ws["C5"].value or "").strip()
    header_row = find_header_row_by_keyword(ws, 2, "Description of Goods")
    if header_row is None:
        return [], sub_order_no
    products = []
    for r in range(header_row + 1, ws.max_row + 1):
        desc = ws.cell(r, 2).value
        if desc is None:
            break
        if isinstance(desc, str) and desc.strip().upper() == "SUM":
            break
        hs = ws.cell(r, 3).value
        val = ws.cell(r, 9).value
        mark = ws.cell(r, 7).value
        if hs is None or desc is None:
            continue
        products.append({
            "hs_code": str(hs).strip(),
            "desc": str(desc).strip(),
            "custom_value": to_float(val),
            "mark": str(mark).strip() if mark not in (None, "") else "",
            "sub_order_no": sub_order_no,
        })
    return products, sub_order_no

def build_invoice_index(products):
    by_mark_desc_seq = defaultdict(list)
    by_desc = defaultdict(list)
    for p in products:
        if p.get("mark"):
            by_mark_desc_seq[(p["mark"], p["desc"])].append(p["hs_code"])
        by_desc[p["desc"]].append(p["hs_code"])
    return by_mark_desc_seq, by_desc

def parse_packing_list_rows(ws):
    header_row = find_header_row_by_keyword(ws, 2, "Description of Goods")
    if header_row is None:
        return []
    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        desc = ws.cell(r, 2).value
        if desc is None:
            break
        if isinstance(desc, str) and desc.strip().upper() == "SUM":
            break
        mark_val, _ = _get_cell_value_merged(ws, r, 5)
        carton_val, carton_rng = _get_cell_value_merged(ws, r, 7)
        net_val, net_rng = _get_cell_value_merged(ws, r, 9)
        gross_val, gross_rng = _get_cell_value_merged(ws, r, 10)
        rows.append({
            "desc": str(desc).strip(),
            "mark": str(mark_val).strip() if mark_val not in (None, "") else "",
            "carton": to_float(carton_val),
            "net": to_float(net_val),
            "gross": to_float(gross_val),
            "carton_rng": carton_rng or f"R{r}C7",
            "net_rng": net_rng or f"R{r}C9",
            "gross_rng": gross_rng or f"R{r}C10",
        })
    return rows

def collect_from_file(filepath):
    def norm_desc(s):
        s = str(s or "").strip().lower()
        s = " ".join(s.split())
        s = "".join(ch if ch.isalnum() or ch.isspace() else " " for ch in s)
        return " ".join(s.split())

    wb = load_workbook(filepath, data_only=True)
    ws_inv = sheet_by_name_ci(wb, "INVOICE")
    ws_pl = sheet_by_name_ci(wb, "PACKING LIST")
    if ws_inv is None or ws_pl is None:
        return [], [], []

    inv_products, sub_order_no = parse_invoice(ws_inv)
    by_mark_desc_seq, by_desc = build_invoice_index(inv_products)
    pl_rows = parse_packing_list_rows(ws_pl)

    groups, idx_map = [], {}
    for pr in pl_rows:
        crng = pr["carton_rng"]
        if crng not in idx_map:
            idx_map[crng] = len(groups)
            groups.append({"carton_rng": crng, "rows": []})
        groups[idx_map[crng]]["rows"].append(pr)

    group_calc = []
    pl_md_counter = Counter()
    for g in groups:
        rows = g["rows"]
        total_cartons = float(rows[0]["carton"] or 0.0) if rows else 0.0
        hs_list = []
        for pr in rows:
            if pr["mark"] and (pr["mark"], pr["desc"]) in by_mark_desc_seq:
                seq = by_mark_desc_seq[(pr["mark"], pr["desc"])]
                idx = pl_md_counter[(pr["mark"], pr["desc"])]
                if idx >= len(seq):
                    idx = len(seq) - 1
                hs = seq[idx] if seq else "PL_UNMATCHED"
                pl_md_counter[(pr["mark"], pr["desc"])] += 1
            else:
                cands = by_desc.get(pr["desc"], [])
                if cands:
                    freq = Counter(cands)
                    hs = sorted(freq.items(), key=lambda x: (-x[1], x[0]))[0][0]
                else:
                    hs = "PL_UNMATCHED"
            hs_list.append(hs)
        group_calc.append({"rows": rows, "hs_list": hs_list, "total_cartons": total_cartons})

    def row_key(hs, desc):
        return (hs, norm_desc(desc))

    occ_counter = Counter()
    for gc in group_calc:
        seen = set(row_key(gc["hs_list"][i], pr["desc"]) for i, pr in enumerate(gc["rows"]))
        for k in seen:
            occ_counter[k] += 1

    assigned_carton_so_far = Counter()
    for gc in group_calc:
        rows, hs_list = gc["rows"], gc["hs_list"]
        n = len(rows)
        total_cartons = int(gc["total_cartons"] or 0)
        cartons_assigned = [0.0] * n
        if total_cartons >= n and n:
            for i in range(n): cartons_assigned[i] = 1.0
            extra = total_cartons - n
            if extra > 0: cartons_assigned[0] += float(extra)
        elif total_cartons > 0:
            weights = []
            for i, pr in enumerate(rows):
                k = row_key(hs_list[i], pr["desc"])
                weights.append((assigned_carton_so_far[k], occ_counter[k], -float(pr.get("net") or 0.0), i))
            weights.sort()
            for _, _, _, idx in weights[:total_cartons]:
                cartons_assigned[idx] = 1.0
        recipient_idx = next((i for i, v in enumerate(cartons_assigned) if v > 0), None)
        for i, v in enumerate(cartons_assigned):
            if v > 0:
                assigned_carton_so_far[row_key(hs_list[i], rows[i]["desc"])] += v
        gc["cartons_assigned"] = cartons_assigned
        gc["recipient_idx"] = recipient_idx

    used_gross_rng, used_net_rng = set(), set()
    pl_lines, transfers = [], []
    for gc in group_calc:
        rows, hs_list = gc["rows"], gc["hs_list"]
        recipient_idx = gc.get("recipient_idx")
        r_key = (hs_list[recipient_idx], rows[recipient_idx]["desc"]) if recipient_idx is not None else None
        for i, pr in enumerate(rows):
            net = float(pr.get("net") or 0.0)
            gross = float(pr.get("gross") or 0.0)
            if pr.get("net_rng") in used_net_rng: net = 0.0
            else: used_net_rng.add(pr.get("net_rng"))
            if pr.get("gross_rng") in used_gross_rng: gross = 0.0
            else: used_gross_rng.add(pr.get("gross_rng"))
            carton = float(gc["cartons_assigned"][i] or 0.0)
            pl_lines.append((hs_list[i], pr["desc"], carton, net, gross, sub_order_no))
            if carton == 0.0 and r_key is not None and i != recipient_idx:
                transfers.append(((hs_list[i], pr["desc"]), r_key, gross))

    value_lines = [(p["hs_code"], p["desc"], p["custom_value"], sub_order_no) for p in inv_products]
    return value_lines, pl_lines, transfers

def aggregate(value_lines, pl_lines, transfers=None):
    agg = defaultdict(lambda: {"carton": 0.0, "net": 0.0, "gross": 0.0, "value": 0.0, "invoices": set()})
    for hs, desc, val, inv in value_lines:
        agg[(hs, desc)]["value"] += float(val or 0.0)
        if inv: agg[(hs, desc)]["invoices"].add(inv)
    for hs, desc, carton, net, gross, inv in pl_lines:
        agg[(hs, desc)]["carton"] += float(carton or 0.0)
        agg[(hs, desc)]["net"] += float(net or 0.0)
        agg[(hs, desc)]["gross"] += float(gross or 0.0)
        if inv: agg[(hs, desc)]["invoices"].add(inv)
    if transfers:
        for (from_hs, from_desc), (to_hs, to_desc), amt in transfers:
            amt = float(amt or 0.0)
            from_key, to_key = (from_hs, from_desc), (to_hs, to_desc)
            if from_key in agg and float(agg[from_key]["carton"] or 0.0) == 0.0:
                agg[from_key]["gross"] -= amt
                agg[to_key]["gross"] += amt
        for key, v in agg.items():
            if float(v["carton"] or 0.0) == 0.0 or v["gross"] < 0:
                v["gross"] = 0.0
    return agg

def parse_sub_orders(inv_str: str) -> str:
    nums = []
    for p in str(inv_str or "").split(","):
        m = re.search(r"(\d+)\s*$", p.strip())
        if m: nums.append(str(int(m.group(1))))
    seen, out = set(), []
    for n in nums:
        if n not in seen:
            seen.add(n); out.append(n)
    return " ".join(out)

def copy_row_style(ws, source_row, target_row, max_col=83):
    for c in range(1, max_col + 1):
        src = ws.cell(source_row, c)
        dst = ws.cell(target_row, c)
        if src.has_style:
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)
            dst.alignment = copy(src.alignment)
            dst.number_format = src.number_format
            dst.protection = copy(src.protection)
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height

def create_sgs_workbook(agg_data, container_no, bl_number, terminal_choice, template_path):
    wb = load_workbook(template_path)
    if "Data" not in wb.sheetnames:
        raise ValueError("Template T1_SGS.xlsx: feuille 'Data' introuvable")
    ws = wb["Data"]

    items = []
    for (hs, desc) in sorted(agg_data.keys(), key=lambda k: (k[0], k[1])):
        d = agg_data[(hs, desc)]
        if not str(hs or "").strip() or not str(desc or "").strip():
            continue
        items.append({
            "hs": str(hs).strip(),
            "desc": str(desc).strip(),
            "carton": round(float(d.get("carton", 0) or 0), 2),
            "net": round(float(d.get("net", 0) or 0), 2),
            "gross": round(float(d.get("gross", 0) or 0), 2),
            "value": round(float(d.get("value", 0) or 0), 2),
            "sub_orders": parse_sub_orders(", ".join(sorted(d.get("invoices", set()) or []))),
        })

    if not items:
        raise ValueError("Aucune ligne SGS à générer")

    # Clear old rows from row 3 down, columns A:CE (83 cols)
    clear_to = max(ws.max_row, 3 + len(items) + 50)
    for r in range(3, clear_to + 1):
        for c in range(1, 84):
            ws.cell(r, c).value = None

    # Terminal columns O/P/Q
    sender_address = ""
    sender_city = ""
    sender_postcode = ""
    if terminal_choice == "Delta":
        sender_address = "EUROPEAWEG 875"
        sender_city = "Rotterdam"
        sender_postcode = "3199 LD"
    elif terminal_choice == "Euromax":
        sender_address = "Maasvlakteweg 951"
        sender_city = "Rotterdam"
        sender_postcode = "3199 LZ"
    # Empty => leave O/P/Q empty

    year = datetime.now().year
    const = {
        "H": "EUR-Euro",
        "K": "SGSac94206311f08d",
        "M": "NL-Dutch",
        "O": sender_address,
        "P": sender_city,
        "Q": sender_postcode,
        "R": "BE0417688928",
        "T": "BE-Belgian",
        "U": "Hawe",
        "V": "Kruiningenstraat 188",
        "W": "Schoten",
        "X": 2900,
        "AC": "NM",
        "AD": "CT-Carton",
        "AF": "N705-Bill of lading",
        "AG": bl_number,
        "AN": "N380-Commercial Invoice",
        "AR": year,
        "AS": "N730-Road consignment note",
        "AT": container_no,
        "BL": "CN-China",
        "BM": "BE-Belgian",
        "BN": container_no,
    }
    col_idx = {k: column_index_from_string(k) for k in const}

    source_style_row = 3 if ws.max_row >= 3 else 2
    for i, it in enumerate(items):
        r = 3 + i
        copy_row_style(ws, source_style_row, r, 83)
        ws.cell(r, 1).value = it["hs"]          # A
        ws.cell(r, 2).value = it["desc"]        # B
        ws.cell(r, 7).value = it["value"]       # G
        ws.cell(r, 9).value = it["net"]         # I
        ws.cell(r, 10).value = it["gross"]      # J
        ws.cell(r, 31).value = it["carton"]     # AE
        ws.cell(r, 41).value = it["sub_orders"] # AO
        for letter, val in const.items():
            ws.cell(r, col_idx[letter]).value = val

    # Delete empty rows after last generated line
    last_data_row = 3 + len(items) - 1
    if ws.max_row > last_data_row:
        ws.delete_rows(last_data_row + 1, ws.max_row - last_data_row)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue(), len(items)

# =========================
# STREAMLIT UI
# =========================
st.set_page_config(page_title="SGS Generator", layout="wide")
st.title("SGS Generator")

uploaded_files = st.file_uploader(
    "Upload invoice Excel files",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
)

template_upload = st.file_uploader(
    "Template T1_SGS.xlsx (optionnel si le fichier est dans le même dossier que l'app)",
    type=["xlsx"],
    accept_multiple_files=False,
)

if uploaded_files:
    files_sorted = sorted(uploaded_files, key=lambda f: natural_key(f.name))

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        saved_paths = []
        for uf in files_sorted:
            p = tmpdir / uf.name
            p.write_bytes(uf.getbuffer())
            saved_paths.append(p)

        # Save template
        if template_upload is not None:
            template_path = tmpdir / TEMPLATE_NAME
            template_path.write_bytes(template_upload.getbuffer())
        else:
            template_path = APP_DIR / TEMPLATE_NAME

        # Final check
        results = [final_check_file(p) for p in saved_paths]
        df = pd.DataFrame([{
            "File": r["file"],
            "Status": "ERROR" if r["errors"] else ("WARNING" if r["warnings"] else "OK"),
            "Errors": len(r["errors"]),
            "Warnings": len(r["warnings"]),
            "Cartons": r["cartons"],
            "Gross Weight": r["gross"],
        } for r in results])

        files_checked = len(results)
        files_errors = sum(1 for r in results if r["errors"])
        files_warnings_only = sum(1 for r in results if (not r["errors"] and r["warnings"]))
        total_cartons = sum(float(r["cartons"] or 0) for r in results)
        total_gross = sum(float(r["gross"] or 0) for r in results)

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Files checked", files_checked)
        c2.metric("Files with errors", files_errors)
        c3.metric("Files with warnings only", files_warnings_only)
        c4.metric("Total cartons", f"{total_cartons:,.0f}")
        c5.metric("Total gross weight", f"{total_gross:,.2f}")

        if files_errors == 0:
            st.success(f"{files_checked} file(s) passed with no blocking issue(s).")
        else:
            st.error(f"{files_errors} file(s) have errors. SGS generation is blocked.")

        st.subheader("Summary")
        st.dataframe(df, use_container_width=True, hide_index=True)

        st.subheader("Files with issues")
        issue_rows = []
        for r in results:
            for e in r["errors"]:
                issue_rows.append({"File": r["file"], "Type": "ERROR", "Message": e})
            for w in r["warnings"]:
                issue_rows.append({"File": r["file"], "Type": "WARNING", "Message": w})
        if issue_rows:
            st.dataframe(pd.DataFrame(issue_rows), use_container_width=True, hide_index=True)
        else:
            st.info("No issues found.")

        st.divider()
        st.subheader("SGS Generation")

        terminal = st.radio("Terminal", ["Delta", "Euromax", "Empty"], horizontal=True)
        bl_number = st.text_input("Numéro BL")

        generate = st.button("Générer T1_SGS", type="primary", disabled=(files_errors > 0))

        if generate:
            if not template_path.exists():
                st.error("Template T1_SGS.xlsx introuvable. Mets T1_SGS.xlsx dans le même dossier que app_sgs_streamlit.py ou upload le template.")
            elif not bl_number.strip():
                st.error("Complète le numéro BL avant de générer.")
            else:
                try:
                    all_value_lines, all_pl_lines, all_transfers = [], [], []
                    for p in saved_paths:
                        v, pl, tr = collect_from_file(p)
                        all_value_lines.extend(v)
                        all_pl_lines.extend(pl)
                        all_transfers.extend(tr)
                    agg = aggregate(all_value_lines, all_pl_lines, all_transfers)
                    container_no = re.sub(r"-\d+$", "", saved_paths[0].stem).upper()
                    data, xlines = create_sgs_workbook(agg, container_no, bl_number.strip(), terminal, template_path)
                    out_name = f"T1_SGS-{container_no}-{datetime.now().strftime('%Y%m%d')}-{xlines}lines.xlsx"
                    st.success(f"Fichier généré: {out_name}")
                    st.download_button(
                        "Télécharger T1_SGS",
                        data=data,
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except Exception as e:
                    st.error(f"Erreur génération SGS: {e}")
else:
    st.info("Upload les fichiers INVOICE pour commencer.")

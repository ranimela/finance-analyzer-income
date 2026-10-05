"""Reconciliation working file engine matching income ledger with Agreements Payments."""

import csv
import datetime
from pathlib import Path
from typing import Any
from collections import defaultdict
import openpyxl


def generate_income_working_file(
    ledger_csv_path: Path,
    agreements_payments_path: Path,
    output_csv_path: Path,
) -> dict[str, Any]:
    """Generate income_working.csv by enriching ledger transactions with Agreements Payments.

    Per line in ledger_csv:
    - If Column F (אסמ') > 0:
        - Look up occurrences in Agreements Payments.
        - Add column 'occurences found' with match count.
        - Add columns 'Invoice Start Date' and 'Invoice End Date'.
        - If matched > 1 times and not all values of Start Date or End Date are identical,
          mark both date columns with 'ERROR - multiple dates'.
    - If Column F <= 0 or non-numeric:
        - Leave new columns blank.

    Args:
        ledger_csv_path: Path to the clean ledger CSV (e.g. step9 CSV).
        agreements_payments_path: Path to Agreements Payments.xlsx.
        output_csv_path: Path to write income_working.csv.

    Returns:
        dict[str, Any]: Audit metrics of lookups, matches, and error conditions.
    """
    if not ledger_csv_path.exists():
        raise FileNotFoundError(f"Clean ledger CSV not found: {ledger_csv_path}")
    if not agreements_payments_path.exists():
        raise FileNotFoundError(f"Agreements Payments XLSX not found: {agreements_payments_path}")

    output_csv_path.parent.mkdir(parents=True, exist_ok=True)

    # 1. Parse Agreements Payments.xlsx
    wb = openpyxl.load_workbook(agreements_payments_path, data_only=True, read_only=True)
    ws = wb.active
    if ws is None:
        raise ValueError("Agreements Payments workbook has no active sheet.")

    invoice_map: dict[str, list[tuple[str, str]]] = defaultdict(list)

    for r_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
        if r_idx <= 7:
            continue
        if len(row) < 6:
            continue

        raw_inv = row[3]  # Col D is index 3
        if raw_inv is None:
            continue

        inv_str = str(raw_inv).strip()
        if inv_str.endswith(".0") and inv_str[:-2].isdigit():
            inv_str = inv_str[:-2]

        if not inv_str or inv_str == "0":
            continue

        start_cell = row[4]  # Col E is index 4
        end_cell = row[5]    # Col F is index 5

        def _to_dmy(cell: Any) -> str:
            if isinstance(cell, datetime.datetime):
                return cell.strftime("%d/%m/%Y")
            s = str(cell or "").strip()
            if not s:
                return ""
            if len(s) == 10 and s[4] == "-" and s[7] == "-":
                parts = s.split("-")
                return f"{parts[2]}/{parts[1]}/{parts[0]}"
            return s

        s_date = _to_dmy(start_cell)
        e_date = _to_dmy(end_cell)

        invoice_map[inv_str].append((s_date, e_date))

    wb.close()

    # 2. Process clean ledger CSV
    with open(ledger_csv_path, mode="r", encoding="utf-8-sig", newline="") as f_in:
        reader = list(csv.reader(f_in))

    if not reader:
        raise ValueError(f"Ledger CSV is empty: {ledger_csv_path}")

    header = reader[0]
    data_rows = reader[1:]

    # Locate Column C (index 2), Column D (index 3), Column F (index 5), Column G (index 6), Column J (index 9)
    col_c_idx = 2
    if "קוד מיון" in header:
        col_c_idx = header.index("קוד מיון")

    col_d_idx = 3
    if "ת.אסמכ" in header:
        col_d_idx = header.index("ת.אסמכ")

    col_f_idx = 5
    if "אסמ'" in header:
        col_f_idx = header.index("אסמ'")

    col_g_idx = 6
    if "אסמ'2" in header:
        col_g_idx = header.index("אסמ'2")

    col_j_idx = 9
    if "חובה / זכות (שקל) חובה" in header and "חובה / זכות (שקל) זכות" in header:
        col_j_idx = header.index("חובה / זכות (שקל) זכות")
    elif "חובה / זכות (שקל) זכות" in header:
        col_j_idx = header.index("חובה / זכות (שקל) זכות")

    new_header = header + [
        "occurences found",
        "Invoice Start Date",
        "Invoice End Date",
        "invoice month duration",
        "adjusted monthly income",
    ]
    enriched_rows: list[list[str]] = [new_header]

    metrics: dict[str, Any] = {
        "total_ledger_rows": len(data_rows),
        "zero_or_non_positive_rows": 0,
        "ap_zero_matches": 0,
        "ap_single_matches": 0,
        "ap_multiple_matches_consistent": 0,
        "ap_multiple_matches_has_error": 0,
        "start_date_errors": 0,
        "end_date_errors": 0,
        "duration_calculated_rows": 0,
        "duration_skipped_rows": 0,
        "adjusted_income_calculated_rows": 0,
        "special_mion_direct_allocation_rows": 0,
    }

    intermediate_records: list[dict[str, Any]] = []

    for row in data_rows:
        raw_c = row[col_c_idx].strip() if len(row) > col_c_idx else ""
        raw_d = row[col_d_idx].strip() if len(row) > col_d_idx else ""
        raw_f = row[col_f_idx].strip() if len(row) > col_f_idx else ""
        raw_g = row[col_g_idx].strip() if len(row) > col_g_idx else ""
        raw_j = row[col_j_idx].strip() if len(row) > col_j_idx else ""

        # Determine lookup invoice key: primary Column F (אסמ'), fallback to Column G (אסמ'2) if Col F <= 0
        try:
            val_f = float(raw_f)
            has_pos_f = val_f > 0
        except ValueError:
            has_pos_f = False

        lookup_key = ""
        if has_pos_f:
            lookup_key = raw_f
        else:
            # Fallback to Column G (אסמ'2)
            try:
                val_g = float(raw_g)
                has_pos_g = val_g > 0
            except ValueError:
                has_pos_g = False

            if has_pos_g:
                lookup_key = raw_g

        occ_str = ""
        s_date_str = ""
        e_date_str = ""

        if not lookup_key:
            metrics["zero_or_non_positive_rows"] += 1
        else:
            inv_key = lookup_key
            if inv_key.endswith(".0") and inv_key[:-2].isdigit():
                inv_key = inv_key[:-2]

            matches = invoice_map.get(inv_key, [])
            occ = len(matches)

            if occ == 0:
                metrics["ap_zero_matches"] += 1
                occ_str = "0"
            elif occ == 1:
                metrics["ap_single_matches"] += 1
                occ_str = "1"
                s_date_str, e_date_str = matches[0]
            else:
                occ_str = str(occ)
                unique_starts = set(m[0] for m in matches)
                unique_ends = set(m[1] for m in matches)

                start_err = len(unique_starts) > 1
                end_err = len(unique_ends) > 1

                if start_err or end_err:
                    metrics["ap_multiple_matches_has_error"] += 1
                    if start_err:
                        metrics["start_date_errors"] += 1
                    if end_err:
                        metrics["end_date_errors"] += 1

                    s_date_str = "ERROR - multiple dates" if start_err else matches[0][0]
                    e_date_str = "ERROR - multiple dates" if end_err else matches[0][1]
                else:
                    metrics["ap_multiple_matches_consistent"] += 1
                    s_date_str, e_date_str = matches[0]

        # Calculate duration and adjusted monthly income
        # If Column C is 94501 or 93002: do NOT split income to quarters, skip duration calculation
        duration_str = ""
        adj_income_str = ""
        is_special_mion = raw_c in ("94501", "93002")

        if is_special_mion:
            metrics["special_mion_direct_allocation_rows"] += 1
        elif (
            not s_date_str
            or not e_date_str
            or "ERROR" in s_date_str
            or "ERROR" in e_date_str
        ):
            metrics["duration_skipped_rows"] += 1
        else:
            try:
                ds = datetime.datetime.strptime(s_date_str, "%d/%m/%Y")
                de = datetime.datetime.strptime(e_date_str, "%d/%m/%Y")
                # Calendar month span inclusive: (end.year - start.year)*12 + (end.month - start.month) + 1
                duration = (de.year - ds.year) * 12 + (de.month - ds.month) + 1
                duration_str = str(duration)
                metrics["duration_calculated_rows"] += 1

                # Income from Col J (Credit)
                cleaned_income = raw_j.replace(",", "").strip()
                if cleaned_income:
                    income_val = float(cleaned_income)
                    adj_val = income_val / duration
                    adj_income_str = f"{adj_val:.2f}"
                    metrics["adjusted_income_calculated_rows"] += 1
            except Exception:
                duration_str = ""
                adj_income_str = ""

        # Intermediate collection before quarter allocation
        intermediate_records.append({
            "base_row": row,
            "raw_c": raw_c,
            "raw_d": raw_d,
            "raw_j": raw_j,
            "occ_str": occ_str,
            "s_date_str": s_date_str,
            "e_date_str": e_date_str,
            "duration_str": duration_str,
            "adj_income_str": adj_income_str,
        })

    # 3. Determine latest quarter needed
    # Default minimum end year/quarter: 2024 Q1
    max_year = 2024
    max_quarter = 1

    for rec in intermediate_records:
        e_s = rec["e_date_str"]
        if e_s and "ERROR" not in e_s:
            try:
                de = datetime.datetime.strptime(e_s, "%d/%m/%Y")
                q_num = (de.month - 1) // 3 + 1
                if (de.year > max_year) or (de.year == max_year and q_num > max_quarter):
                    max_year = de.year
                    max_quarter = q_num
            except Exception:
                pass

    # Build sequence of quarter tuples: (year, q_num, label)
    quarter_cols: list[tuple[int, int, str]] = []
    curr_y = 2024
    curr_q = 1
    while (curr_y < max_year) or (curr_y == max_year and curr_q <= max_quarter):
        yy_str = str(curr_y)[-2:]
        lbl = f"{yy_str}Q{curr_q}"
        quarter_cols.append((curr_y, curr_q, lbl))
        curr_q += 1
        if curr_q > 4:
            curr_q = 1
            curr_y += 1

    quarter_headers = [col[2] for col in quarter_cols]
    final_header = new_header + quarter_headers
    final_rows: list[list[str]] = [final_header]

    metrics["quarter_columns_count"] = len(quarter_headers)
    metrics["max_allocated_quarter"] = quarter_headers[-1] if quarter_headers else "24Q1"
    metrics["rows_with_quarter_allocations"] = 0

    for rec in intermediate_records:
        row_ext = [
            rec["occ_str"],
            rec["s_date_str"],
            rec["e_date_str"],
            rec["duration_str"],
            rec["adj_income_str"],
        ]

        q_allocations = [""] * len(quarter_cols)
        raw_c = rec["raw_c"]
        raw_d = rec["raw_d"]
        raw_j = rec["raw_j"]
        adj_s = rec["adj_income_str"]
        s_s = rec["s_date_str"]
        e_s = rec["e_date_str"]

        # Step A: If Column C is 94501 or 93002, do NOT split income to quarters.
        # Instead, add the entire income amount to the appropriate quarter based on Column D (ת.אסמכ).
        if raw_c in ("94501", "93002"):
            cleaned_income = raw_j.replace(",", "").strip()
            if cleaned_income and raw_d:
                try:
                    inc_val = float(cleaned_income)
                    dt_d = datetime.datetime.strptime(
                        raw_d, "%Y-%m-%d" if "-" in raw_d else "%d/%m/%Y"
                    ).date()
                    target_q_num = (dt_d.month - 1) // 3 + 1
                    target_lbl = f"{str(dt_d.year)[-2:]}Q{target_q_num}"

                    for q_idx, (_, _, q_lbl) in enumerate(quarter_cols):
                        if q_lbl == target_lbl:
                            q_allocations[q_idx] = f"{inc_val:.2f}"
                            metrics["rows_with_quarter_allocations"] += 1
                            break
                except Exception:
                    pass

        # Step B: Normal multi-quarter split based on start and end dates
        elif adj_s and s_s and e_s and "ERROR" not in s_s and "ERROR" not in e_s:
            try:
                adj_float = float(adj_s)
                ds = datetime.datetime.strptime(s_s, "%d/%m/%Y").date()
                de = datetime.datetime.strptime(e_s, "%d/%m/%Y").date()

                had_alloc = False
                for q_idx, (qy, qnum, _) in enumerate(quarter_cols):
                    qm_start = (qnum - 1) * 3 + 1
                    qm_end = qnum * 3

                    # Count months of the invoice that fall into this quarter
                    months_overlap = 0
                    cy, cm = ds.year, ds.month
                    while (cy < de.year) or (cy == de.year and cm <= de.month):
                        if cy == qy and qm_start <= cm <= qm_end:
                            months_overlap += 1
                        cm += 1
                        if cm > 12:
                            cm = 1
                            cy += 1

                    if months_overlap > 0:
                        allocated_val = months_overlap * adj_float
                        q_allocations[q_idx] = f"{allocated_val:.2f}"
                        had_alloc = True

                if had_alloc:
                    metrics["rows_with_quarter_allocations"] += 1
            except Exception:
                pass

        final_rows.append(rec["base_row"] + row_ext + q_allocations)

    with open(output_csv_path, mode="w", encoding="utf-8-sig", newline="") as f_out:
        writer = csv.writer(f_out)
        writer.writerows(final_rows)

    metrics["total_output_rows"] = len(final_rows) - 1
    return metrics

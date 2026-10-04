"""Module for cleaning Israeli General Ledger Excel exports (Income accounts).

Implements the 6-step cleaning pipeline adapted for income ledgers:
- Step 1: Skips rows 1-3, extracts row 4 as headers with synthetic names for cols A-C,
          and filters opening balance ('יתרת פתיחה') and closing balance ('יתרת סגירה').
- Step 2: Locates 'סה"כ מפתח חשבון' rows and deletes them along with subsequent follower rows.
- Step 3: Forward-fills empty cells in columns A, B, and C.
- Step 4: Deletes rows where transaction columns D to Q are all empty.
- Step 5: Deletes rows where Column A contains 'הפרשה'.
- Step 6: Deletes rows where Column A contains 'סהכ לדוח' / 'סה"כ לדוח'.
Outputs the final clean dataset directly to step6 CSV encoded in utf-8-sig.
"""

import csv
import datetime
from pathlib import Path
from typing import Any
import openpyxl


def clean_income_ledger(
    input_path: Path,
    output_path: Path,
) -> dict[str, Any]:
    """Execute the full 6-step cleaning pipeline on an income general ledger.

    Args:
        input_path: Path to the source .xlsx file.
        output_path: Target path for the final step6 .csv file.

    Returns:
        dict[str, Any]: Audit metrics tracking row counts and filtering impact at each step.
    """
    if not input_path.exists():
        raise FileNotFoundError(f"Input file does not exist: {input_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    metrics: dict[str, Any] = {
        "step0_raw_rows_read": 0,
        "step1_metadata_skipped": 0,
        "step1_opening_balance_removed": 0,
        "step1_closing_balance_removed": 0,
        "step1_retained_rows": 0,
        "step2_subtotal_anchor_matches": 0,
        "step2_follower_rows_deleted": 0,
        "step2_retained_rows": 0,
        "step3_col_a_cells_filled": 0,
        "step3_col_b_cells_filled": 0,
        "step3_col_c_cells_filled": 0,
        "step4_empty_tx_rows_deleted": 0,
        "step4_retained_rows": 0,
        "step5_hafrasha_rows_deleted": 0,
        "step5_retained_rows": 0,
        "step6_report_totals_deleted": 0,
        "step6_final_emitted_data_rows": 0,
    }

    # Open workbook in read-only and data-only mode to prevent memory exhaustion
    wb = openpyxl.load_workbook(input_path, data_only=True, read_only=True)
    ws = wb.active
    if ws is None:
        raise ValueError("Workbook contains no active worksheet.")

    rows_iter = ws.iter_rows(values_only=True)

    # -------------------------------------------------------------
    # Step 1: Skip metadata rows 1-3, extract headers, filter balances
    # -------------------------------------------------------------
    for _ in range(3):
        try:
            next(rows_iter)
            metrics["step1_metadata_skipped"] += 1
            metrics["step0_raw_rows_read"] += 1
        except StopIteration:
            break

    try:
        raw_header = list(next(rows_iter))
        metrics["step0_raw_rows_read"] += 1
    except StopIteration:
        raise ValueError("File contains fewer than 4 rows; header missing.")

    headers: list[str] = []
    col_count = max(len(raw_header), 17)
    for idx in range(col_count):
        val = raw_header[idx] if idx < len(raw_header) else None
        match idx:
            case 0:
                headers.append("שם חשבון")
            case 1:
                headers.append("מפתח חשבון")
            case 2:
                headers.append("קוד מיון")
            case _:
                headers.append(str(val).strip() if val is not None else f"Col_{idx + 1}")

    step1_rows: list[list[str]] = []
    for row in rows_iter:
        metrics["step0_raw_rows_read"] += 1
        row_list = list(row)

        if len(row_list) < len(headers):
            row_list.extend([None] * (len(headers) - len(row_list)))
        elif len(row_list) > len(headers):
            row_list = row_list[: len(headers)]

        # Column C is index 2: filter 'יתרת פתיחה'
        col_c = str(row_list[2]).strip() if row_list[2] is not None else ""
        if col_c == "יתרת פתיחה":
            metrics["step1_opening_balance_removed"] += 1
            continue

        # Column M is index 12: filter 'יתרת סגירה'
        col_m = str(row_list[12]).strip() if len(row_list) > 12 and row_list[12] is not None else ""
        if col_m == "יתרת סגירה":
            metrics["step1_closing_balance_removed"] += 1
            continue

        formatted_row: list[str] = []
        for cell in row_list:
            if isinstance(cell, datetime.datetime):
                formatted_row.append(cell.strftime("%Y-%m-%d"))
            elif cell is None:
                formatted_row.append("")
            else:
                formatted_row.append(str(cell).strip())

        step1_rows.append(formatted_row)

    wb.close()
    metrics["step1_retained_rows"] = len(step1_rows)

    # -------------------------------------------------------------
    # Step 2: Remove subtotal lines and trailing summary follower rows
    # -------------------------------------------------------------
    indices_to_delete: set[int] = set()
    num_step1 = len(step1_rows)
    idx = 0
    while idx < num_step1:
        row = step1_rows[idx]
        col_a = row[0].strip() if row else ""

        if col_a in ('סה"כ מפתח חשבון', "סה״כ מפתח חשבון"):
            follower_indices: list[int] = []

            for offset in range(1, 4):
                next_idx = idx + offset
                if next_idx < num_step1:
                    next_row = step1_rows[next_idx]
                    next_col_a = next_row[0].strip() if next_row else ""
                    if next_col_a == "":
                        follower_indices.append(next_idx)
                    else:
                        break
                else:
                    break

            metrics["step2_subtotal_anchor_matches"] += 1
            indices_to_delete.add(idx)
            for f_idx in follower_indices:
                indices_to_delete.add(f_idx)
                metrics["step2_follower_rows_deleted"] += 1
            idx += 1 + len(follower_indices)
            continue

        idx += 1

    step2_rows = [row for i, row in enumerate(step1_rows) if i not in indices_to_delete]
    metrics["step2_retained_rows"] = len(step2_rows)

    # -------------------------------------------------------------
    # Step 3: Forward-fill empty cells in columns A, B, and C
    # -------------------------------------------------------------
    last_val = ["", "", ""]
    step3_rows: list[list[str]] = []
    for row in step2_rows:
        new_row = list(row)
        while len(new_row) < 3:
            new_row.append("")

        for col_idx in range(3):
            cell_val = new_row[col_idx].strip()
            if cell_val != "":
                last_val[col_idx] = cell_val
            else:
                new_row[col_idx] = last_val[col_idx]
                match col_idx:
                    case 0:
                        metrics["step3_col_a_cells_filled"] += 1
                    case 1:
                        metrics["step3_col_b_cells_filled"] += 1
                    case 2:
                        metrics["step3_col_c_cells_filled"] += 1

        step3_rows.append(new_row)

    # -------------------------------------------------------------
    # Step 4: Delete lines where transaction columns D to Q are empty
    # -------------------------------------------------------------
    step4_rows: list[list[str]] = []
    for row in step3_rows:
        # Columns D to Q are indices 3 to 16 inclusive
        tx_cells = [row[i].strip() if i < len(row) else "" for i in range(3, 17)]
        if all(cell == "" for cell in tx_cells):
            metrics["step4_empty_tx_rows_deleted"] += 1
            continue
        step4_rows.append(row)

    metrics["step4_retained_rows"] = len(step4_rows)

    # -------------------------------------------------------------
    # Step 5: Delete rows where Column A contains 'הפרשה'
    # -------------------------------------------------------------
    step5_rows: list[list[str]] = []
    for row in step4_rows:
        col_a = row[0].strip() if row else ""
        if "הפרשה" in col_a:
            metrics["step5_hafrasha_rows_deleted"] += 1
            continue
        step5_rows.append(row)

    metrics["step5_retained_rows"] = len(step5_rows)

    # -------------------------------------------------------------
    # Step 6: Delete rows where Column A contains 'סהכ לדוח'
    # -------------------------------------------------------------
    step6_rows: list[list[str]] = []
    for row in step5_rows:
        col_a = row[0].strip() if row else ""
        col_a_clean = (
            col_a.replace('"', "")
            .replace("'", "")
            .replace("״", "")
            .replace("׳", "")
            .strip()
        )
        if "סהכ לדוח" in col_a_clean:
            metrics["step6_report_totals_deleted"] += 1
            continue
        step6_rows.append(row)

    metrics["step6_final_emitted_data_rows"] = len(step6_rows)

    # -------------------------------------------------------------
    # Step 7: Delete columns D, E, F, G, H, N, Q
    # -------------------------------------------------------------
    # Exclude indices 3, 4, 5, 6, 7, 13, 16
    del_indices = {3, 4, 5, 6, 7, 13, 16}
    keep_indices = [i for i in range(len(headers)) if i not in del_indices]

    final_headers = [headers[i] for i in keep_indices]
    final_rows: list[list[str]] = []
    for row in step6_rows:
        projected = [row[i] if i < len(row) else "" for i in keep_indices]
        final_rows.append(projected)

    metrics["step7_columns_retained"] = len(final_headers)
    metrics["step7_columns_deleted"] = len(del_indices)
    # -------------------------------------------------------------
    # Step 8: Filter 'קוד מיון'
    # Keep only:
    # - 90001 <= קוד מיון <= 90100 (inclusive)
    # - קוד מיון == 90800
    # - 93000 <= קוד מיון <= 94501 (inclusive)
    # -------------------------------------------------------------
    # 'קוד מיון' is at index 2 of final_headers
    mion_col_idx = final_headers.index("קוד מיון") if "קוד מיון" in final_headers else 2
    step8_rows: list[list[str]] = []
    dropped_step8_count = 0

    for row in final_rows:
        val_str = row[mion_col_idx].strip() if len(row) > mion_col_idx else ""
        try:
            val_int = int(float(val_str))
            keep = (
                (90001 <= val_int <= 90100)
                or (val_int == 90800)
                or (93000 <= val_int <= 94501)
            )
        except ValueError:
            keep = False

        if keep:
            step8_rows.append(row)
        else:
            dropped_step8_count += 1

    metrics["step8_mion_rows_retained"] = len(step8_rows)
    metrics["step8_mion_rows_deleted"] = dropped_step8_count

    # -------------------------------------------------------------
    # Step 9: Filter Column B (מפתח חשבון) provision accounts
    # Delete rows where Column B has 'ה-' / '-ה' (e.g. 920156-ה)
    # -------------------------------------------------------------
    account_col_idx = final_headers.index("מפתח חשבון") if "מפתח חשבון" in final_headers else 1
    step9_rows: list[list[str]] = []
    dropped_step9_count = 0

    for row in step8_rows:
        b_val = row[account_col_idx].strip() if len(row) > account_col_idx else ""
        # Check both logical prefixes/suffixes due to RTL rendering variations
        if b_val.startswith("ה-") or b_val.endswith("-ה") or b_val.endswith("ה-") or "-ה" in b_val:
            dropped_step9_count += 1
            continue
        step9_rows.append(row)

    metrics["step9_col_b_hafrasha_deleted"] = dropped_step9_count
    metrics["step9_final_emitted_rows"] = len(step9_rows)

    # Write final cleaned rows directly to output CSV
    with open(output_path, mode="w", encoding="utf-8-sig", newline="") as f_out:
        writer = csv.writer(f_out)
        writer.writerow(final_headers)
        writer.writerows(step9_rows)

    return metrics


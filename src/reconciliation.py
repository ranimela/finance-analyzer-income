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

    # Locate Column F (index 5) and Column J (index 9)
    col_f_idx = 5
    if "אסמ'" in header:
        col_f_idx = header.index("אסמ'")

    col_j_idx = 9
    if "חובה / זכות (שקל) זכות" in header:
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
    }

    for row in data_rows:
        raw_f = row[col_f_idx].strip() if len(row) > col_f_idx else ""
        raw_j = row[col_j_idx].strip() if len(row) > col_j_idx else ""

        try:
            val_f = float(raw_f)
            is_positive = val_f > 0
        except ValueError:
            is_positive = False

        occ_str = ""
        s_date_str = ""
        e_date_str = ""

        if not is_positive:
            metrics["zero_or_non_positive_rows"] += 1
        else:
            inv_key = raw_f
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
        duration_str = ""
        adj_income_str = ""

        if (
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

        enriched_rows.append(
            row + [occ_str, s_date_str, e_date_str, duration_str, adj_income_str]
        )

    with open(output_csv_path, mode="w", encoding="utf-8-sig", newline="") as f_out:
        writer = csv.writer(f_out)
        writer.writerows(enriched_rows)

    metrics["total_output_rows"] = len(enriched_rows) - 1
    return metrics

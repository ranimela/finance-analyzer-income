"""Validation test suite for Phase 1 Income Ledger Cleaning Pipeline."""

import csv
from pathlib import Path
import pytest

from src.cleaner import clean_income_ledger


def test_clean_income_ledger_pipeline(tmp_path: Path) -> None:
    """Test full 6-step cleaning pipeline on real input file."""
    input_file = Path("data/inputs/כרטסות הכנסות 24-26.xlsx")
    assert input_file.exists(), f"Input file not found at {input_file}"

    output_csv = tmp_path / "test_stage1_v2.csv"
    output_class_csv = tmp_path / "test_classification_v2.csv"
    output_canc_csv = tmp_path / "test_cancellations_v2.csv"
    output_xlsx = tmp_path / "test_stage1_v2.xlsx"
    metrics = clean_income_ledger(
        input_file,
        output_csv,
        output_cancellations_path=output_canc_csv,
        output_classification_path=output_class_csv,
        output_xlsx_path=output_xlsx,
    )

    # 1. Verify metrics accounting
    assert metrics["step0_raw_rows_read"] == 9725
    assert metrics["step1_metadata_skipped"] == 3
    assert metrics["step1_opening_balance_removed"] == 757
    assert metrics["step1_closing_balance_removed"] == 975
    assert metrics["step1_retained_rows"] == 7989

    assert metrics["step2_subtotal_anchor_matches"] == 757
    assert metrics["step2_follower_rows_deleted"] == 1514
    assert metrics["step2_retained_rows"] == 5718

    assert metrics["step4_empty_tx_rows_deleted"] == 759
    assert metrics["step4_retained_rows"] == 4959

    assert metrics["step5_hafrasha_rows_deleted"] == 879
    assert metrics["step5_retained_rows"] == 4080

    assert metrics["step6_report_totals_deleted"] == 3
    assert metrics["step6_final_emitted_data_rows"] == 4077
    assert metrics["step7_columns_retained"] == 10
    assert metrics["step7_columns_deleted"] == 7
    assert metrics["step8_mion_rows_retained"] == 1455
    assert metrics["step8_mion_rows_deleted"] == 2622
    assert metrics["step9_col_b_hafrasha_deleted"] == 88
    assert metrics["step9_final_emitted_rows"] == 1367
    assert metrics["step10_total_column_added"] is True
    assert metrics["step11_closing_entries_deleted"] == 12
    assert metrics["step12_classification_target_invoices"] == 74
    assert metrics["step12_classification_rows_cut"] == 150
    assert metrics["step13_cancellation_notices_found"] == 91
    assert metrics["step13_target_invoices_extracted"] == 57
    assert metrics["step13_original_invoices_cut"] == 64
    assert metrics["step13_total_cancellations_cut"] == 145
    assert metrics["step13_stage1_final_active_rows"] == 1060

    # 2. Verify output CSV file properties (Active clean rows)
    assert output_csv.exists()
    with open(output_csv, mode="r", encoding="utf-8-sig", newline="") as f:
        reader = list(csv.reader(f))

    assert len(reader) == 1061  # 1 header + 1060 active data rows

    header = reader[0]
    expected_headers = [
        "שם חשבון",
        "מפתח חשבון",
        "קוד מיון",
        "ת.אסמכ",
        "ת.ערך",
        "אסמ'",
        "אסמ'2",
        "פרטים",
        "חובה / זכות (שקל) חובה",
        "חובה / זכות (שקל) זכות",
        "Total",
    ]
    assert header == expected_headers

    # 3. Verify income classification CSV file properties
    assert output_class_csv.exists()
    with open(output_class_csv, mode="r", encoding="utf-8-sig", newline="") as f_cl:
        reader_class = list(csv.reader(f_cl))
    assert len(reader_class) == 151  # 1 header + 150 cut classification rows
    assert reader_class[0] == expected_headers

    # 4. Verify cancellations CSV file properties
    assert output_canc_csv.exists()
    with open(output_canc_csv, mode="r", encoding="utf-8-sig", newline="") as f_c:
        reader_canc = list(csv.reader(f_c))
    assert len(reader_canc) == 146  # 1 header + 145 cut rows
    assert reader_canc[0] == expected_headers

    # 5. Verify Excel workbook (3 tabs)
    import openpyxl
    wb = openpyxl.load_workbook(output_xlsx)
    assert wb.sheetnames == ["כרטסות הכנסות פעילות", "income classification", "Cancellations"]
    assert wb["כרטסות הכנסות פעילות"].max_row == 1061
    assert wb["income classification"].max_row == 151
    assert wb["Cancellations"].max_row == 146

    # 5. Verify clean content rules across all emitted active rows
    for row_idx, row in enumerate(reader[1:], 2):
        col_a, col_b, col_c = row[0], row[1], row[2]
        # Forward fill check: account columns must never be blank
        assert col_a != "", f"Row {row_idx}: Col A ('שם חשבון') is empty"
        assert col_b != "", f"Row {row_idx}: Col B ('מפתח חשבון') is empty"
        assert col_c != "", f"Row {row_idx}: Col C ('קוד מיון') is empty"

        # Banned values checks
        assert "יתרת פתיחה" not in col_c, f"Row {row_idx}: contains יתרת פתיחה"
        assert 'סה"כ מפתח חשבון' not in col_a, f"Row {row_idx}: contains subtotal anchor"
        assert "הפרשה" not in col_a, f"Row {row_idx}: contains הפרשה"
        assert "סהכ לדוח" not in col_a.replace('"', "").replace("'", ""), f"Row {row_idx}: contains report total"

        # Verify Total calculation: Total = Col J (Credit) - Col I (Debit)
        val_i_str = row[8].replace(",", "").strip()
        val_j_str = row[9].replace(",", "").strip()
        total_str = row[10].strip()
        val_i = float(val_i_str) if val_i_str else 0.0
        val_j = float(val_j_str) if val_j_str else 0.0
        expected_total = round(val_j - val_i, 2)
        assert abs(float(total_str) - expected_total) < 0.01, f"Row {row_idx}: Total mismatch"

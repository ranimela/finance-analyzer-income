"""Validation test suite for Phase 1 Income Ledger Cleaning Pipeline."""

import csv
from pathlib import Path
import pytest

from src.cleaner import clean_income_ledger


def test_clean_income_ledger_pipeline(tmp_path: Path) -> None:
    """Test full 6-step cleaning pipeline on real input file."""
    input_file = Path("data/inputs/כרטסות הכנסות 24-26.xlsx")
    assert input_file.exists(), f"Input file not found at {input_file}"

    output_csv = tmp_path / "test_step6.csv"
    metrics = clean_income_ledger(input_file, output_csv)

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

    # 2. Verify output CSV file properties
    assert output_csv.exists()
    with open(output_csv, mode="r", encoding="utf-8-sig", newline="") as f:
        reader = list(csv.reader(f))

    assert len(reader) == 1368  # 1 header + 1367 data rows

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
    ]
    assert header == expected_headers

    # 3. Verify clean content rules across all emitted rows
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

        # Check closing balance isn't present in details column (index 12)
        if len(row) > 12:
            assert "יתרת סגירה" not in row[12], f"Row {row_idx}: contains יתרת סגירה"

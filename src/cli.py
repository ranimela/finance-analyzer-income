import argparse
import json
import sys
from pathlib import Path

from src.cleaner import clean_income_ledger
from src.reconciliation import generate_income_working_file, process_cancellations


def main() -> None:
    """CLI entrypoint."""
    parser = argparse.ArgumentParser(
        description="Finance Analyzer Income: Clean ledger and generate working file."
    )
    subparsers = parser.add_subparsers(dest="command", help="Command to run")

    # Clean command (Stage 1)
    clean_parser = subparsers.add_parser("clean", help="Clean raw income ledger Excel file (Stage 1).")
    clean_parser.add_argument(
        "--input",
        type=Path,
        default=Path("data/inputs/כרטסות הכנסות 24-26.xlsx"),
        help="Path to raw Excel input file.",
    )
    clean_parser.add_argument(
        "--output",
        type=Path,
        default=Path("data/outputs/כרטסות הכנסות 24-26_stage1_v2.csv"),
        help="Path for cleaned Stage 1 active output CSV file.",
    )
    clean_parser.add_argument(
        "--classification",
        type=Path,
        default=Path("data/outputs/כרטסות הכנסות 24-26_classification_v2.csv"),
        help="Path for income classification output CSV file.",
    )
    clean_parser.add_argument(
        "--cancellations",
        type=Path,
        default=Path("data/outputs/כרטסות הכנסות 24-26_cancellations_v2.csv"),
        help="Path for cancellations output CSV file.",
    )
    clean_parser.add_argument(
        "--output-xlsx",
        type=Path,
        default=Path("data/outputs/כרטסות הכנסות 24-26_stage1_v2.xlsx"),
        help="Path for Stage 1 multi-tab Excel workbook.",
    )

    # Reconcile command
    recon_parser = subparsers.add_parser("reconcile", help="Generate income_working.csv.")
    recon_parser.add_argument(
        "--ledger",
        type=Path,
        default=Path("data/outputs/כרטסות הכנסות 24-26_step9.csv"),
        help="Path to cleaned ledger CSV.",
    )
    recon_parser.add_argument(
        "--agreements",
        type=Path,
        default=Path("data/inputs/Agreements Payments.xlsx"),
        help="Path to Agreements Payments XLSX.",
    )
    recon_parser.add_argument(
        "--output",
        type=Path,
        default=Path("data/outputs/income_working.csv"),
        help="Path for income_working.csv output.",
    )

    # Cancel command
    cancel_parser = subparsers.add_parser(
        "cancel", help="Filter cancellations into Cancellations tab."
    )
    cancel_parser.add_argument(
        "--input",
        type=Path,
        default=Path("data/outputs/income_working.csv"),
        help="Path to income_working.csv.",
    )
    cancel_parser.add_argument(
        "--output-csv",
        type=Path,
        default=Path("data/outputs/income_working_step2.csv"),
        help="Path for income_working_step2.csv output.",
    )
    cancel_parser.add_argument(
        "--output-xlsx",
        type=Path,
        default=Path("data/outputs/income_working_step2.xlsx"),
        help="Path for income_working_step2.xlsx output.",
    )

    args = parser.parse_args()

    sys.stdout.reconfigure(encoding="utf-8")
    try:
        if args.command == "clean" or args.command is None:
            input_p = getattr(args, "input", Path("data/inputs/כרטסות הכנסות 24-26.xlsx"))
            output_p = getattr(args, "output", Path("data/outputs/כרטסות הכנסות 24-26_stage1_v2.csv"))
            class_p = getattr(args, "classification", Path("data/outputs/כרטסות הכנסות 24-26_classification_v2.csv"))
            canc_p = getattr(args, "cancellations", Path("data/outputs/כרטסות הכנסות 24-26_cancellations_v2.csv"))
            xlsx_p = getattr(args, "output_xlsx", Path("data/outputs/כרטסות הכנסות 24-26_stage1_v2.xlsx"))
            metrics = clean_income_ledger(
                input_p,
                output_p,
                output_cancellations_path=canc_p,
                output_classification_path=class_p,
                output_xlsx_path=xlsx_p,
            )
            print("\n--- Stage 1 Cleaning & Classification Pipeline Success ---")
            print(f"Active Output CSV:       {output_p}")
            print(f"Classification CSV:      {class_p}")
            print(f"Cancellations CSV:       {canc_p}")
            print(f"Stage 1 Workbook:        {xlsx_p}")
            print(json.dumps(metrics, indent=2, ensure_ascii=False))

        if args.command == "reconcile" or args.command is None:
            ledger_p = getattr(args, "ledger", Path("data/outputs/כרטסות הכנסות 24-26_step9.csv"))
            agreements_p = getattr(args, "agreements", Path("data/inputs/Agreements Payments.xlsx"))
            recon_output_p = getattr(args, "output", Path("data/outputs/income_working.csv"))
            r_metrics = generate_income_working_file(ledger_p, agreements_p, recon_output_p)
            print("\n--- Reconciliation Working File Success ---")
            print(f"Output File: {recon_output_p}")
            print(json.dumps(r_metrics, indent=2, ensure_ascii=False))

        if args.command == "cancel" or args.command is None:
            in_p = getattr(args, "input", Path("data/outputs/income_working.csv"))
            out_csv_p = getattr(args, "output_csv", Path("data/outputs/income_working_step2.csv"))
            out_xlsx_p = getattr(args, "output_xlsx", Path("data/outputs/income_working_step2.xlsx"))
            c_metrics = process_cancellations(
                in_p,
                out_csv_p,
                output_cancellations_csv_path=Path("data/outputs/cancellations.csv"),
                output_step2_xlsx_path=out_xlsx_p,
            )
            print("\n--- Cancellations Processing Success ---")
            print(f"Output CSV: {out_csv_p}")
            print(f"Output XLSX: {out_xlsx_p}")
            print(json.dumps(c_metrics, indent=2, ensure_ascii=False))

    except Exception as exc:
        print(f"Error during execution: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

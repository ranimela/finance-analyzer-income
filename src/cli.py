import argparse
import json
import sys
from pathlib import Path

from src.cleaner import clean_income_ledger
from src.reconciliation import generate_income_working_file


def main() -> None:
    """CLI entrypoint."""
    parser = argparse.ArgumentParser(
        description="Finance Analyzer Income: Clean ledger and generate working file."
    )
    subparsers = parser.add_subparsers(dest="command", help="Command to run")

    # Clean command
    clean_parser = subparsers.add_parser("clean", help="Clean raw income ledger Excel file.")
    clean_parser.add_argument(
        "--input",
        type=Path,
        default=Path("data/inputs/כרטסות הכנסות 24-26.xlsx"),
        help="Path to raw Excel input file.",
    )
    clean_parser.add_argument(
        "--output",
        type=Path,
        default=Path("data/outputs/כרטסות הכנסות 24-26_step9.csv"),
        help="Path for cleaned Step 9 output CSV file.",
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

    args = parser.parse_args()

    sys.stdout.reconfigure(encoding="utf-8")
    try:
        if args.command == "clean" or args.command is None:
            input_p = getattr(args, "input", Path("data/inputs/כרטסות הכנסות 24-26.xlsx"))
            output_p = getattr(args, "output", Path("data/outputs/כרטסות הכנסות 24-26_step9.csv"))
            metrics = clean_income_ledger(input_p, output_p)
            print("\n--- Cleaning Pipeline Success ---")
            print(f"Output File: {output_p}")
            print(json.dumps(metrics, indent=2, ensure_ascii=False))

        if args.command == "reconcile" or args.command is None:
            ledger_p = getattr(args, "ledger", Path("data/outputs/כרטסות הכנסות 24-26_step9.csv"))
            agreements_p = getattr(args, "agreements", Path("data/inputs/Agreements Payments.xlsx"))
            recon_output_p = getattr(args, "output", Path("data/outputs/income_working.csv"))
            r_metrics = generate_income_working_file(ledger_p, agreements_p, recon_output_p)
            print("\n--- Reconciliation Working File Success ---")
            print(f"Output File: {recon_output_p}")
            print(json.dumps(r_metrics, indent=2, ensure_ascii=False))

    except Exception as exc:
        print(f"Error during execution: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

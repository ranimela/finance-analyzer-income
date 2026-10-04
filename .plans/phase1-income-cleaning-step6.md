# Architectural Specification: Phase 1 Ledger Cleaning Pipeline (הכנסות)

## Architectural Overview
This specification adapts and executes the deterministic 6-step cleaning pipeline designed in `finance-analyzer-provisions` to the income ledger file `data/inputs/כרטסות הכנסות 24-26.xlsx`.
Per architecture decisions:
1. The cleaning pipeline applies the exact six filtering stages in memory/streaming.
2. Only the final clean CSV (`data/outputs/כרטסות הכנסות 24-26_step6.csv`) is emitted to disk, avoiding redundant intermediate writes.
3. The project source code adheres to modern Python 3.10+ typing, strict formatting, and stateless execution.
4. Data files remain excluded from git tracking via `.gitignore`.

## Immutable Data Contracts

### Target Output Contract (`data/outputs/כרטסות הכנסות 24-26_step6.csv`)
- **Encoding:** `utf-8-sig` (ensuring correct Excel viewing for Hebrew characters).
- **Line Terminating:** CRLF / standard newline dialect.
- **Header Columns (17 columns):**
  1. `שם חשבון` (Synthetic from Column A / Row 4)
  2. `מפתח חשבון` (Synthetic from Column B / Row 4)
  3. `קוד מיון` (Synthetic from Column C / Row 4)
  4. `כותרת` (Row 4 Col D)
  5. `תנועה` (Row 4 Col E)
  6. `מנה` (Row 4 Col F)
  7. `ס"ת` (Row 4 Col G)
  8. `ח-ן נגדי` (Row 4 Col H)
  9. `ת.אסמכ` (Row 4 Col I)
  10. `ת.ערך` (Row 4 Col J)
  11. `אסמ'` (Row 4 Col K)
  12. `אסמ'2` (Row 4 Col L)
  13. `פרטים` (Row 4 Col M)
  14. `תמחיר` (Row 4 Col N)
  15. `חובה / זכות (שקל) חובה` (Row 4 Col O)
  16. `חובה / זכות (שקל) זכות` (Row 4 Col P)
  17. `יתרה (שקל)` (Row 4 Col Q)

## 6-Step Cleaning Algorithm
1. **Step 1 - Metadata & Balance Strip:**
   - Skip rows 1–3 (report titles).
   - Read row 4 as column headers, injecting synthetic names for Col A (`שם חשבון`), Col B (`מפתח חשבון`), Col C (`קוד מיון`).
   - Discard rows where Column C (`קוד מיון`) == `'יתרת פתיחה'`.
   - Discard rows where Column M (`פרטים`) == `'יתרת סגירה'`.
2. **Step 2 - Subtotal Anchor & Follower Removal:**
   - Locate rows where Column A equals `'סה"כ מפתח חשבון'` or contains `'סה״כ מפתח חשבון'`.
   - Remove the anchor row plus up to 3 subsequent follower rows if their Column A is empty.
3. **Step 3 - Forward Fill:**
   - Forward-fill empty cells in Columns A (`שם חשבון`), B (`מפתח חשבון`), and C (`קוד מיון`) using the most recent non-empty value.
4. **Step 4 - Empty Transaction Strip:**
   - Discard any row where all columns from D to Q (indices 3 to 16) are empty strings / None.
5. **Step 5 - Hafrasha Filter:**
   - Discard rows where Column A (`שם חשבון`) contains `'הפרשה'`.
6. **Step 6 - Report Totals Strip:**
   - Discard rows where Column A normalized contains `'סהכ לדוח'` or `'סה"כ לדוח'`.
   - Emit the final dataset to `data/outputs/כרטסות הכנסות 24-26_step6.csv`.

## Affected Files
- `src/__init__.py` (Package root)
- `src/cleaner.py` (Core cleaning engine)
- `src/cli.py` (Console runner)
- `.plans/phase1-income-cleaning-step6.md` (This blueprint)

## Step-by-Step Micro-Tasks for Builder
1. Create `src/__init__.py` and implement `src/cleaner.py` with typed, documented methods for `clean_income_ledger()` implementing Steps 1–6.
2. Implement CLI command interface in `src/cli.py` accepting `--input` and `--output` arguments with sensible defaults pointing to `data/inputs/כרטסות הכנסות 24-26.xlsx` and `data/outputs/כרטסות הכנסות 24-26_step6.csv`.
3. Execute the CLI to generate `data/outputs/כרטסות הכנסות 24-26_step6.csv`.
4. Run validation checks verifying:
   - Output row count matches expected ~4,077 data rows (plus header).
   - Zero occurrences of `'יתרת פתיחה'`, `'יתרת סגירה'`, `'סה"כ מפתח חשבון'`, or `'הפרשה'` in output.
   - Header integrity with all 17 target columns in utf-8-sig.

## Verification Criteria
- Automated python test / verification script confirms:
  - Input: 9,725 rows -> Emitted: 4,077 data rows (+ 1 header row = 4,078 total rows).
  - Valid UTF-8-sig encoding.
  - Proper forward-filled account codes in cols A-C across all transaction rows.

## Context Pruning
The Builder is permitted to read only:
- `c:\Users\rmelamed\Projects\finance-analyzer-income\.plans\phase1-income-cleaning-step6.md`
- `c:\Users\rmelamed\Projects\finance-analyzer-provisions\src\cleaner.py` (reference only)
Ignoring all other files.

# WL_Recruitment_Analysis

This repo uses an OpenAI LLM classifier to assign one to three UN Sustainable
Development Goals (SDGs) to each professor research profile.

## OpenAI SDG classifier

Use this for the current small workbook.

1. Open `.env`.
2. Replace `replace_with_your_openai_api_key` with your real OpenAI API key.
3. Run:

   ```bash
   python classify_sdgs_openai.py
   ```

   To classify a specific file:

   ```bash
   python classify_sdgs_openai.py vw_analysis_recruitment.csv
   ```

   To run faster with parallel API workers:

   ```bash
   python classify_sdgs_openai.py vw_analysis_recruitment.csv --workers 5
   ```

   To also classify multiple rows per API call:

   ```bash
   python classify_sdgs_openai.py vw_analysis_recruitment.csv --workers 5 --batch-size 3
   ```

   To resume from a checkpoint output:

   ```bash
   python classify_sdgs_openai.py vw_analysis_recruitment.csv --resume-from outputs/vw_analysis_recruitment_with_openai_sdgs_YYYYMMDD_HHMMSS.xlsx
   ```

Input:

- Any `.xlsx`, `.xls`, or `.csv` file. If no file is provided, the script uses
  `professor_research.xlsx`.

Excel output:

- `outputs/<input_file_name>_with_openai_sdgs_YYYYMMDD_HHMMSS.xlsx`

Classification column:

- `sdg`

The input workbook already has an `sdg` column, so the script writes the LLM
classification there. If a future workbook does not have an `sdg` column, the
script creates only that one column.

The script keeps only `Medium` and `High` confidence SDGs, with at least 1 and
at most 3 SDGs per row. If no SDG reaches `Medium` confidence, it keeps the
single strongest `Low` confidence SDG as a fallback so the row is not blank.

Manual review is set by deterministic Python rules after thresholding. It does
not come from keyword matching. A row needs manual review when:

- no SDG is returned
- only a fallback `Low` confidence SDG is available

Run log output:

- `logs/openai_sdg_YYYYMMDD_HHMMSS.log`

Each run logs the input file, output file, model, worker count, rows per API
call, number of rows detected, periodic progress, retry warnings, checkpoint
saves, rows needing manual review, and final row counts. It does not log full per-row
rationale/evidence, so logs stay readable for large workbooks.

During long runs, the script retries transient API failures and saves checkpoint
Excel output every 25 processed rows. If a later row still fails, the output file
keeps the rows completed before the failure. Rerun with `--resume-from` to skip
rows that already have an `sdg` value in the checkpoint.

The script uses the standard API because the current workbook has only 6 rows.
For a later 2,000-row run, use the same prompt/schema but submit the requests
through OpenAI Batch API to reduce cost.

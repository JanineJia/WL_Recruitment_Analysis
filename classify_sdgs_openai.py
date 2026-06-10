"""
Classify professor research rows into UN Sustainable Development Goals using OpenAI.

This script uses LLM classification only, not keyword matching.

Before running:
    1. Open .env
    2. Replace replace_with_your_openai_api_key with your real API key

Run:
    python classify_sdgs_openai.py
    python classify_sdgs_openai.py vw_analysis_recruitment.csv
    python classify_sdgs_openai.py vw_analysis_recruitment.csv --workers 1
    python classify_sdgs_openai.py vw_analysis_recruitment.csv --workers 2

Output:
    outputs/<input_file_name>_with_openai_sdgs_YYYYMMDD_HHMMSS.xlsx
    logs/openai_sdg_YYYYMMDD_HHMMSS.log

The Excel output uses the existing sdg column if present. If the input workbook
does not have an sdg column, the script creates only that one classification
column.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

import pandas as pd


DEFAULT_INPUT_FILE = "professor_research.xlsx"
OUTPUT_DIR = "outputs"
MODEL = "gpt-5.4-nano"
ENV_FILE = ".env"
LOG_DIR = "logs"

PROGRESS_EVERY_ROWS = 10
PROGRESS_DIVIDER = "-" * 72
CHECKPOINT_EVERY_ROWS = 25

OPENAI_TIMEOUT_SECONDS = 120
RETRY_ATTEMPTS = 5
RETRY_BACKOFF_SECONDS = 20

MAX_SDGS_PER_ROW = 3
COUNTED_CONFIDENCES = {"High", "Medium"}

TEXT_COLUMNS = [
    "title",
    "dpt_profile",
    "research_tags",
    "grants",
    "scopus_publications",
    "DPT",
    "affiliation",
    "specializations",
    "faculty",
]

# Reduced from 6000 to make rate limits less likely.
MAX_TEXT_CHARS = 3000
MAX_OUTPUT_TOKENS = 800

# Fixes openpyxl IllegalCharacterError.
ILLEGAL_EXCEL_CHARACTERS = re.compile(r"[\000-\010]|[\013-\014]|[\016-\037]")

SDG_NAMES = {
    1: "No Poverty",
    2: "Zero Hunger",
    3: "Good Health and Well-being",
    4: "Quality Education",
    5: "Gender Equality",
    6: "Clean Water and Sanitation",
    7: "Affordable and Clean Energy",
    8: "Decent Work and Economic Growth",
    9: "Industry, Innovation and Infrastructure",
    10: "Reduced Inequalities",
    11: "Sustainable Cities and Communities",
    12: "Responsible Consumption and Production",
    13: "Climate Action",
    14: "Life Below Water",
    15: "Life on Land",
    16: "Peace, Justice and Strong Institutions",
    17: "Partnerships for the Goals",
}

SDG_RESPONSE_SCHEMA: Dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "matched_sdgs": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "sdg_number": {
                        "type": "integer",
                        "enum": list(SDG_NAMES.keys()),
                    },
                    "sdg_name": {
                        "type": "string",
                        "enum": list(SDG_NAMES.values()),
                    },
                    "confidence": {
                        "type": "string",
                        "enum": ["High", "Medium", "Low"],
                    },
                    "rationale": {"type": "string"},
                    "evidence": {"type": "string"},
                },
                "required": [
                    "sdg_number",
                    "sdg_name",
                    "confidence",
                    "rationale",
                    "evidence",
                ],
            },
        },
        "overall_notes": {"type": "string"},
    },
    "required": ["matched_sdgs", "overall_notes"],
}

BATCH_SDG_RESPONSE_SCHEMA: Dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "rows": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "row_number": {"type": "integer"},
                    "matched_sdgs": SDG_RESPONSE_SCHEMA["properties"]["matched_sdgs"],
                    "overall_notes": {"type": "string"},
                },
                "required": ["row_number", "matched_sdgs", "overall_notes"],
            },
        }
    },
    "required": ["rows"],
}

SYSTEM_PROMPT = """
You classify university professor research profiles into UN Sustainable
Development Goals (SDGs).

Use only these 17 predetermined SDGs:
1 No Poverty
2 Zero Hunger
3 Good Health and Well-being
4 Quality Education
5 Gender Equality
6 Clean Water and Sanitation
7 Affordable and Clean Energy
8 Decent Work and Economic Growth
9 Industry, Innovation and Infrastructure
10 Reduced Inequalities
11 Sustainable Cities and Communities
12 Responsible Consumption and Production
13 Climate Action
14 Life Below Water
15 Life on Land
16 Peace, Justice and Strong Institutions
17 Partnerships for the Goals

Classify based on substantive research connection, not incidental words.
Return at least 1 and at most 3 SDGs for one professor.
Order SDGs from strongest to weakest.
Only include Medium or High confidence SDGs.
If no SDG reaches Medium confidence, return the single strongest Low confidence SDG.
Use High only when the SDG is central to the profile.
Use Medium when the SDG is clearly supported but not central.
Use Low when the evidence is plausible but weak.
Do not use keyword matching. Classify by research meaning and evidence.
""".strip()


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def remove_illegal_excel_chars(value: object) -> object:
    """Remove hidden control characters that Excel/openpyxl cannot write."""
    if isinstance(value, str):
        return ILLEGAL_EXCEL_CHARACTERS.sub("", value)
    return value


def clean_dataframe_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy of df that is safe to write to .xlsx."""
    return df.map(remove_illegal_excel_chars)


def safe_to_excel(df: pd.DataFrame, output_path: Path) -> None:
    """Save dataframe to Excel after removing illegal characters."""
    clean_df = clean_dataframe_for_excel(df)
    clean_df.to_excel(output_path, index=False, engine="openpyxl")


def load_env_file(path: str = ENV_FILE) -> None:
    if not os.path.exists(path):
        return

    with open(path, "r", encoding="utf-8") as env_file:
        for line in env_file:
            clean_line = line.strip()
            if not clean_line or clean_line.startswith("#") or "=" not in clean_line:
                continue

            name, value = clean_line.split("=", 1)
            name = name.strip()
            value = value.strip().strip('"').strip("'")
            if name:
                os.environ[name] = value


def setup_logger() -> logging.Logger:
    Path(LOG_DIR).mkdir(exist_ok=True)
    log_path = Path(LOG_DIR) / f"openai_sdg_{datetime.now():%Y%m%d_%H%M%S}.log"

    logger = logging.getLogger("openai_sdg_classifier")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    logger.info("Log file: %s", log_path)
    return logger


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Classify professor research rows into SDGs with OpenAI."
    )
    parser.add_argument(
        "input_file",
        nargs="?",
        default=DEFAULT_INPUT_FILE,
        help=f"Input .xlsx or .csv file. Default: {DEFAULT_INPUT_FILE}",
    )
    parser.add_argument(
        "--resume-from",
        help="Existing checkpoint .xlsx output to continue from.",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=1,
        help="Number of parallel API workers. Default: 1",
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=1,
        help="Rows to classify per API call. Default: 1",
    )
    return parser.parse_args()


def build_output_path(input_file: str) -> Path:
    Path(OUTPUT_DIR).mkdir(exist_ok=True)
    input_stem = Path(input_file).stem
    return (
        Path(OUTPUT_DIR)
        / f"{input_stem}_with_openai_sdgs_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    )


def read_input_file(input_file: str) -> pd.DataFrame:
    input_path = Path(input_file)
    suffix = input_path.suffix.lower()

    if suffix == ".csv":
        return pd.read_csv(input_path)
    if suffix in {".xlsx", ".xls"}:
        return pd.read_excel(input_path)

    raise ValueError("Input file must be .xlsx, .xls, or .csv")


def load_resume_sdg_values(df: pd.DataFrame, resume_file: str) -> pd.Series:
    resume_df = pd.read_excel(resume_file)
    if len(resume_df) != len(df):
        raise ValueError("Resume file row count does not match input file row count.")
    if "sdg" not in resume_df.columns:
        raise ValueError("Resume file does not contain an sdg column.")
    return resume_df["sdg"].fillna("").astype(str)


def build_profile_text(row: pd.Series, columns: List[str]) -> str:
    parts = []

    for column in columns:
        if column in row.index:
            value = clean_text(row[column])
            if value:
                parts.append(f"{column}: {value}")

    extra_values = []
    for column in row.index:
        if str(column).startswith("Unnamed"):
            value = clean_text(row[column])
            if value:
                extra_values.append(value)

    if extra_values:
        parts.append("extra_text: " + " ".join(extra_values))

    profile_text = "\n".join(parts)

    if len(profile_text) > MAX_TEXT_CHARS:
        return profile_text[:MAX_TEXT_CHARS] + "\n[truncated for cost control]"

    return profile_text


def extract_response_text(response_data: Dict[str, Any]) -> str:
    if response_data.get("output_text"):
        return str(response_data["output_text"])

    for output_item in response_data.get("output", []):
        for content_item in output_item.get("content", []):
            if content_item.get("type") in {"output_text", "text"}:
                return str(content_item.get("text", ""))

    raise RuntimeError("OpenAI response did not contain output text.")


def call_openai(profile_text: str, api_key: str) -> Dict[str, Any]:
    payload = {
        "model": MODEL,
        "input": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": "Classify this professor profile:\n\n" + profile_text,
            },
        ],
        "text": {
            "format": {
                "type": "json_schema",
                "name": "sdg_classification",
                "strict": True,
                "schema": SDG_RESPONSE_SCHEMA,
            }
        },
        "max_output_tokens": MAX_OUTPUT_TOKENS,
    }

    request = Request(
        "https://api.openai.com/v1/responses",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        with urlopen(request, timeout=OPENAI_TIMEOUT_SECONDS) as response:
            response_data = json.loads(response.read().decode("utf-8"))
    except HTTPError as error:
        message = error.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"OpenAI API request failed: {message}") from error

    response_text = extract_response_text(response_data)
    return json.loads(response_text)


def is_rate_limit_error(error: Exception) -> bool:
    message = str(error).lower()
    return "rate limit" in message or "tokens per min" in message or "tpm" in message


def call_openai_with_retries(
    profile_text: str,
    api_key: str,
    logger: logging.Logger,
    row_index: int,
    total_rows: int,
) -> Dict[str, Any]:
    for attempt in range(1, RETRY_ATTEMPTS + 1):
        try:
            return call_openai(profile_text, api_key)

        except (TimeoutError, URLError, RuntimeError) as error:
            if attempt == RETRY_ATTEMPTS:
                raise

            if is_rate_limit_error(error):
                wait_seconds = 60 * attempt
            else:
                wait_seconds = RETRY_BACKOFF_SECONDS * attempt

            logger.warning(
                "Retry row %s/%s after API error attempt %s/%s | waiting %ss | %s: %s",
                row_index,
                total_rows,
                attempt,
                RETRY_ATTEMPTS,
                wait_seconds,
                type(error).__name__,
                str(error)[:300],
            )

            time.sleep(wait_seconds)

    raise RuntimeError("OpenAI retry loop exited unexpectedly.")


def call_openai_batch(rows: List[tuple[int, pd.Series]], api_key: str) -> Dict[int, Dict[str, Any]]:
    row_blocks = []
    expected_row_numbers = set()

    for row_number, row in rows:
        row_index = row_number + 1
        expected_row_numbers.add(row_index)
        row_blocks.append(
            f"ROW {row_index}\n{build_profile_text(row, TEXT_COLUMNS)}"
        )

    payload = {
        "model": MODEL,
        "input": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": (
                    "Classify each professor row below. Return one result for each "
                    "ROW number and preserve the same row_number values.\n\n"
                    + "\n\n---\n\n".join(row_blocks)
                ),
            },
        ],
        "text": {
            "format": {
                "type": "json_schema",
                "name": "sdg_batch_classification",
                "strict": True,
                "schema": BATCH_SDG_RESPONSE_SCHEMA,
            }
        },
        "max_output_tokens": MAX_OUTPUT_TOKENS * len(rows),
    }

    request = Request(
        "https://api.openai.com/v1/responses",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        with urlopen(request, timeout=OPENAI_TIMEOUT_SECONDS) as response:
            response_data = json.loads(response.read().decode("utf-8"))
    except HTTPError as error:
        message = error.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"OpenAI API request failed: {message}") from error

    response_text = extract_response_text(response_data)
    parsed = json.loads(response_text)

    results_by_row = {
        int(row_result["row_number"]): {
            "matched_sdgs": row_result["matched_sdgs"],
            "overall_notes": row_result["overall_notes"],
        }
        for row_result in parsed["rows"]
    }

    if set(results_by_row) != expected_row_numbers:
        raise RuntimeError("OpenAI batch response did not return every requested row.")

    return results_by_row


def call_openai_batch_with_retries(
    rows: List[tuple[int, pd.Series]],
    api_key: str,
    logger: logging.Logger,
    total_rows: int,
) -> Dict[int, Dict[str, Any]]:
    row_numbers = [row_number + 1 for row_number, _ in rows]
    row_range = f"{min(row_numbers)}-{max(row_numbers)}"

    for attempt in range(1, RETRY_ATTEMPTS + 1):
        try:
            return call_openai_batch(rows, api_key)

        except (TimeoutError, URLError, RuntimeError) as error:
            if attempt == RETRY_ATTEMPTS:
                raise

            if is_rate_limit_error(error):
                wait_seconds = 60 * attempt
            else:
                wait_seconds = RETRY_BACKOFF_SECONDS * attempt

            logger.warning(
                "Retry batch rows %s/%s after API error attempt %s/%s | waiting %ss | %s: %s",
                row_range,
                total_rows,
                attempt,
                RETRY_ATTEMPTS,
                wait_seconds,
                type(error).__name__,
                str(error)[:300],
            )

            time.sleep(wait_seconds)

    raise RuntimeError("OpenAI batch retry loop exited unexpectedly.")


def format_sdg_numbers(result: Dict[str, Any]) -> str:
    return "; ".join(f"SDG {match['sdg_number']}" for match in result["matched_sdgs"])


def apply_sdg_threshold(result: Dict[str, Any]) -> Dict[str, Any]:
    matches = result["matched_sdgs"]
    counted_matches = [
        match for match in matches if match["confidence"] in COUNTED_CONFIDENCES
    ]

    if counted_matches:
        result["matched_sdgs"] = counted_matches[:MAX_SDGS_PER_ROW]
    else:
        result["matched_sdgs"] = matches[:1]

    return result


def manual_review_reason(result: Dict[str, Any]) -> str:
    matches = result["matched_sdgs"]

    if not matches:
        return "No SDG matches returned"

    confidence_values = [match["confidence"] for match in matches]

    if all(confidence == "Low" for confidence in confidence_values):
        return "Only fallback Low confidence SDG available"

    return ""


def needs_manual_review(result: Dict[str, Any]) -> bool:
    return manual_review_reason(result) != ""


def log_progress_checkpoint(logger: logging.Logger, row_index: int, total_rows: int) -> None:
    logger.info(PROGRESS_DIVIDER)
    logger.info("Progress checkpoint: %s/%s rows completed", row_index, total_rows)


def save_checkpoint(
    df: pd.DataFrame,
    output_path: Path,
    logger: logging.Logger,
    row_index: int,
    total_rows: int,
) -> None:
    safe_to_excel(df, output_path)
    logger.info(
        "Checkpoint saved after %s/%s rows completed: %s",
        row_index,
        total_rows,
        output_path,
    )


def classify_row(
    row_number: int,
    row: pd.Series,
    api_key: str,
    logger: logging.Logger,
    total_rows: int,
) -> Dict[str, Any]:
    profile_text = build_profile_text(row, TEXT_COLUMNS)
    row_index = row_number + 1

    result = call_openai_with_retries(
        profile_text,
        api_key,
        logger,
        row_index,
        total_rows,
    )

    return apply_sdg_threshold(result)


def classify_batch(
    rows: List[tuple[int, pd.Series]],
    api_key: str,
    logger: logging.Logger,
    total_rows: int,
) -> Dict[int, Dict[str, Any]]:
    if len(rows) == 1:
        row_number, row = rows[0]
        result = classify_row(row_number, row, api_key, logger, total_rows)
        return {row_number + 1: result}

    results_by_row = call_openai_batch_with_retries(rows, api_key, logger, total_rows)

    return {
        row_index: apply_sdg_threshold(result)
        for row_index, result in results_by_row.items()
    }


def chunk_rows(
    rows: List[tuple[int, pd.Series]],
    batch_size: int,
) -> List[List[tuple[int, pd.Series]]]:
    return [
        rows[index : index + batch_size]
        for index in range(0, len(rows), batch_size)
    ]


def main() -> None:
    args = parse_args()
    input_file = args.input_file

    logger = setup_logger()

    output_path = (
        Path(args.resume_from)
        if args.resume_from
        else build_output_path(input_file)
    )

    load_env_file()

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key or api_key == "replace_with_your_openai_api_key":
        raise RuntimeError("Set OPENAI_API_KEY in .env before running this script.")

    logger.info("Starting OpenAI SDG classification")
    logger.info("Input file: %s", input_file)
    logger.info("Output file: %s", output_path)

    if args.resume_from:
        logger.info("Resume file: %s", args.resume_from)

    logger.info("Model: %s", MODEL)
    logger.info("Max text chars per row: %s", MAX_TEXT_CHARS)
    logger.info("Max SDGs per row: %s", MAX_SDGS_PER_ROW)

    df = read_input_file(input_file)
    df = df.dropna(axis=1, how="all")

    total_rows = len(df)

    logger.info("Rows detected: %s", total_rows)
    logger.info("Text columns: %s", ", ".join(TEXT_COLUMNS))

    if "sdg" in df.columns:
        non_empty_sdg_rows = df["sdg"].notna().sum()
        logger.info(
            "Existing sdg column found; replacing values with LLM classifications. "
            "Non-empty existing sdg cells before replacement: %s",
            non_empty_sdg_rows,
        )
    else:
        logger.info("No sdg column found; creating sdg classification column.")

    df["sdg"] = ""

    if args.resume_from:
        df["sdg"] = load_resume_sdg_values(df, args.resume_from)
        logger.info(
            "Rows already classified in resume file: %s",
            df["sdg"].astype(str).str.strip().ne("").sum(),
        )

    pending_rows = []

    for row_number, row in df.iterrows():
        if clean_text(row.get("sdg", "")):
            continue
        pending_rows.append((row_number, row.copy()))

    completed_rows = total_rows - len(pending_rows)

    worker_count = max(1, args.workers)
    batch_size = max(1, args.batch_size)

    pending_batches = chunk_rows(pending_rows, batch_size)

    logger.info("API workers: %s", worker_count)
    logger.info("Rows per API call: %s", batch_size)
    logger.info("Rows to classify in this run: %s", len(pending_rows))

    if completed_rows:
        logger.info("Rows skipped because sdg already exists: %s", completed_rows)

    results = []

    executor = ThreadPoolExecutor(max_workers=worker_count)
    executor_closed = False

    try:
        future_to_batch = {
            executor.submit(
                classify_batch,
                batch,
                api_key,
                logger,
                total_rows,
            ): batch
            for batch in pending_batches
        }

        for future in as_completed(future_to_batch):
            batch = future_to_batch[future]

            try:
                batch_results = future.result()

            except Exception:
                first_row_number = batch[0][0]
                row_index = first_row_number + 1
                row = df.loc[first_row_number]
                title = clean_text(row.get("title", f"row {row_index}"))

                logger.exception(
                    "Failed batch starting row %s/%s | title=%s",
                    row_index,
                    total_rows,
                    title[:80],
                )

                save_checkpoint(df, output_path, logger, completed_rows, total_rows)

                executor.shutdown(wait=False, cancel_futures=True)
                executor_closed = True
                raise

            for row_index, result in sorted(batch_results.items()):
                row_number = row_index - 1
                row = df.loc[row_number]
                title = clean_text(row.get("title", f"row {row_index}"))

                matched_sdgs = format_sdg_numbers(result) or "none"
                df.at[row_number, "sdg"] = matched_sdgs

                review_needed = needs_manual_review(result)
                review_reason = manual_review_reason(result) or "none"

                if review_needed:
                    logger.warning(
                        "Manual review row %s/%s | title=%s | sdgs=%s | reason=%s",
                        row_index,
                        total_rows,
                        title[:80],
                        matched_sdgs,
                        review_reason,
                    )

                results.append(result)
                completed_rows += 1

                if (
                    completed_rows == 1
                    or completed_rows == total_rows
                    or completed_rows % PROGRESS_EVERY_ROWS == 0
                ):
                    log_progress_checkpoint(logger, completed_rows, total_rows)

                if completed_rows % CHECKPOINT_EVERY_ROWS == 0:
                    save_checkpoint(df, output_path, logger, completed_rows, total_rows)

    finally:
        if not executor_closed:
            executor.shutdown(wait=True, cancel_futures=False)

    safe_to_excel(df, output_path)

    rows_with_match = df["sdg"].astype(str).str.strip().ne("").sum()
    rows_needing_review = sum(1 for result in results if needs_manual_review(result))

    logger.info("Summary")
    logger.info("Rows processed: %s", total_rows)
    logger.info("Rows with SDG classifications: %s", rows_with_match)
    logger.info("Rows needing manual review in this run: %s", rows_needing_review)
    logger.info("Saved classified results to: %s", output_path)


if __name__ == "__main__":
    main()
"""
Classify Excel rows into UN Sustainable Development Goals (SDGs).

How to use this script:

1. Put this script in the same folder as your Excel file, or provide the full
   path to the Excel file in INPUT_FILE.

2. Change INPUT_FILE to your Excel file name.

3. Change OUTPUT_FILE to the file name you want to create.

4. Set HEADER_ROW to the row number that contains your column titles.
   Important: Python uses zero-based indexing.
   - If your title/header row is row 1 in Excel, use HEADER_ROW = 0
   - If your title/header row is row 2 in Excel, use HEADER_ROW = 1
   - If your title/header row is row 3 in Excel, use HEADER_ROW = 2

5. Run the script once with PRINT_ONLY_COLUMNS = True if you only want to
   inspect the detected column names first.

6. Add the columns you want to classify into SELECTED_TEXT_COLUMNS.
   Examples:
       SELECTED_TEXT_COLUMNS = [
           "Professor Name",
           "Research Topic",
           "Research Tags",
           "Publication Title",
           "Publication Abstract",
           "Department",
           "Keywords",
       ]

7. Run:
       python classify_sdgs.py

Required packages:
    pip install pandas openpyxl

This script does not use OpenAI, ChatGPT, or any external paid API.
It runs locally using keyword matching.
"""

from __future__ import annotations

import re
from typing import Dict, List, Tuple

import pandas as pd


# ---------------------------------------------------------------------------
# User settings: edit these values for your spreadsheet
# ---------------------------------------------------------------------------

INPUT_FILE = "professor_research.xlsx"
OUTPUT_FILE = "professor_research_with_sdgs.xlsx"

# Python counts from 0, so Excel row 1 is HEADER_ROW = 0.
HEADER_ROW = 0

# Set to True when you only want to print the detected columns and preview rows.
# Set to False when you are ready to classify and save results.
PRINT_ONLY_COLUMNS = False

# Put the exact column names you want to use for SDG classification here.
# The script prints detected column names first so you can copy/paste them.
SELECTED_TEXT_COLUMNS = [
    "research_tags",
    "specializations",
    "dpt_profile",
    "title",
]

# If SELECTED_TEXT_COLUMNS is empty, this option can automatically use likely
# text columns. For best results, keep this False and choose columns yourself.
AUTO_USE_LIKELY_TEXT_COLUMNS_IF_EMPTY = False

# If True, rows with only Low confidence matches are also flagged for review.
REVIEW_LOW_CONFIDENCE = False

# Confidence thresholds based on the number of matched strong phrases per SDG.
HIGH_CONFIDENCE_MIN_KEYWORDS = 3
MEDIUM_CONFIDENCE_MIN_KEYWORDS = 1


# ---------------------------------------------------------------------------
# SDG keyword dictionary
# Edit or expand these lists to better match your research area.
#
# strong_phrases: specific research phrases that can directly identify an SDG.
# weak_terms: broad terms that are useful hints, but are too vague on their own.
# ---------------------------------------------------------------------------

SDG_KEYWORDS: Dict[int, Dict[str, object]] = {
    1: {
        "name": "No Poverty",
        "strong_phrases": [
            "poverty reduction",
            "poverty alleviation",
            "low-income communities",
            "low income households",
            "income insecurity",
            "economic insecurity",
            "financial hardship",
            "social assistance",
            "social protection",
            "welfare policy",
            "affordable housing",
            "housing affordability",
            "homelessness",
            "basic services access",
            "financial inclusion",
            "economic vulnerability",
        ],
        "weak_terms": ["poverty", "welfare"],
    },
    2: {
        "name": "Zero Hunger",
        "strong_phrases": [
            "food security",
            "food insecurity",
            "sustainable agriculture",
            "food systems",
            "agricultural productivity",
            "crop production",
            "soil fertility",
            "food supply chain",
            "nutrition security",
            "malnutrition",
            "agroecology",
            "smallholder farmers",
            "sustainable food production",
            "urban agriculture",
            "food access",
            "food affordability",
        ],
        "weak_terms": ["hunger", "nutrition", "agriculture"],
    },
    3: {
        "name": "Good Health and Well-being",
        "strong_phrases": [
            "public health",
            "global health",
            "mental health",
            "health equity",
            "health outcomes",
            "healthcare access",
            "health care access",
            "clinical outcomes",
            "disease prevention",
            "infectious disease",
            "chronic disease",
            "maternal health",
            "child health",
            "community health",
            "population health",
            "epidemiology",
            "health policy",
            "substance use",
            "patient care",
        ],
        "weak_terms": ["health", "disease", "medicine", "clinical", "cancer", "patient"],
    },
    4: {
        "name": "Quality Education",
        "strong_phrases": [
            "quality education",
            "educational access",
            "education equity",
            "inclusive education",
            "learning outcomes",
            "student learning",
            "student success",
            "curriculum development",
            "teacher training",
            "higher education",
            "early childhood education",
            "educational technology",
            "online learning",
            "literacy development",
            "numeracy development",
            "skills development",
            "pedagogical practice",
        ],
        "weak_terms": ["education", "learning", "teaching", "curriculum", "literacy"],
    },
    5: {
        "name": "Gender Equality",
        "strong_phrases": [
            "gender equality",
            "gender equity",
            "gender inequality",
            "women's rights",
            "women empowerment",
            "women's empowerment",
            "gender-based violence",
            "domestic violence",
            "sexual harassment",
            "reproductive rights",
            "female leadership",
            "gender representation",
            "feminist theory",
            "queer studies",
            "lgbtq inclusion",
            "intersectional gender",
        ],
        "weak_terms": ["gender", "women", "girls", "feminism", "lgbtq", "queer"],
    },
    6: {
        "name": "Clean Water and Sanitation",
        "strong_phrases": [
            "clean water",
            "drinking water",
            "water quality",
            "water treatment",
            "wastewater treatment",
            "water management",
            "water resources",
            "water pollution",
            "freshwater systems",
            "groundwater contamination",
            "watershed management",
            "sanitation systems",
            "hygiene access",
            "desalination",
            "hydrological systems",
            "safe drinking water",
        ],
        "weak_terms": ["water", "sanitation", "wastewater", "groundwater", "freshwater", "hygiene"],
    },
    7: {
        "name": "Affordable and Clean Energy",
        "strong_phrases": [
            "clean energy",
            "renewable energy",
            "solar energy",
            "wind energy",
            "energy efficiency",
            "energy storage",
            "smart grid",
            "power grid",
            "battery technology",
            "fuel cells",
            "hydrogen energy",
            "energy transition",
            "energy access",
            "electricity access",
            "decarbonized energy",
            "electrification",
            "low-carbon energy",
        ],
        "weak_terms": ["energy", "solar", "electricity", "battery", "fuel cell"],
    },
    8: {
        "name": "Decent Work and Economic Growth",
        "strong_phrases": [
            "decent work",
            "labor market",
            "labour market",
            "employment outcomes",
            "job creation",
            "workforce development",
            "workplace safety",
            "occupational health",
            "economic growth",
            "economic development",
            "productivity growth",
            "entrepreneurship",
            "small business development",
            "youth employment",
            "informal employment",
            "sustainable tourism",
            "labor rights",
            "labour rights",
        ],
        "weak_terms": ["employment", "labor", "labour", "worker", "workforce", "productivity"],
    },
    9: {
        "name": "Industry, Innovation and Infrastructure",
        "strong_phrases": [
            "resilient infrastructure",
            "sustainable infrastructure",
            "digital infrastructure",
            "transport infrastructure",
            "industrial innovation",
            "technological innovation",
            "research and development",
            "r&d investment",
            "advanced manufacturing",
            "smart manufacturing",
            "automation systems",
            "robotics",
            "materials science",
            "engineering innovation",
            "broadband access",
            "industrial development",
            "infrastructure planning",
        ],
        "weak_terms": ["infrastructure", "industry", "innovation", "technology", "manufacturing", "engineering", "broadband"],
    },
    10: {
        "name": "Reduced Inequalities",
        "strong_phrases": [
            "social inequality",
            "income inequality",
            "racial inequality",
            "health inequality",
            "education inequality",
            "reduced inequalities",
            "social inclusion",
            "economic inclusion",
            "marginalized communities",
            "underrepresented groups",
            "disability inclusion",
            "accessibility",
            "migration policy",
            "refugee integration",
            "indigenous communities",
            "racial justice",
            "equity and inclusion",
            "systemic discrimination",
        ],
        "weak_terms": ["inequality", "inequalities", "equity", "inclusion", "discrimination", "marginalized", "refugee", "indigenous"],
    },
    11: {
        "name": "Sustainable Cities and Communities",
        "strong_phrases": [
            "sustainable cities",
            "urban planning",
            "urban development",
            "urban sustainability",
            "smart cities",
            "public transportation",
            "sustainable transportation",
            "affordable housing",
            "housing policy",
            "built environment",
            "walkability",
            "urban resilience",
            "climate-resilient cities",
            "community resilience",
            "municipal planning",
            "disaster risk reduction",
            "heritage conservation",
            "land use planning",
        ],
        "weak_terms": ["city", "cities", "urban", "community", "housing", "transportation", "municipal"],
    },
    12: {
        "name": "Responsible Consumption and Production",
        "strong_phrases": [
            "responsible consumption",
            "sustainable production",
            "circular economy",
            "waste reduction",
            "food waste",
            "plastic waste",
            "recycling systems",
            "resource efficiency",
            "life cycle assessment",
            "sustainable supply chain",
            "sustainable procurement",
            "consumer behavior",
            "ethical consumption",
            "sustainable packaging",
            "reuse systems",
            "industrial waste",
            "environmental footprint",
        ],
        "weak_terms": ["consumption", "production", "recycling", "waste", "packaging", "reuse"],
    },
    13: {
        "name": "Climate Action",
        "strong_phrases": [
            "climate change",
            "climate action",
            "climate policy",
            "climate adaptation",
            "climate mitigation",
            "climate resilience",
            "greenhouse gas emissions",
            "carbon emissions",
            "carbon reduction",
            "carbon capture",
            "net zero",
            "decarbonization",
            "low-carbon transition",
            "extreme weather",
            "sea level rise",
            "climate risk",
            "climate justice",
            "global warming",
        ],
        "weak_terms": ["climate", "carbon", "emissions", "mitigation", "adaptation"],
    },
    14: {
        "name": "Life Below Water",
        "strong_phrases": [
            "marine ecosystems",
            "marine biodiversity",
            "marine conservation",
            "ocean conservation",
            "ocean acidification",
            "marine pollution",
            "plastic pollution",
            "coastal ecosystems",
            "coral reefs",
            "fisheries management",
            "sustainable fisheries",
            "blue economy",
            "aquatic ecosystems",
            "estuaries",
            "seawater systems",
            "marine protected areas",
            "ocean health",
        ],
        "weak_terms": ["ocean", "marine", "coastal", "aquatic", "fisheries", "estuaries", "seawater"],
    },
    15: {
        "name": "Life on Land",
        "strong_phrases": [
            "terrestrial ecosystems",
            "biodiversity conservation",
            "forest conservation",
            "deforestation",
            "land use change",
            "habitat loss",
            "wildlife conservation",
            "endangered species",
            "ecosystem restoration",
            "restoration ecology",
            "soil conservation",
            "desertification",
            "protected areas",
            "invasive species",
            "agroforestry",
            "species conservation",
            "land degradation",
            "ecosystem services",
        ],
        "weak_terms": ["biodiversity", "ecosystem", "forest", "land use", "conservation", "wildlife", "habitat", "species", "soil"],
    },
    16: {
        "name": "Peace, Justice and Strong Institutions",
        "strong_phrases": [
            "peace and justice",
            "strong institutions",
            "human rights",
            "rule of law",
            "criminal justice",
            "access to justice",
            "legal institutions",
            "democratic governance",
            "public accountability",
            "anti-corruption",
            "conflict resolution",
            "political violence",
            "institutional trust",
            "public policy governance",
            "security governance",
            "justice system",
            "law and society",
        ],
        "weak_terms": ["peace", "justice", "institution", "governance", "law", "legal", "conflict", "violence", "corruption"],
    },
    17: {
        "name": "Partnerships for the Goals",
        "strong_phrases": [
            "global partnership",
            "international cooperation",
            "development cooperation",
            "multi-stakeholder collaboration",
            "stakeholder engagement",
            "capacity building",
            "knowledge sharing",
            "technology transfer",
            "policy coherence",
            "sustainable development goals",
            "sdg implementation",
            "united nations",
            "development finance",
            "cross-sector partnership",
            "community partnership",
            "research collaboration for development",
        ],
        "weak_terms": ["collaboration", "partnership", "development", "stakeholder", "sdgs"],
    },
}


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def normalize_column_name(column_name: object) -> str:
    """Convert a spreadsheet column name to a clean string."""
    return str(column_name).strip()


def clean_text(value: object) -> str:
    """Safely convert any cell value into searchable text."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def combine_text_columns(row: pd.Series, text_columns: List[str]) -> str:
    """Combine selected columns into one text field for a row."""
    parts = [clean_text(row[column]) for column in text_columns]
    return " ".join(part for part in parts if part)


def remove_empty_and_unnamed_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Remove blank Excel columns that pandas labels as Unnamed."""
    # Remove columns where every cell is empty.
    df = df.dropna(axis=1, how="all")

    # Remove columns with no real header, such as "Unnamed: 16".
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

    return df


def keyword_found(text: str, keyword: str) -> bool:
    """
    Check whether a keyword or phrase appears in text.

    The regex boundaries help avoid accidental partial matches.
    For example, "law" should match "law" but not "lawn".
    """
    pattern = r"(?<!\w)" + re.escape(keyword.lower()) + r"(?!\w)"
    return re.search(pattern, text.lower()) is not None


def confidence_from_matches(strong_count: int, weak_count: int) -> str:
    """Convert strong/weak keyword counts into a confidence label."""
    if strong_count >= HIGH_CONFIDENCE_MIN_KEYWORDS:
        return "High"
    if strong_count >= MEDIUM_CONFIDENCE_MIN_KEYWORDS:
        return "Medium"
    if weak_count > 0:
        return "Low"
    return ""


def classify_text(text: str) -> List[Dict[str, object]]:
    """
    Classify one combined text value into zero, one, or many SDGs.

    Strong phrases can create High/Medium matches.
    Weak terms only create Low-confidence possible matches for review.
    """
    matches = []

    if not text.strip():
        return matches

    for sdg_number, sdg_info in SDG_KEYWORDS.items():
        matched_strong_phrases = []
        matched_weak_terms = []

        for phrase in sdg_info["strong_phrases"]:
            if keyword_found(text, phrase):
                matched_strong_phrases.append(phrase)

        for term in sdg_info["weak_terms"]:
            if keyword_found(text, term):
                matched_weak_terms.append(term)

        # Remove duplicates while preserving order.
        unique_strong_phrases = list(dict.fromkeys(matched_strong_phrases))
        unique_weak_terms = list(dict.fromkeys(matched_weak_terms))
        confidence = confidence_from_matches(
            len(unique_strong_phrases), len(unique_weak_terms)
        )

        if confidence:
            matches.append(
                {
                    "sdg_number": sdg_number,
                    "sdg_name": sdg_info["name"],
                    "matched_keywords": unique_strong_phrases + unique_weak_terms,
                    "matched_strong_phrases": unique_strong_phrases,
                    "matched_weak_terms": unique_weak_terms,
                    "confidence": confidence,
                }
            )

    # Sort results by SDG number so the output is stable and easy to read.
    return sorted(matches, key=lambda item: int(item["sdg_number"]))


def format_sdg_numbers(matches: List[Dict[str, object]]) -> str:
    """Format matched SDG numbers for the Excel output column."""
    return "; ".join(f"SDG {match['sdg_number']}" for match in matches)


def format_sdg_names(matches: List[Dict[str, object]]) -> str:
    """Format matched SDG names for the Excel output column."""
    return "; ".join(str(match["sdg_name"]) for match in matches)


def format_keywords(matches: List[Dict[str, object]]) -> str:
    """Format matched keywords by SDG for the Excel output column."""
    return "; ".join(
        f"SDG {match['sdg_number']}: {', '.join(match['matched_keywords'])}"
        for match in matches
    )


def format_confidence(matches: List[Dict[str, object]]) -> str:
    """Format confidence by SDG for the Excel output column."""
    return "; ".join(
        f"SDG {match['sdg_number']}: {match['confidence']}" for match in matches
    )


def keep_confirmed_matches(matches: List[Dict[str, object]]) -> List[Dict[str, object]]:
    """Keep only High and Medium confidence SDG matches."""
    return [
        match
        for match in matches
        if match["confidence"] in {"High", "Medium"}
    ]


def keep_review_matches(matches: List[Dict[str, object]]) -> List[Dict[str, object]]:
    """Keep only Low confidence SDG matches for manual review."""
    return [match for match in matches if match["confidence"] == "Low"]


def format_review_matches(matches: List[Dict[str, object]]) -> str:
    """Format Low confidence SDGs with their names and keywords."""
    return "; ".join(
        f"SDG {match['sdg_number']} - {match['sdg_name']}: "
        f"{', '.join(match['matched_keywords'])}"
        for match in matches
    )


def likely_text_columns(columns: List[str]) -> List[str]:
    """
    Guess likely research-related text columns.

    This is optional and only used when AUTO_USE_LIKELY_TEXT_COLUMNS_IF_EMPTY
    is True and SELECTED_TEXT_COLUMNS is empty.
    """
    likely_terms = [
        "name",
        "professor",
        "faculty",
        "research",
        "topic",
        "tag",
        "title",
        "abstract",
        "department",
        "keyword",
        "publication",
        "description",
        "interest",
        "area",
    ]

    guessed_columns = []
    for column in columns:
        column_lower = column.lower()
        if any(term in column_lower for term in likely_terms):
            guessed_columns.append(column)
    return guessed_columns


def validate_selected_columns(
    selected_columns: List[str], available_columns: List[str]
) -> Tuple[List[str], List[str]]:
    """
    Check that the selected columns exist in the dataframe.

    Returns:
    - valid columns
    - missing columns
    """
    available_set = set(available_columns)
    valid_columns = [column for column in selected_columns if column in available_set]
    missing_columns = [column for column in selected_columns if column not in available_set]
    return valid_columns, missing_columns


# ---------------------------------------------------------------------------
# Main script
# ---------------------------------------------------------------------------

def main() -> None:
    """Read the Excel file, classify rows, and save the result."""
    print(f"Reading Excel file: {INPUT_FILE}")
    df = pd.read_excel(INPUT_FILE, header=HEADER_ROW, engine="openpyxl")

    # Remove blank Excel columns before printing, classifying, or saving.
    df = remove_empty_and_unnamed_columns(df)

    # Clean column names so accidental leading/trailing spaces are less painful.
    df.columns = [normalize_column_name(column) for column in df.columns]
    detected_columns = list(df.columns)

    print("\nDetected column names:")
    for index, column in enumerate(detected_columns, start=1):
        print(f"{index}. {column}")

    print("\nFirst 5 rows:")
    print(df.head())

    if PRINT_ONLY_COLUMNS:
        print("\nPRINT_ONLY_COLUMNS is True, so no classification was performed.")
        print("Set PRINT_ONLY_COLUMNS = False when you are ready to classify rows.")
        return

    selected_columns = SELECTED_TEXT_COLUMNS.copy()

    if not selected_columns and AUTO_USE_LIKELY_TEXT_COLUMNS_IF_EMPTY:
        selected_columns = likely_text_columns(detected_columns)
        print("\nSELECTED_TEXT_COLUMNS is empty.")
        print("Automatically selected likely text columns:")
        for column in selected_columns:
            print(f"- {column}")

    valid_columns, missing_columns = validate_selected_columns(
        selected_columns, detected_columns
    )

    if missing_columns:
        print("\nWarning: these selected columns were not found:")
        for column in missing_columns:
            print(f"- {column}")

    if not valid_columns:
        print("\nNo valid text columns were selected.")
        print("Please copy exact column names from the detected column list into")
        print("SELECTED_TEXT_COLUMNS near the top of this script.")
        return

    print("\nUsing these columns for SDG classification:")
    for column in valid_columns:
        print(f"- {column}")

    # Combine selected text columns into one searchable field per row.
    df["_combined_sdg_text"] = df.apply(
        lambda row: combine_text_columns(row, valid_columns), axis=1
    )

    # Classify each row.
    all_matches = df["_combined_sdg_text"].apply(classify_text)

    # Split matches into confirmed matches and possible review matches.
    # High / Medium confidence matches go into matched_sdgs.
    # Low confidence matches go into possible_sdgs_for_review.
    confirmed_matches = all_matches.apply(keep_confirmed_matches)
    review_matches = all_matches.apply(keep_review_matches)

    # Add requested output columns.
    df["matched_sdgs"] = confirmed_matches.apply(format_sdg_numbers)
    df["matched_sdg_names"] = confirmed_matches.apply(format_sdg_names)
    df["matched_keywords"] = confirmed_matches.apply(format_keywords)
    df["confidence"] = confirmed_matches.apply(format_confidence)
    df["possible_sdgs_for_review"] = review_matches.apply(format_review_matches)
    df["manual_review_needed"] = review_matches.apply(
        lambda matches: len(matches) > 0
    )

    # Remove the temporary combined text column before saving.
    df = df.drop(columns=["_combined_sdg_text"])

    rows_processed = len(df)
    rows_with_match = (df["matched_sdgs"].str.strip() != "").sum()
    rows_needing_review = df["manual_review_needed"].sum()

    print("\nDebugging summary:")
    print(f"Number of rows processed: {rows_processed}")
    print(f"Number of rows that received at least one SDG match: {rows_with_match}")
    print(f"Number of rows that need manual review: {rows_needing_review}")

    print(f"\nSaving classified results to: {OUTPUT_FILE}")
    df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    print("Done.")


if __name__ == "__main__":
    main()

import pandas as pd
import os
import time
import random
import re
import unicodedata
from tqdm import tqdm
from dotenv import load_dotenv
from llm_provider import LLMProvider

load_dotenv()

try:
    llm_provider = LLMProvider()
except Exception as error:
    print(f"⚠️ LLM setup failed: {error}")
    llm_provider = None

THEME_KNOWLEDGE_BASE = """
1. Poverty and Economic Barriers: Financial hardship, child labour. Keywords: Poor, no money.
2. Legal Document-linked Barriers: Missing Aadhaar, birth certificates. Keywords: No Aadhar, no ID.
3. Child Marriage Issue: Early marriage preventing education. Keywords: Child marriage.
4. Distance and Accessibility Issues: School far, bad roads, weather. Keywords: Far, no bus, rain.
5. Parental Attitudes & Socio-Cultural: Mindsets against girls, dowry, domestic roles.
6. School Infrastructure & Facility: Toilets, water, Mid-day meals, books, govt schemes.
7. Teacher Capacity & Quality: Shortage of teachers, irregular attendance.
8. Safety Issues: Harassment, unsafe routes, stray dogs.
9. Substance Abuse & Addiction: Alcohol, drugs, gambling, mobile addiction.
10. Other Factors: General awareness, migration. (Target <10%)
"""

MAX_RETRIES = int(os.getenv("LLM_MAX_RETRIES", "5"))
BASE_RETRY_SECONDS = float(os.getenv("LLM_RETRY_BASE_SECONDS", "2"))
MAX_RETRY_SECONDS = float(os.getenv("LLM_RETRY_MAX_SECONDS", "45"))
TAGGER_BATCH_SIZE = int(os.getenv("TAGGER_BATCH_SIZE", "200"))
TAGGER_MAX_TOKENS = int(os.getenv("TAGGER_MAX_TOKENS", "12000"))
TAGGER_THINKING_BUDGET = int(os.getenv("TAGGER_THINKING_BUDGET", "512"))
FUZZY_THEME_MIN_SCORE = float(os.getenv("FUZZY_THEME_MIN_SCORE", "0.92"))
RULE_THEME_MIN_SCORE = float(os.getenv("RULE_THEME_MIN_SCORE", "0.60"))

CANONICAL_THEMES = [
    "Poverty and Economic Barriers",
    "Legal Document-linked Barriers",
    "Child Marriage Issue",
    "Distance and Accessibility Issues",
    "Parental Attitudes & Socio-Cultural",
    "School Infrastructure & Facility",
    "Teacher Capacity & Quality",
    "Safety Issues",
    "Substance Abuse & Addiction",
    "Other Factors",
]

THEME_INDEX_MAP = {str(index + 1): theme for index, theme in enumerate(CANONICAL_THEMES)}

THEME_PATTERN_RULES = {
    "Legal Document-linked Barriers": [
        r"\baad?ha?r\b", r"\baadhar\b", r"\baadhaar\b", r"birth certificate", r"\btc\b",
        r"transfer certificate", r"document", r"id card", r"certificate"
    ],
    "Poverty and Economic Barriers": [
        r"\bpoverty\b", r"poor financial", r"financial", r"\bno money\b", r"lack of money",
        r"economic", r"labou?r", r"wages?", r"income", r"brick kiln", r"field work",
        r"\bthe poor\b", r"\bdebt\b"
    ],
    "Child Marriage Issue": [
        r"child marriage", r"early marriage", r"marr(y|ied|iage).*18", r"under\s*18", r"marry.*young"
    ],
    "Distance and Accessibility Issues": [
        r"\bdistance\b", r"far away", r"far from school", r"transport", r"bicycle", r"road",
        r"access", r"commut", r"nearby village", r"anganwadi.*far", r"far.*anganwadi"
    ],
    "Parental Attitudes & Socio-Cultural": [
        r"parent", r"not aware", r"awareness", r"mindset", r"discrimination", r"household chore",
        r"household work", r"domestic work", r"girls? .*not allowed", r"importance of education",
        r"socio", r"do not send", r"not send.*school", r"going out.*house", r"prevent.*school",
        r"not feel.*stud", r"not attend.*school"
    ],
    "School Infrastructure & Facility": [
        r"toilet", r"drinking water", r"mid[- ]?day meal", r"uniform", r"books?", r"anganwadi",
        r"infrastructure", r"facility", r"school building", r"no school", r"no provision",
        r"higher secondary"
    ],
    "Teacher Capacity & Quality": [
        r"teacher", r"does not teach", r"shortage of teachers", r"teacher absentee", r"teaching quality"
    ],
    "Safety Issues": [
        r"harass", r"unsafe", r"safety", r"fear", r"violence", r"molest", r"kidney",
        r"go.*alone", r"walk.*alone", r"travel.*alone"
    ],
    "Substance Abuse & Addiction": [
        r"alcohol", r"drug", r"addiction", r"gambl", r"drunk", r"mobile addiction"
    ],
}

THEME_PRIORITY = [
    "Legal Document-linked Barriers",
    "Child Marriage Issue",
    "Substance Abuse & Addiction",
    "Safety Issues",
    "Teacher Capacity & Quality",
    "School Infrastructure & Facility",
    "Distance and Accessibility Issues",
    "Poverty and Economic Barriers",
    "Parental Attitudes & Socio-Cultural",
]


def is_retryable_error(error):
    message = str(error).lower()
    retry_signals = [
        "429",
        "rate limit",
        "resource_exhausted",
        "too many requests",
        "throttle",
        "temporarily unavailable",
        "timeout",
        "max_tokens",
        "missing text parts",
        "thinking budget",
    ]
    return any(signal in message for signal in retry_signals)


def save_progress(output_csv, batch_df):
    if batch_df.empty:
        return
    file_exists = os.path.exists(output_csv)
    batch_df.to_csv(output_csv, mode='a', index=False, header=not file_exists)


def normalize_for_match(text):
    normalized = unicodedata.normalize("NFKC", str(text))
    normalized = normalized.replace("’", "'").replace("‘", "'").replace("“", '"').replace("”", '"')
    normalized = re.sub(r"\s+", " ", normalized).strip().lower()
    return normalized


def normalize_theme_name(theme_text):
    if pd.isna(theme_text):
        return None

    raw = str(theme_text).strip()
    if not raw or raw.lower() == "nan":
        return None

    numeric_match = re.fullmatch(r"(\d+)(?:\.0+)?", raw)
    if numeric_match:
        return THEME_INDEX_MAP.get(numeric_match.group(1))

    cleaned = re.sub(r"^\d+[\.)\s-]*", "", raw).strip()
    cleaned_lower = cleaned.lower()

    for canonical in CANONICAL_THEMES:
        if cleaned_lower == canonical.lower():
            return canonical

    alias_patterns = [
        (r"poverty|economic|financial|labou?r|income", "Poverty and Economic Barriers"),
        (r"legal|document|aadhar|aadhaar|birth certificate|tc", "Legal Document-linked Barriers"),
        (r"child marriage|early marriage", "Child Marriage Issue"),
        (r"distance|access|transport|far", "Distance and Accessibility Issues"),
        (r"parental|socio|attitude|mindset|awareness", "Parental Attitudes & Socio-Cultural"),
        (r"infrastructure|facility|toilet|water|anganwadi|mid[- ]?day", "School Infrastructure & Facility"),
        (r"teacher|quality", "Teacher Capacity & Quality"),
        (r"safety|harass|unsafe|violence", "Safety Issues"),
        (r"substance|alcohol|drug|addiction|gambl|mobile", "Substance Abuse & Addiction"),
        (r"other", "Other Factors"),
    ]

    for pattern, canonical in alias_patterns:
        if re.search(pattern, cleaned_lower):
            return canonical

    return None


def classify_theme_rule_based(text):
    text_lower = normalize_for_match(text)
    if text_lower in {"", "none", "na", "n/a", "null"}:
        return "Other Factors", 0.20, "empty_text"

    scores = {}
    for theme, patterns in THEME_PATTERN_RULES.items():
        scores[theme] = sum(1 for pattern in patterns if re.search(pattern, text_lower))

    max_score = max(scores.values()) if scores else 0
    if max_score <= 0:
        return "Other Factors", 0.35, "rule_no_match"

    tied_themes = [theme for theme, score in scores.items() if score == max_score]
    if len(tied_themes) == 1:
        confidence = min(0.95, 0.60 + 0.12 * max_score)
        return tied_themes[0], confidence, "rule_single_match"

    for preferred in THEME_PRIORITY:
        if preferred in tied_themes:
            confidence = min(0.75, 0.50 + 0.08 * max_score)
            return preferred, confidence, "rule_tie_break"

    return "Other Factors", 0.40, "rule_unresolved"


def derive_concept_from_text(text, theme):
    cleaned = str(text).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = cleaned.strip(" .,:;|-")

    if not cleaned:
        return "Uncategorized"

    words = cleaned.split()
    short_phrase = " ".join(words[:8])
    if len(short_phrase) < 4:
        short_phrase = "Uncategorized"

    canonical = short_phrase[0].upper() + short_phrase[1:] if short_phrase else "Uncategorized"
    if canonical == "Uncategorized":
        return canonical

    return canonical


def build_reference_lookup(reference_df):
    """Build an O(1) normalized exact-match dict from reference_df.
    Replaces the O(n) SequenceMatcher scan that caused quadratic slowdown.
    Fully vectorised - safe to call on tens of thousands of existing rows at startup."""
    if reference_df.empty:
        return {}

    df = reference_df.copy()

    # Drop rows where Original is NaN or the string "nan"
    orig_raw = df["Original"]
    valid_mask = orig_raw.notna() & orig_raw.astype(str).str.strip().str.lower().ne("nan")
    df = df[valid_mask].reset_index(drop=True)
    if df.empty:
        return {}

    # Vectorised normalize_for_match using Unicode escapes to avoid editor curly-quote corruption
    norms = (
        df["Original"].astype(str).str.strip()
        .str.normalize("NFKC")
        .str.replace("’", "'", regex=False)
        .str.replace("‘", "'", regex=False)
        .str.replace("“", '"', regex=False)
        .str.replace("”", '"', regex=False)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
        .str.lower()
    )

    # Fast theme path: most themes in existing CSVs are already canonical strings
    canonical_set = set(CANONICAL_THEMES)
    themes_raw = df["Theme"].astype(str).str.strip()
    themes = themes_raw.where(themes_raw.isin(canonical_set))
    fallback_idx = themes[themes.isna()].index
    if len(fallback_idx) > 0:
        themes.loc[fallback_idx] = df.loc[fallback_idx, "Theme"].apply(normalize_theme_name)

    # Vectorised concept cleaning (avoids calling clean_merged_concept per row)
    concepts = df.get("Merged_Concept", pd.Series([""] * len(df))).fillna("").astype(str)
    concepts = (
        concepts.str.strip()
        .str.strip('"')
        .str.strip("'")
        .str.replace(r"\s+", " ", regex=True)
        .str.rstrip(" .,:;")
        .str.replace(r"^\d+[\.\)\s-]*", "", regex=True)
        .str.strip()
    )
    has_content = concepts.str.len() > 0
    concepts = (concepts.str[:1].str.upper() + concepts.str[1:]).where(has_content, "Uncategorized")

    confidences = pd.to_numeric(
        df.get("Confidence", pd.Series([0.9] * len(df))), errors="coerce"
    ).fillna(0.9)

    # Filter valid rows and build dict from numpy arrays (avoids iterrows overhead)
    valid = (norms.str.len() > 0) & themes.notna()
    norms_arr    = norms[valid].to_numpy()
    themes_arr   = themes[valid].to_numpy()
    concepts_arr = concepts[valid].to_numpy()
    confs_arr    = confidences[valid].to_numpy()

    return {
        norm: {
            "Theme": str(theme),
            "Merged_Concept": str(concept),
            "Confidence": float(conf),
            "Mapping_Source": "exact_reference",
        }
        for norm, theme, concept, conf in zip(norms_arr, themes_arr, concepts_arr, confs_arr)
    }
def fuzzy_match_from_reference(text, reference_lookup):
    """O(1) exact-match lookup. Previously O(n) SequenceMatcher — caused 10+ min gaps."""
    if not reference_lookup:
        return None
    norm_text = normalize_for_match(text)
    if not norm_text:
        return None
    return reference_lookup.get(norm_text)


def clean_merged_concept(text):
    """Normalize Merged_Concept formatting for consistent canonical labels."""
    if pd.isna(text):
        return ""

    cleaned = str(text).strip()
    cleaned = cleaned.strip('"').strip("'")
    cleaned = cleaned.replace("\n", " ")
    cleaned = ' '.join(cleaned.split())
    cleaned = cleaned.rstrip(' .,:;')

    # Remove numbering/bullets if model adds them
    cleaned = pd.Series([cleaned]).str.replace(r'^\d+[\.)\s-]*', '', regex=True).iloc[0].strip()

    if not cleaned:
        return "Uncategorized"

    return cleaned[0].upper() + cleaned[1:] if len(cleaned) > 1 else cleaned.upper()


def postprocess_mapping_batch(df_batch):
    """Apply deterministic cleanup so model output stays canonical and machine-usable."""
    if df_batch.empty:
        return df_batch

    df_batch['Original'] = df_batch['Original'].astype(str).str.strip()
    df_batch['Theme'] = df_batch['Theme'].apply(normalize_theme_name)
    df_batch['Merged_Concept'] = df_batch['Merged_Concept'].apply(clean_merged_concept)

    # Remove accidental duplicate rows and keep one mapping per Original
    df_batch = df_batch.drop_duplicates(subset=['Original'], keep='last')
    return df_batch

def get_ai_mapping(text_batch, type_label):
    if not llm_provider:
        raise RuntimeError("LLM provider is not configured")

    indexed_rows = "\n".join([f"{index + 1}\t{text}" for index, text in enumerate(text_batch)])

    prompt_content = f"""Act as an expert Social Data Analyst. Use these THEMES:
    {THEME_KNOWLEDGE_BASE}
    
    SEMANTIC DEDUPLICATION PROTOCOL (MANDATORY):
    You must merge semantically similar items into a single "Merged_Concept".
    
     Rules:
    1. Group all variants expressing the same core issue.
    2. Select ONE canonical phrase and reuse it for all equivalent variants in this batch.
    2A. Before writing output, internally create a canonical dictionary for this batch.
    2B. Use only those dictionary labels in final output (no one-off labels for similar meaning).
     3. 'Merged_Concept' MUST follow canonical naming format:
         - concise noun phrase (3-8 words)
         - no ending punctuation
         - avoid sentence-style wording
         - avoid district/person-specific details
         - stable wording across similar records
    3. Examples of merging:
       - "Due to poverty" = "Due to poor financial condition" = "Lack of money" -> Merged_Concept: "Poverty preventing education"
       - "No Aadhaar card" = "Lack of Aadhaar" = "Aadhaar not made" -> Merged_Concept: "Lack of legal documentation (Aadhaar)"
       - "School is far" = "School is very far" = "Distance of school" -> Merged_Concept: "School distance and accessibility issues"
         - "We will drop children to school" = "Arranging transport to school" = "Parents will take children to school" -> Merged_Concept: "Community-supported school transportation"
     4. CRITICAL: Assign EXACTLY ONE theme from the list. Do not combine themes with '+' or 'and'. If multiple apply, choose the most dominant one.
    
    TASK: Categorize these unique {type_label} statements.
    SELF-CHECK (MANDATORY, still same single call):
    - Re-scan your own output and ensure semantically equivalent rows use exactly identical Merged_Concept text.
    - If two labels differ only by wording (e.g., arranging/providing/facilitating same action), unify them.
    OUTPUT FORMAT (STRICT): Return ONLY with three columns: ID|Theme|Merged_Concept
    - ID must be copied exactly from input rows.
    - Return exactly one row per ID from input.
    - Never omit an ID.
    - Use only one of the 10 canonical themes.
    Use the | character as the delimiter. Do not include headers, preamble, or markdown backticks.
    
    DATA:
    {indexed_rows}"""

    raw_output = llm_provider.generate_text(
        prompt_content,
        max_tokens=TAGGER_MAX_TOKENS,
        temperature=0,
        thinking_budget=TAGGER_THINKING_BUDGET,
    )

    raw_output = raw_output.replace('```csv', '').replace('```', '').strip()

    parsed_rows = []
    for line in raw_output.splitlines():
        line = line.strip()
        if not line or '|' not in line:
            continue
        parts = [part.strip() for part in line.split('|')]
        if len(parts) < 3:
            continue

        id_token = parts[0]
        id_match = re.search(r"\d+", id_token)
        if not id_match:
            continue

        row_id = int(id_match.group())
        if row_id < 1 or row_id > len(text_batch):
            continue

        original = str(text_batch[row_id - 1]).strip()
        theme = normalize_theme_name(parts[1])
        merged_concept = clean_merged_concept("|".join(parts[2:]))
        parsed_rows.append(
            {
                "Original": original,
                "Theme": theme,
                "Merged_Concept": merged_concept,
                "Confidence": 0.90 if theme else 0.0,
                "Mapping_Source": "llm_id_mapping",
            }
        )

    df_batch = pd.DataFrame(parsed_rows)
    return postprocess_mapping_batch(df_batch)


def enforce_batch_coverage(current_items, mapped_df, reference_lookup):
    mapped_df = mapped_df.copy()
    if mapped_df.empty:
        mapped_df = pd.DataFrame(columns=["Original", "Theme", "Merged_Concept", "Confidence", "Mapping_Source"])

    # Normalize existing mapped rows and fix invalid/empty values.
    for index in mapped_df.index:
        original = str(mapped_df.at[index, "Original"]).strip()
        theme = normalize_theme_name(mapped_df.at[index, "Theme"])
        concept = clean_merged_concept(mapped_df.at[index, "Merged_Concept"])

        if not theme:
            fuzzy = fuzzy_match_from_reference(original, reference_lookup)
            if fuzzy:
                theme = fuzzy["Theme"]
                concept = fuzzy["Merged_Concept"]
                mapped_df.at[index, "Confidence"] = fuzzy["Confidence"]
                mapped_df.at[index, "Mapping_Source"] = fuzzy["Mapping_Source"]
            else:
                theme_candidate, rule_conf, rule_source = classify_theme_rule_based(original)
                if rule_conf >= RULE_THEME_MIN_SCORE:
                    theme = theme_candidate
                else:
                    theme = "Other Factors"
                mapped_df.at[index, "Confidence"] = round(rule_conf, 4)
                mapped_df.at[index, "Mapping_Source"] = rule_source

        if not concept or concept == "Uncategorized":
            concept = derive_concept_from_text(original, theme)

        mapped_df.at[index, "Original"] = original
        mapped_df.at[index, "Theme"] = theme
        mapped_df.at[index, "Merged_Concept"] = clean_merged_concept(concept)

    mapped_df = mapped_df.drop_duplicates(subset=["Original"], keep="first")
    mapped_lookup = {str(row["Original"]).strip(): row for _, row in mapped_df.iterrows()}

    filled_rows = []
    for item in current_items:
        original = str(item).strip()
        if original in mapped_lookup:
            row = mapped_lookup[original]
            filled_rows.append(
                {
                    "Original": original,
                    "Theme": normalize_theme_name(row.get("Theme")) or "Other Factors",
                    "Merged_Concept": clean_merged_concept(row.get("Merged_Concept", "")),
                    "Confidence": float(row.get("Confidence", 0.9)) if str(row.get("Confidence", "")).strip() else 0.9,
                    "Mapping_Source": str(row.get("Mapping_Source", "llm_id_mapping")),
                }
            )
            continue

        fuzzy = fuzzy_match_from_reference(original, reference_lookup)
        if fuzzy:
            filled_rows.append(
                {
                    "Original": original,
                    "Theme": fuzzy["Theme"],
                    "Merged_Concept": clean_merged_concept(fuzzy["Merged_Concept"]),
                    "Confidence": fuzzy["Confidence"],
                    "Mapping_Source": fuzzy["Mapping_Source"],
                }
            )
            continue

        theme_candidate, rule_conf, rule_source = classify_theme_rule_based(original)
        if rule_conf < RULE_THEME_MIN_SCORE:
            theme_candidate = "Other Factors"

        filled_rows.append(
            {
                "Original": original,
                "Theme": theme_candidate,
                "Merged_Concept": clean_merged_concept(derive_concept_from_text(original, theme_candidate)),
                "Confidence": round(rule_conf, 4),
                "Mapping_Source": rule_source,
            }
        )

    filled_df = pd.DataFrame(filled_rows)
    filled_df = filled_df.drop_duplicates(subset=["Original"], keep="first")
    return filled_df

def process_file(input_csv, output_csv, type_label):
    if not os.path.exists(input_csv):
        print(f"File {input_csv} not found. Skipping.")
        return

    df_unique = pd.read_csv(input_csv)
    unique_list = df_unique['text'].dropna().unique().tolist()

    already_processed = set()
    reference_lookup = {}
    if os.path.exists(output_csv):
        try:
            existing_output = pd.read_csv(output_csv)
            if 'Original' in existing_output.columns:
                existing_clean = postprocess_mapping_batch(existing_output)
                already_processed = set(existing_output['Original'].dropna().astype(str).tolist())
                reference_lookup = build_reference_lookup(existing_clean)
                print(f"♻️ Resume mode: found {len(already_processed)} already processed {type_label} rows in {output_csv} ({len(reference_lookup)} in lookup)")
        except Exception as read_error:
            print(f"⚠️ Could not read existing output for resume: {read_error}")

    pending_list = [item for item in unique_list if str(item) not in already_processed]
    if not pending_list:
        print(f"✅ Nothing pending for {type_label}. {output_csv} is already up to date.")
        return

    batch_size = TAGGER_BATCH_SIZE

    total_batches = (len(pending_list) + batch_size - 1) // batch_size
    provider_name = llm_provider.describe() if llm_provider else "unknown-llm"
    print(f"🔍 Analyzing {len(pending_list)} pending Unique {type_label}s via {provider_name}...")
    print(f"   Total Batches: {total_batches} | Batch Size: {batch_size}")

    for i in tqdm(range(0, len(pending_list), batch_size)):
        current_batch = (i // batch_size) + 1
        print(f"   ⏳ Processing Batch {current_batch}/{total_batches}...")

        current_items = pending_list[i : i + batch_size]
        mapped_df = pd.DataFrame()

        for attempt in range(1, MAX_RETRIES + 1):
            try:
                mapped_df = get_ai_mapping(current_items, type_label)
                break
            except Exception as batch_error:
                retryable = is_retryable_error(batch_error)
                should_retry = retryable and attempt < MAX_RETRIES

                print(f"      ⚠️ Batch {current_batch} attempt {attempt}/{MAX_RETRIES} failed: {batch_error}")
                if should_retry:
                    delay = min(MAX_RETRY_SECONDS, BASE_RETRY_SECONDS * (2 ** (attempt - 1)))
                    delay += random.uniform(0, 0.5)
                    print(f"      🔁 Retrying in {delay:.1f}s...")
                    time.sleep(delay)
                    continue

                print(f"      ❌ Batch {current_batch} failed after {attempt} attempt(s).")
                mapped_df = pd.DataFrame()
                break

        mapped_df = enforce_batch_coverage(current_items, mapped_df, reference_lookup)

        if not mapped_df.empty:
            save_progress(output_csv, mapped_df)
            for _, row in mapped_df.iterrows():
                orig = str(row.get("Original", "")).strip()
                norm = normalize_for_match(orig)
                theme = normalize_theme_name(row.get("Theme"))
                concept = row.get("Merged_Concept", "")
                conf = row.get("Confidence", 0.9)
                if norm and theme:
                    reference_lookup[norm] = {
                        "Theme": theme,
                        "Merged_Concept": clean_merged_concept(str(concept)),
                        "Confidence": float(conf) if str(conf).strip() else 0.9,
                        "Mapping_Source": "exact_reference",
                    }
            print(f"      ✅ Batch {current_batch} done. Saved {len(mapped_df)} rows to {output_csv}.")
        else:
            print(f"      ⚠️ Batch {current_batch} produced no usable rows.")

        time.sleep(0.5) 

    if os.path.exists(output_csv):
        final_rows = len(pd.read_csv(output_csv))
        print(f"✅ Mapping successfully saved to {output_csv} | Total rows now: {final_rows}")

if __name__ == "__main__":
    # Ensure these files exist from Phase 1
    process_file('unique_challenges.csv', 'challenge_mapping.csv', 'Challenge')
    process_file('unique_solutions.csv', 'solution_mapping.csv', 'Solution')
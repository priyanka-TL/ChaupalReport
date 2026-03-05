import pandas as pd
import os
import time
import io
import random
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
TAGGER_BATCH_SIZE = int(os.getenv("TAGGER_BATCH_SIZE", "25"))
TAGGER_MAX_TOKENS = int(os.getenv("TAGGER_MAX_TOKENS", "6000"))
TAGGER_THINKING_BUDGET = int(os.getenv("TAGGER_THINKING_BUDGET", "256"))


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

    if os.path.exists(output_csv):
        existing_df = pd.read_csv(output_csv)
        merged_df = pd.concat([existing_df, batch_df], ignore_index=True)
        merged_df = merged_df.drop_duplicates(subset=["Original"], keep="last")
    else:
        merged_df = batch_df.copy()

    merged_df.to_csv(output_csv, index=False)


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
    df_batch['Theme'] = df_batch['Theme'].astype(str).str.strip()
    df_batch['Merged_Concept'] = df_batch['Merged_Concept'].apply(clean_merged_concept)

    # Remove accidental duplicate rows and keep one mapping per Original
    df_batch = df_batch.drop_duplicates(subset=['Original'], keep='last')
    return df_batch

def get_ai_mapping(text_batch, type_label):
    if not llm_provider:
        raise RuntimeError("LLM provider is not configured")

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
    OUTPUT: Return ONLY a CSV-style format with three columns: Original|Theme|Merged_Concept
    Use the | character as the delimiter. Do not include headers, preamble, or markdown backticks.
    
    DATA:
    {text_batch}"""

    raw_output = llm_provider.generate_text(
        prompt_content,
        max_tokens=TAGGER_MAX_TOKENS,
        temperature=0,
        thinking_budget=TAGGER_THINKING_BUDGET,
    )

    raw_output = raw_output.replace('```csv', '').replace('```', '').strip()
    df_batch = pd.read_csv(io.StringIO(raw_output), sep='|', names=['Original', 'Theme', 'Merged_Concept'], header=None)
    return postprocess_mapping_batch(df_batch)

def process_file(input_csv, output_csv, type_label):
    if not os.path.exists(input_csv):
        print(f"File {input_csv} not found. Skipping.")
        return

    df_unique = pd.read_csv(input_csv)
    unique_list = df_unique['text'].dropna().unique().tolist()

    already_processed = set()
    if os.path.exists(output_csv):
        try:
            existing_output = pd.read_csv(output_csv)
            if 'Original' in existing_output.columns:
                already_processed = set(existing_output['Original'].dropna().astype(str).tolist())
                print(f"♻️ Resume mode: found {len(already_processed)} already processed {type_label} rows in {output_csv}")
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
        batch = "\n".join(current_items)
        mapped_df = pd.DataFrame()

        for attempt in range(1, MAX_RETRIES + 1):
            try:
                mapped_df = get_ai_mapping(batch, type_label)
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

        if not mapped_df.empty:
            save_progress(output_csv, mapped_df)
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
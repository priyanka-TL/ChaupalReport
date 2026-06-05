import pandas as pd
import os
import re
import time
import io
import random
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
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
1. Poverty and Economic Barriers:
   IN SCOPE: child labour, financial hardship, no money for fees, poverty, family income,
             working children, economic necessity, daily wage, hunger, seasonal migration for work,
             harvest work pulling children from school, sibling labour, debt, unable to afford,
             poor economic condition, family struggles financially, no income
   OUT OF SCOPE: academic failure, distance, teacher quality

2. Legal Document-linked Barriers:
   IN SCOPE: Aadhaar, birth certificate, caste certificate, no ID, no documents,
             enrollment blocked due to paperwork, scheme eligibility, document correction,
             registration, transfer certificate, no school leaving certificate
   OUT OF SCOPE: school infrastructure, teacher shortage

3. Child Marriage Issue:
   IN SCOPE: early marriage, girl married before 18, marriage preventing education,
             married off, dowry, girls taken out for marriage, marital responsibilities,
             betrothal, husband's family opposed to education, child bride
   OUT OF SCOPE: pregnancy unrelated to marriage, domestic violence without marriage context

4. Distance and Accessibility Issues:
   IN SCOPE: school far from home, no transport, bad roads, river crossing, seasonal floods,
             road condition, long walk, no bus, remote village, geography barrier,
             no connectivity, broken bridge, rain blocking path, travel difficulty
   OUT OF SCOPE: school infrastructure inside the building

5. Parental Attitudes & Socio-Cultural:
   IN SCOPE: parents unwilling, mindset against girls' education, traditional beliefs,
             domestic work priority, gender discrimination, caste prejudice, social norms,
             community pressure, attitudes, household chores, child not sent to school
             without specific stated reason, general dropout with parental decision,
             parents do not value education, cultural beliefs, lack of awareness about education
   OUT OF SCOPE: legal documents, economic hardship, safety

6. School Infrastructure & Facility:
   IN SCOPE: no toilet, no drinking water, mid-day meal issues, missing classroom,
             broken building, no boundary wall, no furniture, lack of books/uniform,
             scholarship delays, school not functional, no electricity, no playground,
             scheme not implemented, government scheme not reaching students
   OUT OF SCOPE: teacher quality, distance

7. Teacher Capacity & Quality:
   IN SCOPE: teacher absent, teacher shortage, untrained teacher, teacher attitude,
             no subject teacher, poor teaching quality, teacher bias, irregular teacher,
             ghost teacher, teacher harassment, teacher favouritism, teacher not coming
   OUT OF SCOPE: school building, government schemes

8. Safety Issues:
   IN SCOPE: harassment on the way, unsafe route, eve-teasing, stray dogs/animals,
             molestation, fear of safety, crime near school, unsafe environment,
             threats, danger, girls fear going out, sexual harassment
   OUT OF SCOPE: distance alone without safety mention

9. Substance Abuse & Addiction:
   IN SCOPE: alcohol addiction, drug use, gambling, mobile phone addiction, betting,
             substance abuse by parents or children, addiction-related school dropout,
             drunk parent, father's addiction, tobacco, intoxication
   OUT OF SCOPE: poverty alone

10. Other Factors (LAST RESORT — target < 10%):
    ONLY for texts that are completely off-topic, non-educational, or genuinely impossible
    to place in themes 1–9 after careful consideration.
    DO NOT use for vague statements — classify by the most likely theme.
    "Children not going to school" without reason → use Parental Attitudes (theme 5).
    "General awareness" → use Parental Attitudes (theme 5).
    "Migration" or "seasonal migration" → use Poverty (theme 1).
    NEVER use as default fallback.
"""

VALID_THEMES = frozenset({
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
})

NAME_PATTERN_TAGGER = re.compile(
    r'\b(kumari|devi|singh|khatoon|parveen|siddiqui|bano|begum|rani|meera|'
    r'priya|kavita|sunita|rita|nisha|neha|sonam|shobha|rinku|aafia|tabish|asif|'
    r'jyoti|suman|nagma|gudiya|shivani|priyanshu|karuna|chamuni|vidya|rekha|'
    r'savita|mamta|geeta|seema|anita|pooja|renu|mala|lata|usha|saroj)\b',
    re.IGNORECASE
)

JUNK_CONCEPTS_TAGGER = frozenset({
    "uncategorized", "vague", "incomplete statement", "n/a", "none", "",
    "vague or unspecified educational issue", "uncategorized/irrelevant data",
    "unspecified issue", "unclear", "not specified",
})

MAX_RETRIES = int(os.getenv("LLM_MAX_RETRIES", "5"))
BASE_RETRY_SECONDS = float(os.getenv("LLM_RETRY_BASE_SECONDS", "2"))
MAX_RETRY_SECONDS = float(os.getenv("LLM_RETRY_MAX_SECONDS", "45"))
TAGGER_BATCH_SIZE = int(os.getenv("TAGGER_BATCH_SIZE", "25"))
TAGGER_MAX_TOKENS = int(os.getenv("TAGGER_MAX_TOKENS", "6000"))
TAGGER_THINKING_BUDGET = int(os.getenv("TAGGER_THINKING_BUDGET", "256"))
TAGGER_WORKERS = int(os.getenv("TAGGER_WORKERS", "4"))          # parallel API workers
TAGGER_FLUSH_INTERVAL = int(os.getenv("TAGGER_FLUSH_INTERVAL", "80"))  # flush to disk every N batches


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

    return cleaned[0].upper() + cleaned[1:] if len(cleaned) > 1 else cleaned.upper()


def _validate_row(theme, concept):
    """Return True if the row passes schema validation."""
    if theme not in VALID_THEMES:
        return False
    concept_lower = concept.strip().lower()
    if concept_lower in JUNK_CONCEPTS_TAGGER or not concept_lower:
        return False
    if NAME_PATTERN_TAGGER.search(concept):
        return False
    if len(concept.split()) < 3:
        return False
    return True


def postprocess_mapping_batch(df_batch):
    """Apply deterministic cleanup and schema validation so model output stays canonical."""
    if df_batch.empty:
        return df_batch

    df_batch['Original'] = df_batch['Original'].astype(str).str.strip()
    df_batch['Theme'] = df_batch['Theme'].astype(str).str.strip()
    df_batch['Merged_Concept'] = df_batch['Merged_Concept'].apply(clean_merged_concept)

    # Strip personal names from Merged_Concept
    df_batch['Merged_Concept'] = df_batch['Merged_Concept'].apply(
        lambda x: NAME_PATTERN_TAGGER.sub('[Educational Issue]', x) if NAME_PATTERN_TAGGER.search(x) else x
    )

    # Validate theme — fix common near-misses
    theme_fixes = {
        "Parental Attitudes & Socio-Cultural Factors": "Parental Attitudes & Socio-Cultural",
        "Parental Attitude & Socio-Cultural": "Parental Attitudes & Socio-Cultural",
        "School Infrastructure & Facilities": "School Infrastructure & Facility",
        "Poverty & Economic Barriers": "Poverty and Economic Barriers",
    }
    df_batch['Theme'] = df_batch['Theme'].replace(theme_fixes)

    # Filter rows that fail schema validation
    valid_mask = df_batch.apply(
        lambda r: _validate_row(r['Theme'], r['Merged_Concept']), axis=1
    )
    invalid_count = (~valid_mask).sum()
    if invalid_count > 0:
        print(f"      🔎 Filtered {invalid_count} invalid rows (bad theme/PII/junk concept).")
    df_batch = df_batch[valid_mask].copy()

    # Remove accidental duplicate rows and keep one mapping per Original
    df_batch = df_batch.drop_duplicates(subset=['Original'], keep='last')
    return df_batch

def get_ai_mapping(text_batch, type_label, canonical_registry=None):
    """Map a batch of raw texts to (Theme, Merged_Concept).

    canonical_registry: dict {concept_label: occurrence_count} accumulated from
    all prior batches. The top-50 labels are injected into the prompt so the AI
    reuses existing labels rather than inventing synonymous new ones.
    This is the cross-batch normalisation mechanism that prevents label drift.
    """
    if not llm_provider:
        raise RuntimeError("LLM provider is not configured")

    # Build canonical registry context block — top-50 most-used labels from prior batches
    registry_block = ""
    if canonical_registry:
        top_canonicals = sorted(canonical_registry.items(), key=lambda x: -x[1])[:50]
        registry_lines = "\n".join(f"  - {label}" for label, _ in top_canonicals)
        registry_block = f"""
CANONICAL LABEL REGISTRY (from previous batches — REUSE these exact labels when applicable):
{registry_lines}

CRITICAL: If a text in this batch expresses the same core issue as any label above,
use that EXACT label as the Merged_Concept. Do NOT invent a synonym.
For example: if "Poverty Preventing Education" is in the registry, do NOT write
"Financial Hardship Preventing Education" or "Economic Barriers To Schooling".
"""

    # Solution-specific instruction to prevent problem-framing in solution concepts
    solution_framing_block = ""
    if type_label.lower() == "solution":
        solution_framing_block = """
SOLUTION DATASET RULES (apply in addition to all rules above):
- This dataset contains SOLUTIONS and COMMITMENTS made by participants, not problems.
- Even if a participant described a barrier, identify the PROPOSED ACTION or COMMITMENT.
- Merged_Concept must be action-framed: "Parental Commitment To Education", not "Poverty Preventing Education".
- NEVER use "X Preventing Education" — that is challenge framing. Use "Overcoming X" or "Commitment To X".
- Negative-start phrases like "Lack of" / "No " / "Poor " are valid ONLY when the solution is to address that lack (e.g., "Ensuring Legal Documents").
"""

    prompt_content = f"""Act as an expert Social Data Analyst for a Bihar education study. Use these THEMES:
{THEME_KNOWLEDGE_BASE}
{registry_block}
{solution_framing_block}
MANDATORY CLASSIFICATION RULES:
- "Children not going/attending/coming to school" → classify by the REASON stated or implied:
  If reason is migration/harvest/labour → Poverty and Economic Barriers
  If fear/harassment → Safety Issues
  If marriage → Child Marriage Issue
  If no document → Legal Document-linked Barriers
  If distance → Distance and Accessibility Issues
  If parental decision/no specific reason → Parental Attitudes & Socio-Cultural
- THEME PRIORITY RULES (override keyword matching when conflict occurs):
  * ANY text mentioning Aadhaar / birth certificate / caste certificate / transfer certificate / documentation → Legal Document-linked Barriers (NEVER Poverty)
  * ANY text mentioning child marriage / early marriage / girl married / betrothal → Child Marriage Issue (NEVER Other Factors)
  * ANY text mentioning unsafe route / harassment on way / eve-teasing / molestation → Safety Issues
  * ANY text mentioning alcohol / drugs / substance abuse / gambling / mobile addiction → Substance Abuse & Addiction
  * Poverty/financial hardship text that ALSO mentions documents → Legal Document-linked Barriers wins
- NEVER classify as "Other Factors" unless the text is completely non-educational and cannot fit themes 1–9.
- NEVER use "Uncategorized", "Vague", "Incomplete", "N/A" or any junk as Merged_Concept.
- NEVER include person names (Kumari, Devi, Singh, Khatoon, Parveen, Siddiqui, Bano, Begum) in Merged_Concept. Extract only the educational issue.
- "General awareness" → Parental Attitudes & Socio-Cultural
- "Migration" or "seasonal migration" → Poverty and Economic Barriers
- Merged_Concept: 3–8 words, title-case, no trailing punctuation.
- NOISE: If a text is a garbled translation (e.g., food metaphors, nonsensical sentences) with zero educational meaning → Theme: "Other Factors", Merged_Concept: "Non-Educational Irrelevant Statement".

SEMANTIC DEDUPLICATION PROTOCOL (MANDATORY):
You must merge semantically similar items into a single "Merged_Concept".

Rules:
1. Group all variants expressing the same core issue.
2. Select ONE canonical phrase and reuse it for all equivalent variants in this batch.
   If the canonical registry above has a matching label, use that exact label.
2A. Before writing output, internally create a canonical dictionary for this batch.
2B. Use only those dictionary labels in final output (no one-off labels for similar meaning).
3. 'Merged_Concept' MUST follow canonical naming format:
    - concise noun phrase (3-8 words)
    - no ending punctuation
    - avoid sentence-style wording
    - avoid district/person-specific details
    - stable wording across similar records
4. Examples of MANDATORY merging (wording varies → ONE label):
   - "Due to poverty" = "Poor financial condition" = "Lack of money" = "Child labour" (poverty-linked) → "Poverty Preventing Education"
   - "No Aadhaar card" = "Lack of Aadhaar" = "Aadhaar not made" = "Aadhaar card problem" → "Aadhaar Card Issues"
   - "School is far" = "School is very far" = "Distance of school" = "No nearby school" → "School Distance And Accessibility Issues"
   - "Early marriage" = "Child marriage" = "Marriage stopping education" = "Girl married before 18" → "Child Marriage Preventing Education"
   - "Household chores" = "Domestic work" = "Girls doing home work" = "House work priority" → "Domestic Work Priority"
   - "Parents are not interested in sending children to school" = "Lack of parental commitment towards education"
     = "Parents don't care about children's studies" = "Parents not bothered about education"
     = "Parental apathy towards education" = "No parental support for education" → "Lack of Parental Commitment To Education"
   - "Parents do not value education" = "Parents unaware of education importance"
     = "Lack of parental education awareness" = "Parents not educated themselves" → "Lack of Parental Education Awareness"
   - "Teacher not coming" = "Teacher absent regularly" = "Teacher absenteeism" = "No teacher in school" → "Teacher Absence And Absenteeism"
   - "Parental commitment to education" = "Parents will send children to school" = "Parents committed to schooling"
     = "Parents pledged to educate children" → "Parental Commitment To Education"
   - "Community awareness of education" = "Promoting education value" = "Raising education awareness"
     = "Making people understand education importance" → "Community Education Awareness"
5. CRITICAL: Assign EXACTLY ONE theme. If multiple apply, choose the most dominant one. No '+' or 'and' between themes.

TASK: Categorize these unique {type_label} statements.
SELF-CHECK (MANDATORY, still same single call):
- Re-scan your own output and ensure semantically equivalent rows use exactly identical Merged_Concept text.
- If two labels differ only by wording (e.g., arranging/providing/facilitating same action), unify them.
- Verify NO row uses "Other Factors" unless truly impossible to classify elsewhere.
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

def _cluster_sort_texts(texts, batch_size=25):
    """Sort texts by TF-IDF cluster so semantically similar texts land in the same batch.

    When similar texts are batched together, the AI's intra-batch deduplication
    (SELF-CHECK rule) merges them → fewer unique canonical labels created overall.
    Falls back to original order if sklearn is unavailable or dataset is tiny.
    """
    if len(texts) <= batch_size * 2:
        return list(texts)
    try:
        from sklearn.feature_extraction.text import TfidfVectorizer
        from sklearn.cluster import MiniBatchKMeans

        n_clusters = max(len(texts) // batch_size, 2)
        n_clusters = min(n_clusters, len(texts) - 1)

        # Char n-grams catch morphological variants well for Hindi-translated English
        vec = TfidfVectorizer(
            analyzer='char_wb', ngram_range=(3, 5),
            sublinear_tf=True, min_df=1, max_features=8000
        )
        matrix = vec.fit_transform(texts)
        kmeans = MiniBatchKMeans(
            n_clusters=n_clusters, random_state=42,
            batch_size=min(1000, len(texts)), n_init=3, max_iter=30
        )
        cluster_ids = kmeans.fit_predict(matrix)

        # Secondary sort within each cluster by text length (shorter = better cluster seed)
        paired = sorted(zip(cluster_ids, [len(t) for t in texts], texts))
        sorted_texts = [t for _, _, t in paired]
        print(f"   🗂️  Pre-clustered {len(texts)} texts into ~{n_clusters} groups for batch alignment.")
        return sorted_texts
    except Exception as e:
        print(f"   ⚠️  Pre-clustering unavailable ({e}), using original order.")
        return list(texts)


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
                print(f"♻️ Resume mode: {len(already_processed)} already processed {type_label} rows in {output_csv}")
        except Exception as read_error:
            print(f"⚠️ Could not read existing output for resume: {read_error}")

    pending_list = [item for item in unique_list if str(item) not in already_processed]
    if not pending_list:
        print(f"✅ Nothing pending for {type_label}. {output_csv} is already up to date.")
        return

    # Pre-cluster similar texts → similar texts share batches → intra-batch AI deduplication fires
    pending_list = _cluster_sort_texts(pending_list, batch_size=TAGGER_BATCH_SIZE)

    batch_size = TAGGER_BATCH_SIZE
    batches = [pending_list[i:i + batch_size] for i in range(0, len(pending_list), batch_size)]
    total_batches = len(batches)

    provider_name = llm_provider.describe() if llm_provider else "unknown-llm"
    print(f"🔍 Analyzing {len(pending_list)} pending unique {type_label}s via {provider_name}...")
    print(f"   Batches: {total_batches} | Batch size: {batch_size} | Workers: {TAGGER_WORKERS}")

    # Seed canonical registry from prior run (resume-mode and cross-run consistency)
    canonical_registry: dict = {}
    if os.path.exists(output_csv):
        try:
            existing = pd.read_csv(output_csv)
            if 'Merged_Concept' in existing.columns:
                for concept, count in existing['Merged_Concept'].value_counts().items():
                    canonical_registry[str(concept)] = int(count)
        except Exception:
            pass

    # Snapshot the initial registry — all parallel workers get the same starting context.
    # Pre-clustering ensures similar texts are in the same batch so the AI's own SELF-CHECK
    # handles intra-batch deduplication; the registry handles cross-run consistency.
    initial_registry = dict(canonical_registry)

    def _run_one_batch(args):
        """Worker: process one batch (runs in thread pool, no shared state written)."""
        batch_idx, batch_texts = args
        batch_str = "\n".join(batch_texts)
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                return batch_idx, get_ai_mapping(batch_str, type_label,
                                                  canonical_registry=initial_registry)
            except Exception as err:
                if is_retryable_error(err) and attempt < MAX_RETRIES:
                    delay = min(MAX_RETRY_SECONDS, BASE_RETRY_SECONDS * (2 ** (attempt - 1)))
                    delay += random.uniform(0, 0.3)
                    print(f"      ⚠️ Batch {batch_idx+1} retry {attempt}: {err} — sleeping {delay:.1f}s")
                    time.sleep(delay)
                    continue
                print(f"      ❌ Batch {batch_idx+1} failed after {attempt} attempt(s): {err}")
                return batch_idx, pd.DataFrame()
        return batch_idx, pd.DataFrame()

    # ── PARALLEL EXECUTION ────────────────────────────────────────────────────
    pending_flush: list = []   # in-memory accumulator; flushed every TAGGER_FLUSH_INTERVAL
    completed = 0

    _t_start = time.time()
    with ThreadPoolExecutor(max_workers=TAGGER_WORKERS) as executor:
        future_map = {
            executor.submit(_run_one_batch, (i, batch)): i
            for i, batch in enumerate(batches)
        }
        for future in tqdm(as_completed(future_map), total=total_batches, desc=f"  {type_label}"):
            batch_idx, mapped_df = future.result()
            completed += 1

            if not mapped_df.empty:
                for concept, count in mapped_df['Merged_Concept'].value_counts().items():
                    canonical_registry[str(concept)] = (
                        canonical_registry.get(str(concept), 0) + int(count)
                    )
                pending_flush.append(mapped_df)
            else:
                print(f"      ⚠️ Batch {batch_idx+1} produced no usable rows.", flush=True)

            # Periodic disk flush to protect against crashes on large runs
            if len(pending_flush) >= TAGGER_FLUSH_INTERVAL:
                combined = pd.concat(pending_flush, ignore_index=True)
                save_progress(output_csv, combined)
                pending_flush.clear()
                elapsed = time.time() - _t_start
                rps = completed / elapsed if elapsed > 0 else 0
                print(f"   💾 Flush at batch {batch_idx+1}: {completed}/{total_batches} done "
                      f"| {rps:.2f} batches/s | Registry: {len(canonical_registry)} labels", flush=True)

    # Final flush
    if pending_flush:
        combined = pd.concat(pending_flush, ignore_index=True)
        save_progress(output_csv, combined)

    elapsed_total = time.time() - _t_start
    if os.path.exists(output_csv):
        final_df = pd.read_csv(output_csv)
        final_rows = len(final_df)
        final_concepts = final_df['Merged_Concept'].nunique() if 'Merged_Concept' in final_df.columns else 0
        print(f"✅ {type_label} mapping saved → {output_csv} | "
              f"{final_rows:,} rows | {final_concepts:,} unique concepts | "
              f"{elapsed_total/60:.1f} min")

if __name__ == "__main__":
    # Ensure these files exist from Phase 1
    process_file('unique_challenges.csv', 'challenge_mapping.csv', 'Challenge')
    process_file('unique_solutions.csv', 'solution_mapping.csv', 'Solution')
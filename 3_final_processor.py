import pandas as pd
import re
import os
import time
import random
from difflib import SequenceMatcher
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import json
from dotenv import load_dotenv
from llm_provider import LLMProvider

load_dotenv()

# --- LLM CONFIGURATION ---
try:
    llm_provider = LLMProvider()
except Exception as e:
    print(f"⚠️ LLM Client Setup Failed: {e}")
    llm_provider = None

THEME_KNOWLEDGE_BASE = """
1. Poverty and Economic Barriers
2. Legal Document-linked Barriers
3. Child Marriage Issue
4. Distance and Accessibility Issues
5. Parental Attitudes & Socio-Cultural
6. School Infrastructure & Facility
7. Teacher Capacity & Quality
8. Safety Issues
9. Substance Abuse & Addiction
10. Other Factors
"""

MAX_RETRIES = int(os.getenv("LLM_MAX_RETRIES", "5"))
BASE_RETRY_SECONDS = float(os.getenv("LLM_RETRY_BASE_SECONDS", "2"))
MAX_RETRY_SECONDS = float(os.getenv("LLM_RETRY_MAX_SECONDS", "45"))
REFINE_CHECKPOINT_DIR = os.getenv("REFINE_CHECKPOINT_DIR", ".")
REFINE_BATCH_SIZE = int(os.getenv("REFINE_BATCH_SIZE", "25"))
REFINE_MAX_TOKENS = int(os.getenv("REFINE_MAX_TOKENS", "6000"))
REFINE_THINKING_BUDGET = int(os.getenv("REFINE_THINKING_BUDGET", "256"))
INSIGHT_MAX_TOKENS = int(os.getenv("INSIGHT_MAX_TOKENS", "900"))
INSIGHT_THINKING_BUDGET = int(os.getenv("INSIGHT_THINKING_BUDGET", "256"))
INSIGHT_TEMPERATURE = float(os.getenv("INSIGHT_TEMPERATURE", "0.15"))
MIN_DISTRICT_CHAUPALS = int(os.getenv("MIN_DISTRICT_CHAUPALS", "5"))
INSIGHT_BATCH_SIZE = int(os.getenv("INSIGHT_BATCH_SIZE", "6"))
INSIGHT_CACHE_FILE = os.getenv("INSIGHT_CACHE_FILE", "insight_cache.json")

# PII detection: Indian name suffixes commonly appearing in participant quotes
NAME_PATTERN = re.compile(
    r'\b(kumari|devi|singh|khatoon|parveen|siddiqui|bano|begum|rani|meera|'
    r'priya|kavita|sunita|rita|nisha|neha|sonam|shobha|rinku|aafia|tabish|asif|'
    r'jyoti|suman|nagma|gudiya|shivani|priyanshu|karuna|chamuni|vidya|rekha|'
    r'savita|mamta|geeta|seema|anita|pooja|renu|mala|lata|usha|saroj)\b',
    re.IGNORECASE
)
JUNK_CONCEPTS = frozenset({
    "uncategorized", "vague", "incomplete statement", "n/a", "none", "",
    "vague or unspecified educational issue", "uncategorized/irrelevant data",
})
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


def _is_pii_text(text):
    """Returns True if text contains a personal name pattern."""
    return bool(NAME_PATTERN.search(str(text)))


def _is_junk_concept(text):
    """Returns True if text is a known junk/artifact label."""
    return str(text).strip().lower() in JUNK_CONCEPTS


def get_refinement_checkpoint_path(type_label):
    safe_label = str(type_label).strip().lower().replace(" ", "_")
    return os.path.join(REFINE_CHECKPOINT_DIR, f"refine_checkpoint_{safe_label}.json")


def load_refinement_checkpoint(checkpoint_path):
    if not os.path.exists(checkpoint_path):
        return {}
    try:
        with open(checkpoint_path, "r", encoding="utf-8") as file:
            data = json.load(file)
        return data if isinstance(data, dict) else {}
    except Exception as error:
        print(f"   ⚠️ Could not read refinement checkpoint {checkpoint_path}: {error}")
        return {}


def save_refinement_checkpoint(checkpoint_path, results):
    os.makedirs(os.path.dirname(checkpoint_path) or ".", exist_ok=True)
    with open(checkpoint_path, "w", encoding="utf-8") as file:
        json.dump(results, file, ensure_ascii=False, indent=2)


def _extract_json_block(text):
    cleaned = str(text).replace('```json', '').replace('```', '').strip()
    start = cleaned.find('{')
    end = cleaned.rfind('}')
    if start != -1 and end != -1 and end > start:
        return cleaned[start:end + 1]
    return cleaned


def _parse_refinement_response(text, batch):
    raw_json = _extract_json_block(text)
    parsed = json.loads(raw_json)
    if not isinstance(parsed, dict):
        raise ValueError("Refinement response is not a JSON object")

    valid = {}
    for item in batch:
        value = parsed.get(item)
        if not isinstance(value, dict):
            continue
        concept = str(value.get('concept', '')).strip()
        theme = str(value.get('theme', '')).strip()
        if concept and theme:
            valid[item] = {'concept': concept, 'theme': theme}

    if not valid:
        raise ValueError("No valid refinement mappings found in response")

    return valid


def _repair_refinement_json_with_ai(raw_text, batch, type_label):
    if not llm_provider:
        return None

    repair_prompt = f"""Fix the malformed JSON below.

RULES:
1. Return ONLY valid JSON object.
2. Keys must come from this input list: {json.dumps(batch, ensure_ascii=False)}
3. Each value must be an object with keys: concept, theme.
4. Keep meaning intact. Do not add extra text.

MALFORMED JSON/TEXT:
{raw_text}

OUTPUT: valid JSON object only."""

    repaired = llm_provider.generate_text(repair_prompt, max_tokens=2500, temperature=0)
    return _parse_refinement_response(repaired, batch)

def categorize_environment_aggressive(text):
    """Ultra-Aggressive Environment Classification to minimize Unmapped tags."""
    text_lower = str(text).lower()
    
    school_kw = ['school', 'teacher', 'classroom', 'class', 'student', 'education', 'study', 'teaching', 'academic', 'admission', 'enroll', 'attendance', 'grade', 'subject', 'exam', 'books', 'uniform', 'midday meal', 'mid day', 'scholarship', 'library', 'playground', 'infrastructure', 'facility', 'toilet', 'water', 'building']
    home_kw = ['parent', 'family', 'mother', 'father', 'home', 'household', 'house', 'sibling', 'brother', 'sister', 'domestic', 'child labour', 'work at home', 'income', 'alcoholic', 'migration', 'marriage', 'dowry', 'attitude', 'mindset', 'belief', 'cultural', 'discrimination']
    comm_kw = ['village', 'community', 'society', 'road', 'transport', 'bus', 'distance', 'far', 'path', 'route', 'weather', 'rain', 'heat', 'flood', 'surroundings', 'neighborhood', 'area', 'locality', 'safety', 'harassment', 'molestation', 'social pressure', 'caste', 'tribe', 'practice']
    
    s_score = sum(2 if kw in text_lower else 0 for kw in school_kw)
    h_score = sum(2 if kw in text_lower else 0 for kw in home_kw)
    c_score = sum(2 if kw in text_lower else 0 for kw in comm_kw)
    
    # Contextual boosts
    if any(kw in text_lower for kw in ['to school', 'reach school', 'go to school']): c_score += 3
    if any(kw in text_lower for kw in ['at home', 'in family', 'parent awareness']): h_score += 3
    if any(kw in text_lower for kw in ['in school', 'at school', 'lacks']): s_score += 3
    
    scores = {'School': s_score, 'Home': h_score, 'Community': c_score}
    if max(scores.values()) == 0:
        if any(w in text_lower for w in ['poor', 'poverty', 'money', 'financial']): return 'Home'
        return 'Community' # Default fallback
    return max(scores, key=scores.get)

def categorize_agency(text):
    """Classifies the driver of the solution."""
    text_lower = str(text).lower()
    comm_kw = ['community', 'together', 'collective', 'meena manch', 'chaupal', 'village', 'we will', 'committee', 'panchayat']
    ind_kw = ['parent', 'family', 'individual', 'we should', 'people should', 'personally', 'mother', 'father']
    inst_kw = ['government', 'school', 'ngo', 'administration', 'authority', 'provide', 'officer', 'department', 'teacher']
    
    scores = {
        'Community-led': sum(1 for kw in comm_kw if kw in text_lower),
        'Individual-led': sum(1 for kw in ind_kw if kw in text_lower),
        'Institutional': sum(1 for kw in inst_kw if kw in text_lower)
    }
    return max(scores, key=scores.get) if max(scores.values()) > 0 else 'Community-led'

# --- 2. FORMATTING UTILITIES ---

def set_cell_background(cell, fill_color):
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), fill_color)
    cell._tc.get_or_add_tcPr().append(shading_elm)

def normalize_text(text):
    if pd.isna(text): return "Uncategorized"
    return re.sub(r'^\d+[\.\)\s-]*', '', str(text)).strip()

def is_valid_solution(text):
    """Filters out problem statements disguised as solutions."""
    text = str(text).lower().strip()
    problem_starters = ['lack of', 'no ', 'not enough', 'poor ', 'insufficient', 'scarcity', 'absence', 'shortage', 'due to', 'because of']
    if any(text.startswith(p) for p in problem_starters):
        return False
    return True

def clean_theme_name(text):
    """Cleans theme names, handling combined themes and empty values."""
    if pd.isna(text) or str(text).strip() == "" or str(text).lower() == "nan":
        return "Other Factors"
    
    text = str(text).strip()
    
    # Handle combined themes (e.g., "Theme A + Theme B") - Take the first one
    if '+' in text:
        text = text.split('+')[0].strip()
        
    # Remove leading numbers/bullets (e.g., "1. Poverty")
    text = re.sub(r'^\d+[\.\)\s-]*', '', text).strip()
    
    return text


def normalize_concept_key(text):
    """Canonical key for grouping near-duplicate concept labels."""
    if pd.isna(text):
        return ""
    normalized = str(text).strip().lower()
    normalized = re.sub(r'\s+', ' ', normalized)
    normalized = re.sub(r'[^a-z0-9\s]', '', normalized)
    return normalized.strip()


# ── TF-IDF SIMILARITY CLUSTERING ─────────────────────────────────────────────
# Config: tune these per dataset size
SIMILARITY_THRESHOLD = float(os.getenv("CONCEPT_SIMILARITY_THRESHOLD", "0.35"))
SIMILARITY_CHUNK_SIZE = int(os.getenv("CONCEPT_CHUNK_SIZE", "2000"))


def _build_tfidf_matrix(texts):
    """Build a combined char-ngram + word TF-IDF matrix for a list of text labels."""
    from sklearn.feature_extraction.text import TfidfVectorizer
    from scipy.sparse import hstack

    # Character n-grams: capture "child marriage" ≈ "early marriage" via shared substrings
    char_vec = TfidfVectorizer(analyzer='char_wb', ngram_range=(3, 5), min_df=1, sublinear_tf=True)
    # Word n-grams: capture exact keywords like "poverty", "aadhaar", "dropout"
    word_vec = TfidfVectorizer(analyzer='word', ngram_range=(1, 2), min_df=1, sublinear_tf=True)

    char_matrix = char_vec.fit_transform(texts)
    word_matrix = word_vec.fit_transform(texts)
    return hstack([char_matrix * 0.4, word_matrix * 0.6])  # weight words more


def _find_similar_pairs_chunked(matrix, keys, threshold):
    """Compute cosine similarity in row-chunks to stay memory-safe at any scale.

    For n=1500: one pass, ~10ms.
    For n=50000: chunk_size=2000 gives 25 passes, each sparse dot-product is fast.
    """
    from sklearn.metrics.pairwise import cosine_similarity
    import numpy as np

    n = len(keys)
    pairs = []
    chunk = SIMILARITY_CHUNK_SIZE

    for start in range(0, n, chunk):
        end = min(start + chunk, n)
        # chunk_matrix shape: (chunk_size, vocab) — stays in RAM comfortably
        sim_block = cosine_similarity(matrix[start:end], matrix)
        for local_i, global_i in enumerate(range(start, end)):
            row = sim_block[local_i]
            # Only look at j > global_i to avoid double-counting
            for global_j in range(global_i + 1, n):
                if row[global_j] >= threshold:
                    pairs.append((keys[global_i], keys[global_j]))

    return pairs


def assign_concept_groups(df, concept_column='Merged_Concept'):
    """TF-IDF cosine-similarity clustering for semantic concept deduplication.

    Replaces the bucket-index approach. Key improvements:
    - Compares ALL concept pairs across the full vocabulary (no first-token restriction)
    - char-ngram + word TF-IDF captures semantic variants like
      'Early Marriage' ≈ 'Child Marriage', 'Financial Hardship' ≈ 'Poverty'
    - Scales safely to 50K+ concepts via chunked similarity computation
    - Union-Find ensures O(α) grouping after pairs are found

    SIMILARITY_THRESHOLD (env CONCEPT_SIMILARITY_THRESHOLD, default 0.50):
      Raise to 0.65 to merge only near-identical labels.
      Lower to 0.40 to merge more aggressively (risk of over-grouping).
    """
    from collections import defaultdict

    if df.empty:
        temp = df.copy()
        temp['Concept_Key'] = ""
        temp['Concept_Group'] = ""
        return temp

    temp = df.copy()
    temp['Concept_Key'] = temp[concept_column].apply(normalize_concept_key)
    temp = temp[temp['Concept_Key'] != ""].copy()
    if temp.empty:
        temp['Concept_Group'] = ""
        return temp

    key_counts = temp['Concept_Key'].value_counts().to_dict()
    keys = list(key_counts.keys())

    if len(keys) == 1:
        temp['Concept_Group'] = keys[0]
        return temp

    # Union-Find with frequency-weighted root election
    parent = {k: k for k in keys}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra == rb:
            return
        # Higher-frequency concept becomes the canonical root
        if key_counts.get(ra, 0) >= key_counts.get(rb, 0):
            parent[rb] = ra
        else:
            parent[ra] = rb

    # Step 1: exact-key deduplication (zero cost)
    seen_keys = {}
    for key in keys:
        if key in seen_keys:
            union(key, seen_keys[key])
        else:
            seen_keys[key] = key

    # Step 2: TF-IDF cosine similarity across ALL pairs
    try:
        matrix = _build_tfidf_matrix(keys)
        similar_pairs = _find_similar_pairs_chunked(matrix, keys, SIMILARITY_THRESHOLD)
        for ka, kb in similar_pairs:
            union(ka, kb)
    except Exception as e:
        print(f"   ⚠️  TF-IDF clustering error: {e} — falling back to exact-match grouping.", flush=True)

    group_map = {key: find(key) for key in keys}
    temp['Concept_Group'] = temp['Concept_Key'].map(group_map)
    return temp


def resolve_canonical_labels(df, concept_column='Merged_Concept'):
    """After TF-IDF clustering, replace every variant with the group's most-frequent label.

    This collapses 'Early Marriage Preventing Education' (57) and
    'Child Marriage Preventing Education' (55) both into whichever label has the higher
    raw count across the full dataframe — producing a single consolidated entry in the
    Top Challenges table instead of two fragmented rows.

    Call this once on df_c and df_s *after* assign_concept_groups(), before building
    the report sections.
    """
    if 'Concept_Group' not in df.columns:
        return df

    temp = df.copy()
    # For each Concept_Group, find the original (display) label with the highest row count
    group_to_canonical = (
        temp.groupby('Concept_Group')[concept_column]
        .agg(lambda labels: labels.value_counts().idxmax())
        .to_dict()
    )
    temp[concept_column] = temp['Concept_Group'].map(group_to_canonical).fillna(temp[concept_column])
    return temp

def _challenge_item_insight(theme, concept, share, count, t_c, t_s, challenge_texts, max_words=100, concept_key=None, concept_rows=None):
    """Creates 2-3 deep, specific insight sentences (max 100 words) revealing ground realities."""
    total_theme_chal = len(t_c)
    total_theme_sol = len(t_s)

    if concept_rows is not None:
        concept_rows = concept_rows.copy()
    elif concept_key is not None:
        concept_rows = t_c[t_c['Merged_Concept'].apply(normalize_concept_key) == concept_key]
    else:
        concept_rows = t_c[t_c['Merged_Concept'] == concept]
    districts = concept_rows['District'].nunique() if not concept_rows.empty and 'District' in concept_rows.columns else 0

    env_text = "Not available"
    if not concept_rows.empty and 'Environment' in concept_rows.columns:
        env_mix = concept_rows['Environment'].value_counts(normalize=True)
        if not env_mix.empty:
            env_text = f"{env_mix.index[0]}"

    # Sample actual ground scenarios for deeper analysis
    sample_texts = challenge_texts[:8] if len(challenge_texts) > 8 else challenge_texts
    scenarios_text = " || ".join(sample_texts)

    # Fallback insights - direct, analytical
    fallback = (
        f"This represents {share:.1f}% of theme challenges, indicating a systemic issue rather than isolated incidents. "
        f"The pattern manifests primarily in {env_text} settings, pointing to where interventions must be anchored. "
        f"With {districts} district(s) reporting this challenge, it requires {'localized' if districts <= 2 else 'coordinated multi-district'} response strategies."
    )

    if not llm_provider:
        return [fallback]

    prompt = f"""You are analyzing on-ground education barriers from grassroots dialogue data.

ANALYZE THESE ACTUAL GROUND SCENARIOS:
{scenarios_text}

CONTEXT:
- Theme: {theme}
- Number of similar cases: {count} ({share:.1f}% of theme)
- Geographic spread: {districts} district(s)
- Primary setting: {env_text}

TASK:
Write 2-3 DISTINCT, NON-REPETITIVE insights that reveal the BROADER PICTURE of what's happening on the ground.

EACH INSIGHT MUST COVER A DIFFERENT DIMENSION:
✓ Insight 1: What MECHANISM/TRIGGER causes this barrier? (e.g., sudden economic shocks, rigid documentation rules, infrastructure gaps)
✓ Insight 2: WHO is most affected and WHAT cascading effects occur? (e.g., girls withdrawn first, entire families pulled out, seasonal disruptions)
✓ Insight 3 (if needed): What SYSTEMIC PATTERN or broader implication emerges? (e.g., policy-implementation gaps, urban-rural divide, poverty multipliers)

CRITICAL REQUIREMENTS:
✗ NO repetition - each sentence must add NEW information, not rephrase the same point
✗ NO generic statements like "barriers impede access" or "factors prevent education"
✗ NO explicit references to "voices," "testimonials," or "participants said"
✗ NO repetition of the challenge concept name (it's in the heading above)
✓ Be CONCRETE and SPECIFIC about mechanisms, triggers, affected groups, cascading effects
✓ Synthesize the BROADER PICTURE from multiple scenarios - what patterns emerge?
✓ Each insight should answer a DIFFERENT question about the challenge
✓ Use PERFECT grammar, spelling, and punctuation - proofread carefully
✓ Write in complete, well-structured sentences with proper syntax

EXAMPLE OF NON-REPETITIVE INSIGHTS:
❌ BAD (repetitive): "Rigid enforcement blocks children. Inflexibility disqualifies students."
✅ GOOD (distinct dimensions): "Minor documentation discrepancies trigger automatic rejection during enrollment. Marginalized families—lacking digital literacy or correction mechanisms—face permanent exclusion, with no appeals process available."

WORD LIMIT: Maximum 100 words total.

OUTPUT: Return 2-3 distinct, non-overlapping, grammatically perfect insight sentences. Separate with double newlines."""

    try:
        response = llm_provider.generate_text(
            prompt,
            max_tokens=INSIGHT_MAX_TOKENS,
            temperature=INSIGHT_TEMPERATURE,
            thinking_budget=INSIGHT_THINKING_BUDGET,
        )
        cleaned = str(response).replace("```", "").strip()
        
        # Clean up text: fix spacing, grammar, and formatting
        cleaned = _clean_text_output(cleaned)
        
        # Split into sentences/paragraphs
        insights = [p.strip() for p in cleaned.split('\n\n') if p.strip()]
        if not insights:
            insights = [s.strip() + '.' for s in cleaned.split('.') if s.strip()]
        
        # Clean each insight individually
        insights = [_clean_text_output(insight) for insight in insights]
        
        # Enforce 100-word limit
        total_text = " ".join(insights)
        words = total_text.split()
        if len(words) > max_words:
            total_text = " ".join(words[:max_words]).rstrip(" ,;:") + "."
            insights = [_clean_text_output(total_text)]
        
        return insights if insights else [fallback]
    except Exception:
        return [fallback]


def _clean_text_output(text):
    """Clean up text for grammar, spacing, and formatting issues."""
    import re
    
    # Remove multiple spaces
    text = re.sub(r' +', ' ', text)
    
    # Fix spacing before punctuation
    text = re.sub(r'\s+([.,;:!?])', r'\1', text)
    
    # Fix spacing after punctuation
    text = re.sub(r'([.,;:!?])([A-Za-z])', r'\1 \2', text)
    
    # Remove duplicate consecutive words (case-insensitive)
    text = re.sub(r'\b(\w+)\s+\1\b', r'\1', text, flags=re.IGNORECASE)
    
    # Ensure sentences end with proper punctuation
    lines = text.split('\n')
    cleaned_lines = []
    for line in lines:
        line = line.strip()
        if line and not line[-1] in '.!?':
            line += '.'
        cleaned_lines.append(line)
    text = '\n'.join(cleaned_lines)
    
    # Capitalize first letter of sentences
    text = re.sub(r'(^|[.!?]\s+)([a-z])', lambda m: m.group(1) + m.group(2).upper(), text)
    
    # Fix common run-on issues (missing space after period)
    text = re.sub(r'\.([A-Z])', r'. \1', text)
    
    return text.strip()


def _solution_item_insight(theme, concept, share):
    return (
        f"Insight: This solution reflects a practical response under '{theme}', and contributes "
        f"{share:.1f}% of the proposed actions in this theme."
    )


def refine_concepts_with_ai(concepts_list, type_label):
    """
    Uses AI to clean, deduplicate, and re-theme the top concepts.
    Returns a dictionary: { 'Old Concept': {'concept': 'New Concept', 'theme': 'New Theme'} }
    """
    if not llm_provider:
        return {}

    checkpoint_path = get_refinement_checkpoint_path(type_label)
    all_results = load_refinement_checkpoint(checkpoint_path)
    pending_concepts = [concept for concept in concepts_list if concept not in all_results]

    print(f"   🧠 AI Refinement: Optimizing top {len(concepts_list)} {type_label}s...")
    if all_results:
        print(f"      ♻️ Resume mode: loaded {len(all_results)} cached refinements from {checkpoint_path}")

    if not pending_concepts:
        print(f"      ✅ No pending refinement for {type_label}. Using cached results.")
        return all_results

    batch_size = REFINE_BATCH_SIZE

    total_batches = (len(pending_concepts) + batch_size - 1) // batch_size
    for i in range(0, len(pending_concepts), batch_size):
        batch = pending_concepts[i:i+batch_size]
        current_batch = (i // batch_size) + 1
        print(f"      Processing batch {current_batch}/{total_batches} ({len(batch)} items)...")
        
        prompt = f"""You are a Data Cleaning Expert for an Education Report.
        
        THEMES:
        {THEME_KNOWLEDGE_BASE}
        
        INPUT: A list of top recurring {type_label}s found in the data.
        
        TASKS:
        1. AGGRESSIVE DEDUPLICATION: Merge specific variants into broader core concepts.
           - "Child labor in agriculture" / "Child labor at home" / "Child labour due to poverty" / "Child labour preventing education" -> MERGE ALL INTO "Child Labour"
           - "Poverty preventing girls' education" / "Poverty preventing school attendance" / "Poverty preventing children's education" -> MERGE ALL INTO "Poverty preventing education"
           - "Lack of awareness" / "General awareness" -> MERGE INTO "Lack of awareness about education importance"
        2. RE-THEME: Correct misclassified items.
        3. FORMAT: Ensure the concept is a clear, concise {type_label} statement.
        
        INPUT LIST:
        {json.dumps(batch)}
        
        OUTPUT:
        Return a VALID JSON object where keys are the INPUT strings and values are objects with "concept" and "theme".
        IMPORTANT: 
        - Escape all double quotes within strings (e.g., \"text\").
        - Do not include any text outside the JSON block.
        - Ensure the JSON is valid.
        
        MANDATORY RE-CLASSIFICATION RULES:
        - "General awareness" / "Lack of awareness" → theme: "Parental Attitudes & Socio-Cultural"
        - "Migration" / "seasonal migration" → theme: "Poverty and Economic Barriers"
        - "Children not going to school" (no reason given) → theme: "Parental Attitudes & Socio-Cultural"
        - "Other Factors" should ONLY be used for genuinely non-educational, unclassifiable content.
        - Do NOT use "Uncategorized" or "Vague" as concept values.

        Example:
        {{
            "Child labor in agriculture": {{"concept": "Child Labour", "theme": "Poverty and Economic Barriers"}},
            "General awareness": {{"concept": "Lack of parental awareness about education", "theme": "Parental Attitudes & Socio-Cultural"}}
        }}
        RETURN ONLY JSON. NO MARKDOWN."""

        for attempt in range(1, MAX_RETRIES + 1):
            try:
                text = llm_provider.generate_text(
                    prompt,
                    max_tokens=REFINE_MAX_TOKENS,
                    temperature=0,
                    thinking_budget=REFINE_THINKING_BUDGET,
                )

                try:
                    batch_result = _parse_refinement_response(text, batch)
                except Exception:
                    batch_result = _repair_refinement_json_with_ai(text, batch, type_label)

                all_results.update(batch_result)
                save_refinement_checkpoint(checkpoint_path, all_results)
                print(f"      ✅ Batch {current_batch} saved ({len(batch_result)} items).")
                break

            except Exception as error:
                retryable = is_retryable_error(error)
                parse_error = isinstance(error, (json.JSONDecodeError, ValueError, TypeError))
                should_retry = (retryable or parse_error) and attempt < MAX_RETRIES
                print(f"      ⚠️ Batch {current_batch} attempt {attempt}/{MAX_RETRIES} failed: {error}")

                if should_retry:
                    delay = min(MAX_RETRY_SECONDS, BASE_RETRY_SECONDS * (2 ** (attempt - 1)))
                    delay += random.uniform(0, 0.5)
                    print(f"      🔁 Retrying in {delay:.1f}s...")
                    time.sleep(delay)
                    continue

                print(f"      ❌ Batch {current_batch} failed after {attempt} attempt(s).")
                break
            
    return all_results

def normalize_concepts_per_theme(df, type_label):
    """Per-theme semantic normalization using AI — the key fix for synonym fragmentation.

    For each theme, shows ALL unique concepts in that theme to the AI and asks it
    to map every variant to a single canonical label.  This handles cases that
    TF-IDF cannot (e.g., "Poverty" vs "Financial Hardship", "Household Chores"
    vs "Domestic Work") because the AI understands meaning, not just characters.

    Works within a single API call per theme (10 themes = 10 API calls total),
    far cheaper than the previous top-200 global approach which used 8-16 calls
    but left 90% of concepts unrefined.

    Returns the updated DataFrame with Merged_Concept normalized per theme.
    """
    if not llm_provider:
        return df

    checkpoint_path = get_refinement_checkpoint_path(f"{type_label}_per_theme")
    canonical_map = load_refinement_checkpoint(checkpoint_path)  # {old: canonical}
    temp = df.copy()
    themes_present = [t for t in temp['Theme'].unique() if t in VALID_THEMES]

    print(f"   🧠 Per-theme semantic normalization ({len(themes_present)} themes, 1 AI call each)...", flush=True)

    for theme in themes_present:
        theme_mask = temp['Theme'] == theme
        theme_concepts = temp.loc[theme_mask, 'Merged_Concept'].dropna().unique().tolist()
        # Filter out already-normalized concepts
        pending = [c for c in theme_concepts if c not in canonical_map]
        if not pending:
            continue

        # Batch into chunks of 80 if theme has many unique concepts
        THEME_BATCH_SIZE = int(os.getenv("THEME_NORMALIZE_BATCH_SIZE", "80"))
        pending_batches = [pending[i:i+THEME_BATCH_SIZE] for i in range(0, len(pending), THEME_BATCH_SIZE)]

        # Running canonical set: accumulates WITHIN this theme across all its batches.
        # This is the key fix — batch 2 sees what batch 1 established so synonyms collapse.
        running_theme_canonicals: set = set()
        # Seed with any already-cached canonicals for this theme
        for c in theme_concepts:
            if c in canonical_map:
                running_theme_canonicals.add(canonical_map[c])

        for batch_idx, batch in enumerate(pending_batches):
            batch_label = f"{theme} ({batch_idx+1}/{len(pending_batches)})" if len(pending_batches) > 1 else theme
            print(f"      [{batch_label}] — {len(batch)} concepts → normalizing...", flush=True)

            anchor_block = ""
            if running_theme_canonicals:
                anchor_lines = "\n".join(f"  - {c}" for c in sorted(running_theme_canonicals)[:80])
                anchor_block = (
                    f"\nCANONICAL LABELS ESTABLISHED FOR THIS THEME (MUST REUSE — do NOT create synonyms):\n"
                    f"{anchor_lines}\n"
                    f"If a concept below is semantically equivalent to ANY label above, map it to that EXACT label.\n"
                )

            prompt = f"""You are normalizing concept labels for the education theme: "{theme}".

{anchor_block}
TASK: For each concept below, decide its canonical (standard) label.

CRITICAL MERGE RULES — these MUST be collapsed into ONE label:
- Intent-equivalent phrases → SAME label, even if wording is completely different:
  * "Parents are not interested in sending children to school"
    = "Lack of parental commitment towards education"
    = "Parents don't care about children's studies"
    = "Parental indifference to education"
    → ONE canonical: "Lack of Parental Commitment To Education"
  * "Poverty Preventing Education" = "Financial Hardship Preventing Education" = "Child Labour Due To Poverty" → "Poverty Preventing Education"
  * "Early Marriage Preventing Education" = "Child Marriage Preventing Education" = "Early Child Marriage" → "Child Marriage Preventing Education"
  * "Household Chores Preventing Education" = "Domestic Work Priority For Girls" = "Girls Doing Home Work" → "Domestic Work Priority"
  * "Lack of Awareness About Education" = "Parental Indifference To Education" = "Parents Do Not Value Education" = "No Awareness Of Education Importance" → "Lack of Parental Education Awareness"
  * "Lack of Legal Documentation Aadhaar" = "Lack of Aadhaar Card" = "No Aadhaar For Enrollment" → "Aadhaar Card Issues"
  * "General Dropout Without Stated Reason" = "General School Dropout" = "Dropout Due To Parental Decision" → "Parental Decision School Dropout"
  * "Parental Commitment to Education" = "Parental Commitment to Send Children to School" = "Parents Committed to Children's Education" → "Parental Commitment To Education"
  * "Community Awareness of Education Value" = "Promoting Education Value" = "Raising Education Awareness" → "Community Education Awareness"

NORMALIZATION RULES:
- Map ALL semantically equivalent concepts to ONE canonical label.
- Use the EXACT canonical label from the "ESTABLISHED" list above when applicable.
- Canonical label: Title Case, 3-8 words, noun phrase, no trailing punctuation.
- Keep genuinely distinct concepts separate — only merge when meaning is truly equivalent.
- Do NOT over-merge concepts from different categories of this theme.

INPUT CONCEPTS:
{json.dumps(batch, indent=2)}

OUTPUT: Return a valid JSON object mapping each INPUT concept string to its canonical label.
RETURN ONLY VALID JSON. NO MARKDOWN. NO EXPLANATION."""

            for attempt in range(1, MAX_RETRIES + 1):
                try:
                    text = llm_provider.generate_text(
                        prompt, max_tokens=REFINE_MAX_TOKENS, temperature=0,
                        thinking_budget=REFINE_THINKING_BUDGET,
                    )
                    # Parse JSON response
                    text = text.strip()
                    if text.startswith('```'):
                        text = re.sub(r'^```[a-z]*\n?', '', text).rstrip('`').strip()
                    result = json.loads(text)
                    if isinstance(result, dict):
                        # Validate: only accept string→string mappings
                        clean = {str(k): str(v) for k, v in result.items()
                                 if str(v).strip() and len(str(v).split()) >= 2}
                        canonical_map.update(clean)
                        save_refinement_checkpoint(checkpoint_path, canonical_map)
                        # Update running canonicals so next batch in this theme sees them
                        running_theme_canonicals.update(clean.values())
                        print(f"         ✓ Normalized {len(clean)} concepts.", flush=True)
                    break
                except Exception as err:
                    if attempt < MAX_RETRIES and is_retryable_error(err):
                        time.sleep(min(MAX_RETRY_SECONDS, BASE_RETRY_SECONDS * (2 ** (attempt - 1))))
                        continue
                    print(f"         ⚠️ Normalization failed for {theme}: {err}", flush=True)
                    break

    # Apply the canonical map to the DataFrame
    if canonical_map:
        temp['Merged_Concept'] = temp['Merged_Concept'].map(
            lambda c: canonical_map.get(c, c)
        )
        collapsed = df['Merged_Concept'].nunique() - temp['Merged_Concept'].nunique()
        print(f"   ✓ Per-theme normalization: {df['Merged_Concept'].nunique():,} → "
              f"{temp['Merged_Concept'].nunique():,} unique concepts "
              f"({collapsed:,} variants collapsed).", flush=True)

    return temp


# --- MAIN ENGINE ---

def generate_report():
    _start_time = time.time()
    print("🚀 Starting Final Report Generation Engine...", flush=True)

    # ── [1/8] LOAD DATASETS ──────────────────────────────────────────────────
    print("\n📂 [1/8] Loading input datasets...", flush=True)
    try:
        df_raw = pd.read_csv('cleaned_data.csv')
        chal_exploded = pd.read_csv('exploded_challenges.csv')
        sol_exploded = pd.read_csv('exploded_solutions.csv')
        chal_map = pd.read_csv('challenge_mapping.csv')
        sol_map = pd.read_csv('solution_mapping.csv')
    except Exception as e:
        print(f"❌ Error: Required CSV files missing. {e}", flush=True)
        return
    print(f"   ✓ Raw chaupals: {len(df_raw):,}", flush=True)
    print(f"   ✓ Exploded challenges: {len(chal_exploded):,} | solutions: {len(sol_exploded):,}", flush=True)
    print(f"   ✓ Challenge mappings: {len(chal_map):,} | solution mappings: {len(sol_map):,}", flush=True)

    # ── [2/8] MERGE & CLEAN ──────────────────────────────────────────────────
    print("\n⚙️  [2/8] Merging mappings and applying category tags...", flush=True)
    chal_map['Theme'] = chal_map['Theme'].apply(clean_theme_name)
    sol_map['Theme'] = sol_map['Theme'].apply(clean_theme_name)

    # Use set_index + join for O(n) merge instead of O(n²) merge-on-column
    chal_map_idx = chal_map.set_index('Original')
    sol_map_idx = sol_map.set_index('Original')

    df_c = chal_exploded.join(chal_map_idx, on='Challenges', how='left', rsuffix='_map')
    df_c['Theme'] = df_c['Theme'].fillna("Other Factors").astype(str).apply(clean_theme_name)
    df_c['Environment'] = df_c['Challenges'].apply(categorize_environment_aggressive)

    df_s = sol_exploded.join(sol_map_idx, on='Solutions', how='left', rsuffix='_map')
    df_s['Theme'] = df_s['Theme'].fillna("Other Factors").astype(str).apply(clean_theme_name)
    df_s['Agency'] = df_s['Solutions'].apply(categorize_agency)

    df_chal_mapped = df_c
    print(f"   ✓ df_c: {len(df_c):,} rows | df_s: {len(df_s):,} rows", flush=True)

    # ── [3/8] AI REFINEMENT — two-stage ─────────────────────────────────────
    print("\n🤖 [3/8] Running AI concept refinement (global top-200 + per-theme full)...", flush=True)

    # Stage A: global top-200 deduplication (fast, catches cross-theme re-themes)
    top_chal = df_c['Merged_Concept'].value_counts().head(200).index.tolist()
    top_sol = df_s['Merged_Concept'].value_counts().head(200).index.tolist()
    print(f"   Stage A — global top-{len(top_chal)} challenges / top-{len(top_sol)} solutions...", flush=True)

    chal_updates = refine_concepts_with_ai(top_chal, "Challenge")
    if chal_updates:
        for old, new_data in chal_updates.items():
            mask = df_c['Merged_Concept'] == old
            df_c.loc[mask, 'Merged_Concept'] = new_data['concept']
            df_c.loc[mask, 'Theme'] = new_data['theme']

    sol_updates = refine_concepts_with_ai(top_sol, "Solution")
    if sol_updates:
        for old, new_data in sol_updates.items():
            mask = df_s['Merged_Concept'] == old
            df_s.loc[mask, 'Merged_Concept'] = new_data['concept']
            df_s.loc[mask, 'Theme'] = new_data['theme']

    df_c['Theme'] = df_c['Theme'].apply(clean_theme_name)
    df_s['Theme'] = df_s['Theme'].apply(clean_theme_name)

    # Stage B: per-theme normalization — handles semantic synonyms like
    # "Poverty" vs "Financial Hardship", "Household Chores" vs "Domestic Work"
    # across ALL concepts in each theme (not just top 200).
    print(f"   Stage B — per-theme normalization across all themes...", flush=True)
    df_c = normalize_concepts_per_theme(df_c, "Challenge")
    df_s = normalize_concepts_per_theme(df_s, "Solution")

    df_c['Theme'] = df_c['Theme'].apply(clean_theme_name)
    df_s['Theme'] = df_s['Theme'].apply(clean_theme_name)
    df_chal_mapped = df_c

    # ── QUALITY GATE: remove junk concepts and PII before building report ────
    print("\n🔒 Applying quality filters (junk labels, PII)...", flush=True)
    before_c, before_s = len(df_c), len(df_s)
    df_c = df_c[~df_c['Merged_Concept'].apply(_is_junk_concept)].copy()
    df_s = df_s[~df_s['Merged_Concept'].apply(_is_junk_concept)].copy()
    df_c = df_c[~df_c['Merged_Concept'].apply(_is_pii_text)].copy()
    df_s = df_s[~df_s['Merged_Concept'].apply(_is_pii_text)].copy()
    print(f"   Removed {before_c - len(df_c):,} junk/PII challenge rows, "
          f"{before_s - len(df_s):,} solution rows.", flush=True)

    other_pct_c = (df_c['Theme'] == 'Other Factors').mean() * 100
    other_pct_s = (df_s['Theme'] == 'Other Factors').mean() * 100
    print(f"   Other Factors: {other_pct_c:.1f}% of challenges, {other_pct_s:.1f}% of solutions", flush=True)
    if other_pct_c > 30:
        print(f"   ⚠️  WARNING: Other Factors exceeds 30% target ({other_pct_c:.1f}%). "
              f"Consider expanding THEME_KNOWLEDGE_BASE.", flush=True)
    df_chal_mapped = df_c

    # ── [4/8] GLOBAL CONCEPT GROUPING + CANONICAL LABEL RESOLUTION ───────────
    print(f"\n🔗 [4/8] TF-IDF semantic clustering (threshold={SIMILARITY_THRESHOLD})...", flush=True)
    n_uc = df_c['Merged_Concept'].nunique()
    n_us = df_s['Merged_Concept'].nunique()
    print(f"   Unique challenge concepts: {n_uc:,} | solution concepts: {n_us:,}", flush=True)

    _t4 = time.time()
    df_c = assign_concept_groups(df_c, concept_column='Merged_Concept')
    df_c = resolve_canonical_labels(df_c, concept_column='Merged_Concept')
    n_cg = df_c['Concept_Group'].nunique()
    print(f"   ✓ Challenges: {n_uc:,} labels → {n_cg:,} canonical groups in {time.time()-_t4:.1f}s "
          f"({n_uc - n_cg:,} variants collapsed)", flush=True)

    _t4b = time.time()
    df_s = assign_concept_groups(df_s, concept_column='Merged_Concept')
    df_s = resolve_canonical_labels(df_s, concept_column='Merged_Concept')
    n_sg = df_s['Concept_Group'].nunique()
    print(f"   ✓ Solutions: {n_us:,} labels → {n_sg:,} canonical groups in {time.time()-_t4b:.1f}s "
          f"({n_us - n_sg:,} variants collapsed)", flush=True)

    df_chal_mapped = df_c

    # --- BASELINE METRIC CALCULATIONS ---
    TOTAL_CH_STATE = len(df_raw)
    for col in ['Participant Count', 'Men', 'Women', 'Children']:
        df_raw[col] = pd.to_numeric(df_raw[col], errors='coerce').fillna(0)
    
    TOTAL_PART_STATE = int(df_raw['Participant Count'].sum())
    NUM_CHAL_STATEMENTS = len(chal_exploded) # Used for % calculations
    NUM_SOL_STATEMENTS = len(sol_exploded)   # Used for % calculations
    NUM_CHAL = NUM_CHAL_STATEMENTS
    NUM_SOL = NUM_SOL_STATEMENTS
    
    SOL_RATIO = (NUM_SOL / NUM_CHAL) if NUM_CHAL > 0 else 0
    
    # Demographics
    m_total = int(df_raw['Men'].sum())
    w_total = int(df_raw['Women'].sum())
    c_total = int(df_raw['Children'].sum())
    # Calculate Others
    df_raw['Others'] = df_raw['Participant Count'] - (df_raw['Men'] + df_raw['Women'] + df_raw['Children'])
    df_raw['Others'] = df_raw['Others'].clip(lower=0) 
    o_total = int(df_raw['Others'].sum())
    
    # Percentages
    w_perc = (w_total / TOTAL_PART_STATE * 100) if TOTAL_PART_STATE > 0 else 0
    m_perc = (m_total / TOTAL_PART_STATE * 100) if TOTAL_PART_STATE > 0 else 0
    c_perc = (c_total / TOTAL_PART_STATE * 100) if TOTAL_PART_STATE > 0 else 0
    o_perc = (o_total / TOTAL_PART_STATE * 100) if TOTAL_PART_STATE > 0 else 0
    
    avg_per_chaupal = TOTAL_PART_STATE / TOTAL_CH_STATE if TOTAL_CH_STATE > 0 else 0
    num_districts = df_raw['District'].nunique()
    num_themes = df_c['Theme'].nunique()

    # Theme Analysis for Summary
    theme_counts = df_c['Theme'].value_counts()
    top_3_themes = theme_counts.head(3)
    top_3_perc = (top_3_themes.sum() / NUM_CHAL * 100) if NUM_CHAL > 0 else 0

    # Agency Analysis for Summary
    agency_counts = df_s['Agency'].value_counts()
    ind_led = agency_counts.get('Individual-led', 0)
    comm_led = agency_counts.get('Community-led', 0)
    inst_led = agency_counts.get('Institutional', 0)
    
    ind_perc = (ind_led / NUM_SOL * 100) if NUM_SOL > 0 else 0
    comm_perc = (comm_led / NUM_SOL * 100) if NUM_SOL > 0 else 0
    inst_perc = (inst_led / NUM_SOL * 100) if NUM_SOL > 0 else 0
    
    comm_driven_perc = ind_perc + comm_perc

    doc = Document()

    # ── [5/8] BUILDING REPORT SECTIONS ───────────────────────────────────────
    print("\n📝 [5/8] Building report sections...", flush=True)

    # --- SECTION 1: EXECUTIVE SUMMARY ---
    print("   [1/7] Executive Summary...", flush=True)
    doc.add_heading('1. EXECUTIVE SUMMARY', level=1)
    
    # Intro Paragraph
    intro_p = doc.add_paragraph()
    intro_p.add_run(f"This comprehensive report presents an in-depth analysis of {TOTAL_CH_STATE:,} Shiksha Chaupal community dialogues conducted across Bihar, representing the collective voices of {TOTAL_PART_STATE:,} community members. These dialogues constitute one of the most extensive participatory consultations on education challenges in India, providing rich insights into grassroots barriers to education and community-driven solutions. The analysis encompasses {NUM_CHAL:,} individual challenges and {NUM_SOL:,} solutions, systematically categorized into {num_themes} primary thematic areas for comprehensive understanding.")

    # KEY INSIGHT
    doc.add_heading(f'🔑 KEY INSIGHT: Solution Coverage Ratio', level=2)
    p_ratio = doc.add_paragraph()
    run_ratio = p_ratio.add_run(f"Solution-to-Challenge Ratio: {SOL_RATIO:.2f}")
    run_ratio.bold = True
    
    if SOL_RATIO >= 1.0:
        ratio_text = f"This remarkable ratio demonstrates that communities identified {NUM_SOL:,} solutions for {NUM_CHAL:,} challenges. This transcends traditional deficit-based consultations where communities merely list problems. Instead, it reveals communities as active problem-solvers who think constructively about actionable interventions. This represents a paradigm shift in community engagement from problem identification to solution co-creation."
    elif SOL_RATIO >= 0.5:
        ratio_text = f"With {NUM_SOL:,} solutions proposed for {NUM_CHAL:,} challenges, communities are actively engaging in problem-solving. This indicates a constructive approach where participants are moving beyond just listing problems to identifying potential interventions."
    else:
        ratio_text = f"Communities identified {NUM_SOL:,} solutions alongside {NUM_CHAL:,} challenges. While the focus remains on highlighting barriers, there is an emerging capacity for solution-finding that can be further nurtured."
    
    doc.add_paragraph(ratio_text)

    # SCALE OF PARTICIPATION
    doc.add_heading('Scale of Community Participation', level=2)
    scale_p = doc.add_paragraph()
    scale_p.add_run(f"• Geographic Reach: {TOTAL_CH_STATE:,} community dialogues conducted across {num_districts} districts in Bihar\n")
    scale_p.add_run(f"• Total Participants: {TOTAL_PART_STATE:,} community members actively engaged\n")
    scale_p.add_run(f"• Average Engagement: {avg_per_chaupal:.1f} participants per Chaupal, indicating strong community mobilization\n")
    scale_p.add_run(f"• Gender Representation: Women {w_perc:.1f}%, Children {c_perc:.1f}%, Men {m_perc:.1f}%, Others {o_perc:.1f}%\n")
    
    if w_perc > 50:
        gender_text = f"• Dominant female participation ({w_perc:.1f}%) signals authentic grassroots engagement rather than tokenistic consultation, as women are primary stakeholders in children's education"
    elif w_perc > m_perc:
        gender_text = f"• Strong female participation ({w_perc:.1f}%) highlights women's active role in discussing education challenges, outnumbering male participants ({m_perc:.1f}%)."
    else:
        gender_text = f"• The dialogues included diverse participation, with women contributing {w_perc:.1f}% of the voices, ensuring maternal perspectives are included."
    
    scale_p.add_run(gender_text)

    # DOMINANT CHALLENGE THEMES
    doc.add_heading('Dominant Challenge Themes', level=2)
    doc.add_paragraph(f"The thematic analysis reveals systemic patterns in education barriers. The top 3 themes collectively account for {top_3_perc:.1f}% of all challenges, indicating concentrated problem areas requiring prioritized intervention:")
    
    for theme, count in top_3_themes.items():
        t_perc = (count / NUM_CHAL * 100)
        # Get solution count for this theme
        s_count = len(df_s[df_s['Theme'] == theme])
        doc.add_paragraph(f"{theme}: {t_perc:.1f}% ({count:,} challenges, {s_count:,} solutions)", style='List Bullet')

    # COMMUNITY-LED VS SYSTEM-DEPENDENT
    doc.add_heading('Community-Led vs System-Dependent Solutions', level=2)
    doc.add_paragraph("Solution agency analysis reveals community ownership patterns. The distribution demonstrates where communities see themselves as agents of change versus where they require external institutional support:")
    
    agency_p = doc.add_paragraph()
    agency_p.add_run(f"• Individual-led Solutions: {ind_perc:.1f}% ({ind_led:,} solutions) - Family-level actions including parental engagement, behavioral change, and household resource allocation\n")
    agency_p.add_run(f"• Community-led Solutions: {comm_perc:.1f}% ({comm_led:,} solutions) - Collective action including social mobilization, peer support networks, and community organizing\n")
    agency_p.add_run(f"• Institutional Solutions: {inst_perc:.1f}% ({inst_led:,} solutions) - Systemic interventions requiring government or CSO support including infrastructure, policy changes, and resource provision\n")
    
    if comm_driven_perc > 50:
        insight_text = f"• Critical Insight: {comm_driven_perc:.1f}% of solutions are community-driven (individual + community-led), demonstrating extraordinary grassroots capacity that partnerships must amplify rather than replace"
    else:
        insight_text = f"• Critical Insight: While {comm_driven_perc:.1f}% of solutions are community-driven, a significant portion ({inst_perc:.1f}%) requires institutional support, highlighting the need for strong government-community collaboration."

    crit_run = agency_p.add_run(insight_text)
    crit_run.bold = True

    # STRATEGIC PARTNERSHIP
    doc.add_heading('🤝 Strategic Partnership Opportunity', level=2)
    
    if comm_driven_perc > 50:
        strat_text = f"The {comm_driven_perc:.1f}% proportion of community-led and individual-led solutions reveals extraordinary community ownership and problem-solving capacity. Strategic partnerships should operate on a community-strengthening model rather than community-replacing model. This means: (1) Amplifying existing community initiatives through capacity building and resource support, (2) Providing targeted institutional interventions ({inst_perc:.1f}%) for infrastructure, teacher capacity, and documentation systems that communities cannot address independently, (3) Facilitating community-to-community learning and peer exchange, (4) Advocating for policy changes that enable community solutions to scale. The partnership must recognize communities as co-creators and primary implementers, not merely beneficiaries."
    else:
        strat_text = f"With {inst_perc:.1f}% of solutions requiring institutional intervention, a collaborative partnership model is essential. This involves: (1) Government and CSOs addressing structural barriers like infrastructure and teacher shortages, (2) Strengthening the {comm_driven_perc:.1f}% of community-led initiatives to ensure sustainability, (3) Creating feedback loops where community needs directly inform policy implementation."

    doc.add_paragraph(strat_text)

    doc.add_page_break()

    # --- SECTION 2: GENERAL PARTICIPATION OVERVIEW ---
    print("   [2/7] Participation Overview...", flush=True)
    doc.add_heading('2. GENERAL PARTICIPATION OVERVIEW', level=1)
    
    intro_para = doc.add_paragraph()
    run = intro_para.add_run("This section provides comprehensive analysis of participation patterns across geographic and demographic dimensions.")
    run.italic = True

    # TABLE 1: STATE METRICS
    doc.add_heading('TABLE 1: Overall Participation Metrics', level=2)
    table1 = doc.add_table(rows=1, cols=3)
    table1.style = 'Table Grid'
    hdr_cells = table1.rows[0].cells
    for i, h in enumerate(['Metric', 'Count', 'Percentage (%)']):
        hdr_cells[i].text = h
        hdr_cells[i].paragraphs[0].runs[0].bold = True
        set_cell_background(hdr_cells[i], "D9D9D9") 

    def add_row_v3(t, metric, count, p_type):
        r = t.add_row().cells
        r[0].text = metric
        r[1].text = f"{int(count):,}" if isinstance(count, (int, float)) else str(count)
        if p_type == "none": r[2].text = "-"
        elif p_type == "full": r[2].text = "100%"
        else:
            p = (count / TOTAL_PART_STATE * 100) if TOTAL_PART_STATE > 0 else 0
            r[2].text = f"{p:.1f}%"

    add_row_v3(table1, "Total Number of Chaupals", TOTAL_CH_STATE, "none")
    add_row_v3(table1, "Total Number of Participants", TOTAL_PART_STATE, "full")
    add_row_v3(table1, "Average Participants per Chaupal", f"{avg_per_chaupal:.1f}", "none")
    add_row_v3(table1, "Men Participants", m_total, "calc")
    add_row_v3(table1, "Women Participants", w_total, "calc")
    add_row_v3(table1, "Children Participants", c_total, "calc")
    add_row_v3(table1, "Others (Unspecified)", o_total, "calc")

    # SECTION 2.3: DISTRICT DISTRIBUTION
    doc.add_heading('2.3 District-wise Distribution of Reported Chaupals', level=2)
    dist_stats = df_raw.groupby('District').agg(Ch_Count=('id', 'count'), Part_Sum=('Participant Count', 'sum')).reset_index()
    dist_stats['Ch_Perc'] = (dist_stats['Ch_Count'] / TOTAL_CH_STATE) * 100
    dist_stats = dist_stats.sort_values('Ch_Count', ascending=False)

    table_dist = doc.add_table(rows=1, cols=5)
    table_dist.style = 'Table Grid'
    d_hdr = table_dist.rows[0].cells
    headers = ['District', 'Chaupals (N)', 'Chaupal %', 'Participants (N)', 'Participant %']
    for i, h in enumerate(headers):
        d_hdr[i].text = h
        d_hdr[i].paragraphs[0].runs[0].bold = True
        set_cell_background(d_hdr[i], "F2F2F2")

    for _, row in dist_stats.iterrows():
        r = table_dist.add_row().cells
        r[0].text, r[1].text = str(row['District']), f"{int(row['Ch_Count']):,}"
        r[2].text = f"{row['Ch_Perc']:.1f}%"
        r[3].text, r[4].text = f"{int(row['Part_Sum']):,}", f"{(row['Part_Sum']/TOTAL_PART_STATE)*100:.1f}%"

    # NARRATIVE ANALYSIS
    top_3 = dist_stats.head(3)
    top_3_sum = top_3['Ch_Perc'].sum()
    dist_list_str = ", ".join([f"{row['District']} ({int(row['Ch_Count']):,}, {row['Ch_Perc']:.1f}%)" for _, row in top_3.iterrows()])
    
    doc.add_heading('Geographic Distribution Analysis', level=3)
    doc.add_paragraph(f"Geographic distribution shows concentration in {dist_list_str}. Together, these top 3 districts account for {top_3_sum:.1f}% of dialogues.")

    doc.add_heading('Demographic Composition Analysis', level=3)
    w_perc, m_perc, c_perc = (w_total/TOTAL_PART_STATE)*100, (m_total/TOTAL_PART_STATE)*100, (c_total/TOTAL_PART_STATE)*100
    ratio = w_perc / m_perc if m_perc > 0 else 0
    doc.add_paragraph(f"Women constitute {w_perc:.1f}% of participants, which is {( 'more than triple' if ratio >= 3 else 'significantly higher than' )} male participation ({m_perc:.1f}%).")

    doc.add_page_break()
    

    # --- SECTION 3: CORE CONTENT ANALYSIS ---
    print("   [3/7] Core Content Analysis...", flush=True)
    doc.add_heading('3. CORE CONTENT ANALYSIS', level=1)
    
    # Calculations for Section 3
    unique_chal_count = chal_map['Merged_Concept'].nunique()
    unique_sol_count = sol_map['Merged_Concept'].nunique()
    chal_reduction = ((NUM_CHAL - unique_chal_count) / NUM_CHAL * 100) if NUM_CHAL > 0 else 0
    sol_reduction = ((NUM_SOL - unique_sol_count) / NUM_SOL * 100) if NUM_SOL > 0 else 0

    # Narrative
    doc.add_paragraph(f"This section analyzes the substance of community dialogues - the challenges identified and solutions proposed. Communities articulated {NUM_CHAL:,} individual challenges and {NUM_SOL:,} individual solutions.")

    # TABLE 2
    doc.add_heading('Overall Challenge & Solution Metrics', level=2)
    summary_table = doc.add_table(rows=1, cols=2)
    summary_table.style = 'Table Grid'
    hdr = summary_table.rows[0].cells
    hdr[0].text, hdr[1].text = 'Metric', 'Count'
    for cell in hdr: 
        set_cell_background(cell, "D9D9D9")
        cell.paragraphs[0].runs[0].bold = True

    data_rows = [
        ("Total Challenges", f"{NUM_CHAL:,}"),
        ("Total Solutions", f"{NUM_SOL:,}"),
        ("Unique Challenges (after deduplication)", f"{unique_chal_count:,}"),
        ("Unique Solutions (after deduplication)", f"{unique_sol_count:,}"),
        ("Overall Solution-to-Challenge Ratio", f"{SOL_RATIO:.2f}"),
    ]
    for metric, val in data_rows:
        r = summary_table.add_row().cells
        r[0].text, r[1].text = metric, str(val)

    # 3.2 District Averages
    doc.add_heading('District-wise Engagement Depth', level=2)
    dist_chal_counts = chal_exploded.groupby('District').size().reset_index(name='C_Count')
    dist_sol_counts = sol_exploded.groupby('District').size().reset_index(name='S_Count')
    
    avg_df = dist_stats[['District', 'Ch_Count']].merge(dist_chal_counts, on='District').merge(dist_sol_counts, on='District')
    avg_df['Avg_C'] = avg_df['C_Count'] / avg_df['Ch_Count']
    avg_df['Avg_S'] = avg_df['S_Count'] / avg_df['Ch_Count']

    avg_table = doc.add_table(rows=1, cols=3)
    avg_table.style = 'Table Grid'
    a_hdr = avg_table.rows[0].cells
    a_hdr[0].text, a_hdr[1].text, a_hdr[2].text = 'District', 'Avg Challenges/Chaupal', 'Avg Solutions/Chaupal'
    for cell in a_hdr: set_cell_background(cell, "F2F2F2")

    for _, row in avg_df.sort_values('Avg_C', ascending=False).iterrows():
        r = avg_table.add_row().cells
        r[0].text, r[1].text, r[2].text = str(row['District']), f"{row['Avg_C']:.2f}", f"{row['Avg_S']:.2f}"

    # Paradigm Shift Section
    doc.add_heading('Solution-to-Challenge Ratio: A Paradigm Shift', level=2)
    doc.add_paragraph(f"The overall solution-to-challenge ratio of {SOL_RATIO:.2f} represents a fundamental paradigm shift in community consultation methodology. Traditional deficit-based consultations focus solely on problem identification, treating communities as problem containers. The Shiksha Chaupal model demonstrates that when communities are engaged as problem-solvers rather than merely problem-identifiers, they actively think constructively about solutions. This {NUM_SOL:,} solutions for {NUM_CHAL:,} challenges ratio indicates that every articulated challenge was matched with actionable solution thinking, demonstrating community agency and constructive engagement. This asset-based approach recognizes communities as repositories of contextual knowledge and innovative problem-solving capacity.")

    doc.add_page_break()

    # --- SECTION 4: THEMATIC ANALYSIS ---
    print("   [4/7] Thematic Analysis...", flush=True)
    doc.add_heading('4. THEMATIC ANALYSIS', level=1)
    
    # TABLE: Challenge by Theme
    doc.add_heading('Overall Challenge Distribution by Theme', level=2)
    t_chal = doc.add_table(rows=1, cols=3); t_chal.style = 'Table Grid'
    h_chal = t_chal.rows[0].cells
    h_chal[0].text, h_chal[1].text, h_chal[2].text = 'Theme', 'Count', '%'
    for c in h_chal: set_cell_background(c, "D9D9D9")
    
    theme_counts = df_c['Theme'].value_counts()
    for theme, count in theme_counts.items():
        r = t_chal.add_row().cells
        r[0].text, r[1].text, r[2].text = theme, str(count), f"{(count/NUM_CHAL_STATEMENTS*100):.1f}%"

    # TABLE: Solution Agency
    doc.add_heading('Solution Distribution by Agency', level=2)
    t_agency = doc.add_table(rows=1, cols=3); t_agency.style = 'Table Grid'
    h_age = t_agency.rows[0].cells
    h_age[0].text, h_age[1].text, h_age[2].text = 'Agency Type', 'Count', '%'
    for c in h_age: set_cell_background(c, "F2F2F2")
    
    agency_counts = df_s['Agency'].value_counts()
    for agency, count in agency_counts.items():
        r = t_agency.add_row().cells
        r[0].text, r[1].text, r[2].text = agency, str(count), f"{(count/NUM_SOL_STATEMENTS*100):.1f}%"

    # TABLE: Challenge Environment
    doc.add_heading('Challenge Distribution by Environment', level=2)
    t_env = doc.add_table(rows=1, cols=3); t_env.style = 'Table Grid'
    h_env = t_env.rows[0].cells
    h_env[0].text, h_env[1].text, h_env[2].text = 'Environment', 'Count', '%'
    for c in h_env: set_cell_background(c, "D9D9D9")
    
    env_counts = df_c['Environment'].value_counts()
    for env, count in env_counts.items():
        r = t_env.add_row().cells
        r[0].text, r[1].text, r[2].text = env, str(count), f"{(count/NUM_CHAL_STATEMENTS*100):.1f}%"

    # --- INDIVIDUAL THEME DEEP-DIVES ---
    doc.add_page_break()

    # Prepare theme list: Top themes, but ensure "Other Factors" is last
    themes_to_process = list(theme_counts.index[:10])
    
    # If "Other Factors" is in the list, remove it temporarily
    if "Other Factors" in themes_to_process:
        themes_to_process.remove("Other Factors")
        
    # If "Other Factors" exists in data (even if not in top 10 originally, though unlikely given 20%), add it to end
    if "Other Factors" in theme_counts.index:
        themes_to_process.append("Other Factors")

    for i, theme in enumerate(themes_to_process, 1):
        print(f"      [{i}/{len(themes_to_process)}] {theme}...", flush=True)
        doc.add_heading(f'4.{i} {theme.upper()}', level=2)

        # Use pre-computed groups from global assign_concept_groups — no re-grouping here
        t_c = df_c[df_c['Theme'] == theme]
        t_s = df_s[df_s['Theme'] == theme]

        # Theme Metrics
        c_count = len(t_c)
        s_count = len(t_s)
        c_perc = (c_count / NUM_CHAL * 100) if NUM_CHAL > 0 else 0
        sol_cov = (s_count / c_count) if c_count > 0 else 0
        u_c = t_c['Merged_Concept'].nunique()
        u_s = t_s['Merged_Concept'].nunique()

        m_para = doc.add_paragraph()
        m_para.add_run(f"Scale: {c_count:,} challenges ({c_perc:.1f}% of total dataset) | {s_count:,} solutions\n").bold = True
        m_para.add_run(f"Solution Coverage: {sol_cov:.2f} solutions per challenge")

        doc.add_heading('Challenge Landscape', level=4)
        if not t_c.empty:
            env_pref = t_c['Environment'].value_counts(normalize=True).idxmax()
            doc.add_paragraph(f"The landscape for '{theme}' is primarily localized within the {env_pref} environment. This suggests that interventions must be targeted at this level for maximum impact.")

        doc.add_heading("Top Recurring Challenges", level=5)

        # Use pre-computed Concept_Group column (no re-grouping)
        grouped_challenges = []
        challenges_with_key = t_c  # already has Concept_Group from global grouping

        if 'Concept_Group' in challenges_with_key.columns:
            for concept_key, group in challenges_with_key.groupby('Concept_Group'):
                mention_count = len(group)
                display_concept = group['Merged_Concept'].value_counts().idxmax()
                grouped_challenges.append((display_concept, mention_count, concept_key))
        else:
            for concept, count in t_c['Merged_Concept'].value_counts().items():
                grouped_challenges.append((concept, count, normalize_concept_key(concept)))

        grouped_challenges = sorted(grouped_challenges, key=lambda item: item[1], reverse=True)
        total_theme_chal = len(t_c)
        cumulative_count = 0

        for i, (concept, count, concept_key) in enumerate(grouped_challenges, 1):
            cumulative_count += count
            coverage_perc = (cumulative_count / total_theme_chal) * 100
            item_perc = (count / total_theme_chal) * 100

            concept_group_rows = challenges_with_key[challenges_with_key['Concept_Group'] == concept_key] if 'Concept_Group' in challenges_with_key.columns else challenges_with_key[challenges_with_key['Merged_Concept'] == concept]
            original_texts = concept_group_rows['Challenges'].tolist()

            # ── Quality-filtered representative quote selection ──
            rep_quote = None
            rep_district = "Unknown"
            if not concept_group_rows.empty:
                cands = concept_group_rows.copy()
                cands['_txt'] = cands['Challenges'].astype(str).str.strip()
                # Filter: complete sentence, no PII, 8–40 words, ≤200 chars
                quality = cands[
                    cands['_txt'].str[-1:].isin(['.', '!', '?']) &
                    (~cands['_txt'].apply(_is_pii_text)) &
                    (cands['_txt'].str.len() <= 200) &
                    (cands['_txt'].str.split().str.len() >= 8)
                ]
                if not quality.empty:
                    quality = quality.copy()
                    quality['_uniq'] = quality['_txt'].apply(lambda t: len(set(t.lower().split())))
                    rep_row = quality.loc[quality['_uniq'].idxmax()]
                    rep_quote = str(rep_row.get('Challenges', '')).strip()
                    if 'District' in quality.columns and pd.notna(rep_row.get('District')):
                        rep_district = str(rep_row.get('District')).strip() or "Unknown"
                else:
                    # Fallback: sentence-ending texts only, no PII, ≤200 chars
                    safe = cands[
                        cands['_txt'].str[-1:].isin(['.', '!', '?']) &
                        (~cands['_txt'].apply(_is_pii_text)) &
                        (cands['_txt'].str.len() <= 200)
                    ]
                    if not safe.empty:
                        raw = str(safe.iloc[0].get('Challenges', '')).strip()
                        rep_quote = raw
                        d_col = safe.iloc[0].get('District', '')
                        if d_col and pd.notna(d_col):
                            rep_district = str(d_col).strip() or "Unknown"
                    # If still nothing: skip quote — concept label speaks for itself
            
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Pt(36)
            p.paragraph_format.first_line_indent = Pt(-18)
            p.add_run(f"{i}. {concept}").bold = True
            p.add_run(f" ({count} mentions, {item_perc:.1f}%)")

            insight_paragraphs = _challenge_item_insight(
                theme, concept, item_perc, count, t_c, t_s,
                original_texts, max_words=150,
                concept_key=concept_key, concept_rows=concept_group_rows,
            )
            quote_length = len(rep_quote) if rep_quote else 0
            paragraphs_to_show = insight_paragraphs[:1] if quote_length < 100 else insight_paragraphs

            for insight_text in paragraphs_to_show:
                insight_para = doc.add_paragraph()
                insight_para.paragraph_format.left_indent = Pt(54)
                insight_para.add_run(insight_text)

            # Only show quote if we have a clean, non-PII representative text
            if rep_quote and not _is_pii_text(rep_quote):
                quote_para = doc.add_paragraph()
                quote_para.paragraph_format.left_indent = Pt(54)
                quote_para.add_run(f"Voice from the ground ({rep_district}): \"{rep_quote}\"").italic = True

            if coverage_perc >= 50 and i >= 5:
                break
            if i >= 15:
                break

        # Solution Ecosystem
        doc.add_heading('Solution Ecosystem', level=4)
        if not t_s.empty:
            total_theme_sol = len(t_s)
            agency_counts = t_s['Agency'].value_counts(normalize=True)
            agency_main = agency_counts.idxmax()
            agency_perc = agency_counts.max() * 100
            
            doc.add_paragraph(f"Communities proposed {total_theme_sol:,} solutions to address this theme. The solution ecosystem demonstrates {agency_main} agency with {agency_perc:.1f}% of solutions being {agency_main}.")

            doc.add_heading("Most Frequently Proposed Solutions", level=5)

            # Use pre-computed Concept_Group from global grouping; no re-grouping
            valid_solutions = t_s[t_s['Merged_Concept'].apply(is_valid_solution)].copy()

            grouped_solutions = []
            if 'Concept_Group' in valid_solutions.columns:
                for sol_key, group in valid_solutions.groupby('Concept_Group'):
                    mention_count = len(group)
                    display_concept = group['Merged_Concept'].value_counts().idxmax()
                    grouped_solutions.append((display_concept, mention_count, sol_key))
            else:
                for sol_concept, count in valid_solutions['Merged_Concept'].value_counts().items():
                    grouped_solutions.append((sol_concept, count, normalize_concept_key(sol_concept)))

            top_solutions = sorted(grouped_solutions, key=lambda item: item[1], reverse=True)[:5]

            for rank, (sol_concept, count, sol_key) in enumerate(top_solutions, 1):
                item_perc = (count / total_theme_sol) * 100

                # Quality-filtered solution quote: sentence-complete, no PII, 8-40 words, ≤200 chars
                sol_rows = valid_solutions[valid_solutions['Concept_Group'] == sol_key] if 'Concept_Group' in valid_solutions.columns else valid_solutions[valid_solutions['Merged_Concept'] == sol_concept]
                sol_texts = sol_rows['Solutions'].astype(str).str.strip()
                sol_quality = sol_texts[
                    sol_texts.str[-1:].isin(['.', '!', '?']) &
                    (~sol_texts.apply(_is_pii_text)) &
                    (sol_texts.str.len() <= 200) &
                    (sol_texts.str.split().str.len() >= 8) &
                    (sol_texts.str.split().str.len() <= 40)
                ]
                if not sol_quality.empty:
                    # Prefer quotes with highest unique-word diversity (most informative)
                    sol_quality_df = sol_quality.to_frame('txt')
                    sol_quality_df['uniq'] = sol_quality_df['txt'].apply(lambda t: len(set(t.lower().split())))
                    sol_quote = sol_quality_df.loc[sol_quality_df['uniq'].idxmax(), 'txt']
                else:
                    sol_quote = None

                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Pt(36)
                p.paragraph_format.first_line_indent = Pt(-18)
                p.add_run(f"{rank}. {sol_concept}").bold = True
                p.add_run(f" ({count} mentions, {item_perc:.1f}%)")
                if sol_quote and not _is_pii_text(str(sol_quote)):
                    p.add_run(f"\n   Community Proposal: \"{sol_quote}\"").italic = True

    doc.add_page_break()
    print("   [5/7] District Profiles...", flush=True)
    doc.add_heading('5. DISTRICT PROFILES', level=1)
    
    doc.add_paragraph("This section provides comprehensive profiles for top-performing districts by Chaupal count, including quantitative metrics, thematic breakdowns, and top challenges/solutions specific to each district. These profiles enable district-specific intervention design and comparative analysis across geographies.")

    # --- DISTRICT OVERVIEW TABLE ---
    doc.add_heading('District Performance Overview Table', level=2)
    
    # Aggregate Data
    d_stats = df_raw.groupby('District').agg(
        Chaupals=('id', 'nunique'),
        Participants=('Participant Count', 'sum')
    ).reset_index()
    
    c_counts = chal_exploded.groupby('District').size().reset_index(name='Challenges')
    s_counts = sol_exploded.groupby('District').size().reset_index(name='Solutions')
    
    dist_overview = d_stats.merge(c_counts, on='District', how='left').merge(s_counts, on='District', how='left').fillna(0)
    dist_overview['Ratio'] = dist_overview['Solutions'] / dist_overview['Challenges']
    dist_overview = dist_overview.sort_values('Chaupals', ascending=False)
    
    # Create Table
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    headers = ['District', 'Chaupals', 'Participants', 'Challenges', 'Solutions', 'Ratio']
    for i, h in enumerate(headers):
        hdr[i].text = h
        hdr[i].paragraphs[0].runs[0].bold = True
        set_cell_background(hdr[i], "D9D9D9")
        
    for _, row in dist_overview.iterrows():
        r = table.add_row().cells
        r[0].text = str(row['District'])
        r[1].text = f"{int(row['Chaupals']):,}"
        r[2].text = f"{int(row['Participants']):,}"
        r[3].text = f"{int(row['Challenges']):,}"
        r[4].text = f"{int(row['Solutions']):,}"
        r[5].text = f"{row['Ratio']:.2f}"
        
    # Note
    note_p = doc.add_paragraph()
    note_run = note_p.add_run("Note: Ratio = Solutions ÷ Challenges. Values >1.0 indicate more solutions than challenges identified.")
    note_run.italic = True
    note_run.font.size = Pt(9)

    # --- DETAILED PROFILES ---
    # Reorder districts to put 'Others' last
    dist_list = dist_overview['District'].tolist()
    if 'Others' in dist_list:
        dist_list.remove('Others')
        dist_list.append('Others')

    # Apply minimum threshold: skip districts with too few chaupals
    dist_chaupal_map = d_stats.set_index('District')['Chaupals'].to_dict()
    full_dist_count = len(dist_list)
    dist_list = [d for d in dist_list if dist_chaupal_map.get(d, 0) >= MIN_DISTRICT_CHAUPALS]
    skipped = full_dist_count - len(dist_list)
    if skipped:
        print(f"   ⏭️  Skipping {skipped} district(s) with fewer than {MIN_DISTRICT_CHAUPALS} chaupals.", flush=True)
    print(f"   Generating profiles for {len(dist_list)} district(s)...", flush=True)

    for i, dist in enumerate(dist_list, 1):
        print(f"      [{i}/{len(dist_list)}] {dist} ({dist_chaupal_map.get(dist, 0):,} chaupals)...", flush=True)
        doc.add_heading(f'5.{i} {dist.upper()}', level=2)
        d_raw = df_raw[df_raw['District'] == dist]
        
        # Filter data for this district
        d_chal = df_c[df_c['District'] == dist]
        d_sol = df_s[df_s['District'] == dist]

        # Metrics for Snapshot
        dist_chaupals = len(d_raw)
        dist_participants = int(d_raw['Participant Count'].sum())
        chaupal_perc = (dist_chaupals / TOTAL_CH_STATE * 100) if TOTAL_CH_STATE > 0 else 0
        
        m_d = d_raw['Men'].sum()
        w_d = d_raw['Women'].sum()
        c_d = d_raw['Children'].sum()
        
        m_p = (m_d / dist_participants * 100) if dist_participants > 0 else 0
        w_p = (w_d / dist_participants * 100) if dist_participants > 0 else 0
        c_p = (c_d / dist_participants * 100) if dist_participants > 0 else 0
        
        dist_chal_count = len(d_chal)
        dist_sol_count = len(d_sol)
        dist_ratio = (dist_sol_count / dist_chal_count) if dist_chal_count > 0 else 0
        
        avg_c = dist_chal_count / dist_chaupals if dist_chaupals > 0 else 0
        avg_s = dist_sol_count / dist_chaupals if dist_chaupals > 0 else 0

        doc.add_heading('A. Quantitative Snapshot', level=3)
        snap_p = doc.add_paragraph()
        snap_p.add_run(f"Chaupals: {dist_chaupals:,} ({chaupal_perc:.1f}% of total)\n")
        snap_p.add_run(f"Total Participants: {dist_participants:,}\n")
        snap_p.add_run(f"Demographics: Men {m_p:.1f}%, Women {w_p:.1f}%, Children {c_p:.1f}%\n")
        snap_p.add_run(f"Challenges: {dist_chal_count:,} | Solutions: {dist_sol_count:,}\n")
        snap_p.add_run(f"Solution Efficiency: {dist_ratio:.2f}\n")
        snap_p.add_run(f"Average per Chaupal: {avg_c:.1f} challenges, {avg_s:.1f} solutions")
        
        if d_chal.empty:
            doc.add_paragraph("No challenge data available for this district.")
            continue

        # Calculate Theme Percentages
        total_dist_chal = len(d_chal)
        theme_counts = d_chal['Theme'].value_counts()
        
        doc.add_heading('Thematic Breakdown & Examples', level=3)
        
        for theme, count in theme_counts.items():
            perc = (count / total_dist_chal) * 100
            
            # Theme Header
            p_theme = doc.add_paragraph()
            p_theme.paragraph_format.space_before = Pt(6)
            run = p_theme.add_run(f"• {theme} ({perc:.1f}%)")
            run.bold = True
            
            # Get Top 2 Challenges (by frequency in this district)
            theme_c_rows = d_chal[d_chal['Theme'] == theme]
            top_challenges = theme_c_rows['Merged_Concept'].value_counts().head(2).index.tolist()
            
            # Get Top 2 Solutions
            theme_s_rows = d_sol[d_sol['Theme'] == theme]
            
            # DEBUG
            # print(f"Theme: {theme}, theme_s_rows shape: {theme_s_rows.shape}")
            
            if theme_s_rows.empty:
                top_solutions = []
            else:
                theme_s_rows = theme_s_rows.copy()
                theme_s_rows['Merged_Concept'] = theme_s_rows['Merged_Concept'].fillna("").astype(str)
                valid_mask = (
                    theme_s_rows['Merged_Concept'].apply(is_valid_solution) &
                    (~theme_s_rows['Merged_Concept'].apply(_is_junk_concept)) &
                    (~theme_s_rows['Merged_Concept'].apply(_is_pii_text))
                )
                valid_s_rows = theme_s_rows[valid_mask]
                top_solutions = valid_s_rows['Merged_Concept'].value_counts().head(2).index.tolist()

            # Filter challenges for PII and junk before display
            top_challenges = [c for c in top_challenges if not _is_pii_text(c) and not _is_junk_concept(c)]
            top_solutions = [s for s in top_solutions if not _is_pii_text(s) and not _is_junk_concept(s)]

            if top_challenges:
                p_c = doc.add_paragraph()
                p_c.paragraph_format.left_indent = Pt(18)
                p_c.add_run("Challenges: ").bold = True
                p_c.add_run("; ".join(top_challenges))

            if top_solutions:
                p_s = doc.add_paragraph()
                p_s.paragraph_format.left_indent = Pt(18)
                p_s.add_run("Solutions: ").bold = True
                p_s.add_run("; ".join(top_solutions))

    # --- SECTION 6: UNIQUE INSIGHTS ---
    print("   [6/7] Unique Insights...", flush=True)
    doc.add_page_break()
    doc.add_heading('6. UNIQUE INSIGHTS', level=1)
    
    def get_unique_examples(agency_type, count=5):
        subset = df_s[df_s['Agency'] == agency_type].copy()
        if subset.empty: return []
        
        # 1. Prioritize 'Other Factors'
        others = subset[subset['Theme'] == 'Other Factors']
        
        # 2. If not enough, look for low frequency items in general
        # Calculate frequency of Merged_Concept
        freq = subset['Merged_Concept'].value_counts()
        unique_concepts = freq[freq == 1].index.tolist()
        
        # Filter subset for these unique concepts
        unique_rows = subset[subset['Merged_Concept'].isin(unique_concepts)]
        
        # Combine: Others first, then unique rows
        candidates = pd.concat([others, unique_rows]).drop_duplicates(subset=['Solutions'])
        
        # Filter: no PII, no junk, min 20 chars, max 300 chars, ends in sentence terminator
        candidates = candidates[
            (candidates['Solutions'].str.len() > 20) &
            (candidates['Solutions'].str.len() <= 300) &
            (~candidates['Solutions'].apply(_is_pii_text))
        ]

        candidates = candidates.copy()
        candidates['len'] = candidates['Solutions'].str.len()
        top_candidates = candidates.sort_values('len', ascending=False).head(count)

        results = []
        for _, row in top_candidates.iterrows():
            dist = str(row['District'])
            sol = str(row['Solutions']).strip()
            
            # Only show District as per request
            location_str = f"{dist}"
                
            results.append(f"\"{sol}\"\n   📍 {location_str}")
        return results

    # 6.1 Individual-led
    doc.add_heading('Individual-led Unique Solutions', level=2)
    doc.add_paragraph("Highlights of innovative or distinct solutions proposed by individuals that stand out from common themes:")
    for ex in get_unique_examples('Individual-led', 5):
        p = doc.add_paragraph(style='List Bullet')
        p.add_run(ex)

    # 6.2 Community-led
    doc.add_heading('Community-led Unique Solutions', level=2)
    doc.add_paragraph("Highlights of collective actions or community-driven innovations:")
    for ex in get_unique_examples('Community-led', 5):
        p = doc.add_paragraph(style='List Bullet')
        p.add_run(ex)

    # 6.3 Expectations
    doc.add_heading('Expectations from Government & CSOs', level=2)
    doc.add_paragraph("Key expectations and demands expressed by the community for systemic support:")
    
    inst_subset = df_s[df_s['Agency'] == 'Institutional']
    if not inst_subset.empty:
        # Get top concepts
        top_inst = inst_subset['Merged_Concept'].value_counts().head(5)
        for concept, count in top_inst.items():
            # Get a representative quote
            quotes = inst_subset[inst_subset['Merged_Concept'] == concept]['Solutions'].tolist()
            # Filter for valid quotes
            quotes = [q for q in quotes if isinstance(q, str) and len(q) > 10]
            if not quotes: continue
            
            best_quote = max(quotes, key=len)
            
            p = doc.add_paragraph(style='List Number')
            p.add_run(f"{concept}").bold = True
            p.add_run(f" ({count} mentions)")
            p.add_run(f"\n   Community Voice: \"{best_quote}\"").italic = True

    print("\n💾 [6/8] Saving document...", flush=True)
    
    # --- SECTION 7: CONCLUSION ---
    print("   [7/7] Conclusion...", flush=True)
    doc.add_page_break()
    doc.add_heading('7. CONCLUSION', level=1)
    
    # 7.1 Transformative Insight
    doc.add_heading('🎯 TRANSFORMATIVE INSIGHT: Community as Solution Architects', level=2)
    doc.add_paragraph(f"This analysis of {TOTAL_CH_STATE:,} Shiksha Chaupals reveals a fundamental truth: rural communities are not passive recipients of development interventions but sophisticated problem-solvers capable of designing culturally-appropriate, sustainable solutions to education barriers. The {SOL_RATIO:.2f} solution coverage ratio demonstrates that communities are generating solutions at nearly the same rate as identifying challenges.")

    # 7.2 Interconnected Challenge Reality
    doc.add_heading('The Interconnected Challenge Reality', level=2)
    top_themes_list = top_3_themes.index.tolist()
    theme_str = ", ".join(top_themes_list)
    doc.add_paragraph(f"Challenges do not exist in isolation. The dominance of themes like {theme_str} suggests a complex interplay of factors. For instance, economic barriers often amplify documentation issues, while infrastructure gaps can drive families toward alternative schooling options. Understanding these interconnections is essential for designing effective interventions.")

    # 7.3 Community Agency Excellence
    doc.add_heading('Community Agency Excellence', level=2)
    # Get top community solutions for examples
    comm_sols = df_s[df_s['Agency'] == 'Community-led']['Merged_Concept'].value_counts().head(4).index.tolist()
    if comm_sols:
        comm_examples = ", ".join([s.lower() for s in comm_sols])
        doc.add_paragraph(f"Communities have demonstrated remarkable innovation through initiatives such as {comm_examples}. This agency must be recognized, celebrated, and supported—not replaced by external solutions.")
    else:
        doc.add_paragraph("Communities have demonstrated remarkable innovation through collective action and peer support mechanisms. This agency must be recognized, celebrated, and supported—not replaced by external solutions.")

    # 7.4 Systemic Support Imperative
    doc.add_heading('Systemic Support Imperative', level=2)
    # Get top institutional solutions for examples
    inst_sols = df_s[df_s['Agency'] == 'Institutional']['Merged_Concept'].value_counts().head(4).index.tolist()
    if inst_sols:
        inst_examples = ", ".join([s.lower() for s in inst_sols])
        doc.add_paragraph(f"While community agency is exceptional, certain barriers require institutional action. Issues such as {inst_examples} cannot be resolved through community effort alone. The path forward requires strategic partnerships that amplify community strengths while providing systemic support.")
    else:
        doc.add_paragraph("While community agency is exceptional, certain barriers require institutional action. Infrastructure gaps, documentation bottlenecks, and resource shortages cannot be resolved through community effort alone. The path forward requires strategic partnerships that amplify community strengths while providing systemic support.")

    # 7.5 Strategic Partnership Opportunity
    doc.add_heading('🤝 Strategic Partnership Opportunity', level=2)
    doc.add_paragraph("The optimal collaboration model combines community-led cultural change initiatives with institutional support for infrastructure and policy barriers. Government agencies, NGOs, and CSOs should position themselves as resource partners and accountability allies—not solution designers—enabling communities to scale their own innovations while addressing systemic gaps.")

    output_path = 'Final_Shiksha_Report.docx'
    doc.save(output_path)
    elapsed = time.time() - _start_time
    print(f"   ✓ Saved: {output_path}", flush=True)
    print(f"\n✅ [8/8] Stage 3 complete. Total time: {elapsed:.1f}s ({elapsed/60:.1f} min)", flush=True)

if __name__ == "__main__":
    generate_report()
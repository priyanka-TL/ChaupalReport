import pandas as pd
import os
import time
import random
import re
import unicodedata
from difflib import SequenceMatcher
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
TAGGER_BATCH_SIZE = int(os.getenv("TAGGER_BATCH_SIZE", "150"))
TAGGER_MAX_TOKENS = int(os.getenv("TAGGER_MAX_TOKENS", "12000"))
TAGGER_THINKING_BUDGET = int(os.getenv("TAGGER_THINKING_BUDGET", "512"))
FUZZY_THEME_MIN_SCORE = float(os.getenv("FUZZY_THEME_MIN_SCORE", "0.92"))
RULE_THEME_MIN_SCORE = float(os.getenv("RULE_THEME_MIN_SCORE", "0.60"))
CONCEPT_MERGE_MIN_SCORE = float(os.getenv("CONCEPT_MERGE_MIN_SCORE", "0.74"))
ADAPTIVE_MIN_BATCH_SIZE = int(os.getenv("ADAPTIVE_MIN_BATCH_SIZE", "24"))
ADAPTIVE_MAX_SPLIT_DEPTH = int(os.getenv("ADAPTIVE_MAX_SPLIT_DEPTH", "4"))

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
        r"transfer certificate", r"\bdocument\b", r"id card", r"\bcertificate\b",
        # Extended
        r"age proof", r"\bdomicile\b", r"caste certificate", r"income certificate",
        r"proof of age", r"proof of address", r"official record", r"government document",
        r"school record", r"marksheet", r"admit card", r"\bregistration\b",
        r"admission form", r"birth registration", r"\bunderage\b", r"no certificate",
        r"without.*document", r"document.*missing", r"no.*id", r"id proof",
        r"proof.*missing", r"certificate.*not", r"no.*proof", r"document.*issue",
    ],
    "Poverty and Economic Barriers": [
        r"\bpoverty\b", r"poor financial", r"financial", r"\bno money\b", r"lack of money",
        r"economic", r"labou?r", r"wages?", r"income", r"brick kiln", r"field work",
        r"\bthe poor\b", r"\bdebt\b",
        # Extended
        r"\bafford\b", r"school fee", r"\bfees?\b", r"\bexpenses?\b", r"\bhunger\b",
        r"\bstarv", r"\blivelihood\b", r"daily wage", r"seasonal work", r"\bharvest\b",
        r"agricultural work", r"breadwinner", r"\bration\b", r"\bbpl\b",
        r"below poverty", r"\bcost\b.*school", r"school.*\bcost\b", r"poor family",
        r"no money for", r"\bearn\b", r"child work", r"working child",
        r"\bkiln\b", r"\bmine\b", r"\bfactory\b", r"shop work", r"domestic worker",
        r"\bmaid\b", r"\bservant\b", r"cannot afford", r"unable to pay",
        r"money.*problem", r"problem.*money", r"financial.*problem",
        r"not enough money", r"scarce.*money", r"poor.*condition",
    ],
    "Child Marriage Issue": [
        r"child marriage", r"early marriage", r"marr(y|ied|iage).*18", r"under\s*18", r"marry.*young",
        # Extended
        r"\bdowry\b", r"marriage of girl", r"girl.*married", r"married.*off",
        r"forced marriage", r"marry before", r"\bnikaah\b", r"\bshaadi\b",
        r"\bvivah\b", r"\bdoli\b", r"\bengaged\b", r"\bbetrothed\b",
        r"beti ki shadi", r"get married", r"she was married", r"marriage pressure",
        r"marry at \d", r"age of marriage", r"marriage instead", r"dropped.*marriage",
        r"marriage.*dropout", r"dropout.*marriage", r"married.*young",
        r"girl.*marriage", r"marriage.*girl", r"underage.*marr",
    ],
    "Distance and Accessibility Issues": [
        r"\bdistance\b", r"far away", r"far from school", r"transport", r"bicycle", r"road",
        r"access", r"commut", r"nearby village", r"anganwadi.*far", r"far.*anganwadi",
        # Extended
        r"no nearby.*school", r"school.*another village", r"adjacent village",
        r"walk.*school", r"long walk", r"\bpath\b", r"\broute\b", r"\bbridge\b",
        r"\briver\b", r"\bflood\b", r"remote area", r"\bisolated\b",
        r"muddy road", r"dirt road", r"\bkuchcha\b", r"no bus", r"no vehicle",
        r"\brain\b.*school", r"\bweather\b", r"reach school", r"travel.*school",
        r"\bjungle\b", r"\bforest\b", r"going.*alone", r"have to cross",
        r"no footpath", r"no safe road", r"school.*far", r"far.*school",
        r"village.*school", r"school.*village", r"another.*village",
        r"cross.*road", r"cross.*river", r"difficult.*reach",
    ],
    "Parental Attitudes & Socio-Cultural": [
        r"parent", r"not aware", r"awareness", r"mindset", r"discrimination", r"household chore",
        r"household work", r"domestic work", r"girls? .*not allowed", r"importance of education",
        r"socio", r"do not send", r"not send.*school", r"going out.*house", r"prevent.*school",
        r"not feel.*stud", r"not attend.*school",
        # Extended
        r"gender bias", r"girl child.*educ", r"women.*educ", r"illiterate parent",
        r"parents.*not educated", r"do not value", r"not important.*stud",
        r"waste of time", r"will get married", r"girl.*stay home", r"\bpurdah\b",
        r"\bghunghat\b", r"\btradition\b", r"\bcustom\b", r"\bcaste\b",
        r"\btribe\b", r"\bbackward\b", r"social pressure", r"family pressure",
        r"\bneighbor\b", r"community pressure", r"son preference",
        r"brother.*school", r"son.*school", r"\borphan\b", r"step mother",
        r"in.?law", r"\bpatriarchal\b", r"not allowed.*go", r"\brestrict\b",
        r"stop.*school", r"keep.*home", r"does not allow", r"girl.*not.*school",
        r"not send.*girl", r"girl.*dropout", r"socio.?cultural",
        r"social.*attitude", r"attitude.*education", r"believe.*education",
    ],
    "School Infrastructure & Facility": [
        r"toilet", r"drinking water", r"mid[- ]?day meal", r"uniform", r"books?", r"anganwadi",
        r"infrastructure", r"facility", r"school building", r"no school", r"no provision",
        r"higher secondary",
        # Extended
        r"\bwashroom\b", r"\bsanitation\b", r"\bclassroom\b", r"\bfurniture\b",
        r"\bbench\b", r"\bstationery\b", r"\bnotebook\b", r"\blibrary\b",
        r"\belectricity\b", r"\bfan\b", r"\bcomputer\b", r"\blab\b",
        r"\bplayground\b", r"\bkitchen\b", r"\bcook\b", r"\bmeal\b",
        r"\blunch\b", r"\bfood\b.*school", r"\bnutrition\b", r"government scheme",
        r"\bkasturba\b", r"\bkgbv\b", r"residential school", r"\bhostel\b",
        r"\bboarding\b", r"samagra shiksha", r"school closed", r"school not open",
        r"school not functional", r"school building broken", r"leaking roof",
        r"no room", r"\binadequate\b", r"\binsufficient\b",
        r"primary school", r"secondary school", r"high school", r"upper primary",
        r"\benrolment\b", r"school admission", r"school.*enrollment",
        r"mdm\b", r"mid day", r"scholarship scheme", r"school supply",
        r"school material", r"school resource",
    ],
    "Teacher Capacity & Quality": [
        r"teacher", r"does not teach", r"shortage of teachers", r"teacher absentee", r"teaching quality",
        # Extended
        r"teacher absent", r"no teacher", r"teacher vacancy", r"single teacher",
        r"one teacher", r"para teacher", r"shiksha mitra", r"\buntrained\b",
        r"unqualified teacher", r"teacher.*not come", r"teacher.*late",
        r"irregular teacher", r"teacher transfer", r"teacher doesn.t",
        r"private tuition", r"\btutor\b", r"\bcoaching\b", r"poor learning",
        r"students not learning", r"no teaching", r"teacher.*behavior",
        r"teacher.*attitude", r"teacher.*treatment", r"teacher.*irregular",
        r"teacher.*absent", r"rarely.*teach", r"seldom.*teach",
        r"teacher.*quality", r"quality.*teach", r"teach.*quality",
    ],
    "Safety Issues": [
        r"harass", r"unsafe", r"safety", r"fear", r"violence", r"molest", r"kidney",
        r"go.*alone", r"walk.*alone", r"travel.*alone",
        # Extended
        r"stray dog", r"dog attack", r"dog bite", r"wild animal", r"\bsnake\b",
        r"eve teasing", r"\bteasing\b", r"\bbullying\b", r"\babuse\b",
        r"\bthreat\b", r"\bdanger\b", r"dangerous road", r"dark road",
        r"no light", r"\bkidnap\b", r"\babduction\b", r"\btheft\b",
        r"\bcrime\b", r"\bsecurity\b", r"open area", r"\bexposed\b",
        r"night.*school", r"\bunaccompanied\b", r"alone.*school",
        r"school.*alone", r"unsafe.*route", r"route.*unsafe",
        r"dangerous.*path", r"path.*dangerous", r"fear.*go.*school",
        r"scared.*school", r"attack", r"physical.*harm",
    ],
    "Substance Abuse & Addiction": [
        r"alcohol", r"drug", r"addiction", r"gambl", r"drunk", r"mobile addiction",
        # Extended
        r"\btobacco\b", r"\bbidi\b", r"\bcigarette\b", r"\bsmoking\b",
        r"\bliquor\b", r"country liquor", r"desi daru", r"\bintoxicant\b",
        r"\bsubstance\b.*abuse", r"\bnarcotics?\b", r"\bgutkha\b",
        r"pan masala", r"chewing.*tobacco", r"father.*drink", r"husband.*drink",
        r"domestic violence", r"family.*drunk", r"addiction.*family",
        r"drinks at home", r"drink.*home", r"home.*drink",
        r"alcohol.*family", r"family.*alcohol", r"parent.*drink",
        r"drink.*parent", r"intox", r"substance abuse",
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


class RetryableValidationError(Exception):
    """Base class for LLM output validation errors that should be retried/split."""


class LLMEmptyBatchError(RetryableValidationError):
    """Raised when model output parses into zero usable rows."""


class LLMLowCoverageError(RetryableValidationError):
    """Raised when parsed ID coverage is below acceptable threshold."""


class LLMLowQualityError(RetryableValidationError):
    """Raised when merged concept quality is below threshold."""


def is_retryable_error(error):
    if isinstance(error, RetryableValidationError):
        return True

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


def normalize_concept_for_similarity(text):
    if pd.isna(text):
        return ""

    cleaned = clean_merged_concept(text)
    cleaned = unicodedata.normalize("NFKC", cleaned)
    cleaned = cleaned.lower()
    cleaned = re.sub(r"[^a-z0-9\s]", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def tokenize_concept(text):
    key = normalize_concept_for_similarity(text)
    if not key:
        return set()

    stopwords = {
        'a', 'an', 'the', 'and', 'or', 'of', 'to', 'for', 'in', 'on', 'at', 'by',
        'with', 'from', 'is', 'are', 'was', 'were', 'be', 'being', 'this', 'that',
        'about', 'into', 'through', 'across', 'under', 'over'
    }

    tokens = []
    for token in key.split():
        if not token or token in stopwords:
            continue
        if token.endswith('ation') and len(token) > 7:
            token = token[:-5]
        elif token.endswith('tion') and len(token) > 6:
            token = token[:-4]
        if token.endswith('ing') and len(token) > 5:
            token = token[:-3]
        elif token.endswith('ed') and len(token) > 4:
            token = token[:-2]
        elif token.endswith('es') and len(token) > 4:
            token = token[:-2]
        elif token.endswith('s') and len(token) > 3:
            token = token[:-1]
        if token and token not in stopwords:
            tokens.append(token)

    return set(tokens)


def _ensure_theme_bank_bucket(concept_bank, theme):
    if theme not in concept_bank:
        concept_bank[theme] = {
            "concepts": set(),
            "concept_keys": {},
            "token_index": {},
        }
    return concept_bank[theme]


def add_concept_to_bank(concept_bank, theme, concept):
    if not theme:
        return

    canonical = clean_merged_concept(concept)
    if not canonical or canonical == "Uncategorized":
        return

    bucket = _ensure_theme_bank_bucket(concept_bank, theme)
    if canonical in bucket["concepts"]:
        return

    key = normalize_concept_for_similarity(canonical)
    if not key:
        return

    bucket["concepts"].add(canonical)
    bucket["concept_keys"][canonical] = key
    for token in tokenize_concept(canonical):
        if token not in bucket["token_index"]:
            bucket["token_index"][token] = set()
        bucket["token_index"][token].add(canonical)


def build_theme_concept_bank(reference_df):
    concept_bank = {}
    if reference_df.empty:
        return concept_bank

    for _, row in reference_df.iterrows():
        theme = normalize_theme_name(row.get("Theme"))
        concept = row.get("Merged_Concept", "")
        add_concept_to_bank(concept_bank, theme, concept)

    return concept_bank


def concept_similarity_score(concept_a, concept_b):
    key_a = normalize_concept_for_similarity(concept_a)
    key_b = normalize_concept_for_similarity(concept_b)
    if not key_a or not key_b:
        return 0.0

    if key_a == key_b:
        return 1.0

    tokens_a = tokenize_concept(key_a)
    tokens_b = tokenize_concept(key_b)
    if tokens_a and tokens_b:
        overlap = len(tokens_a & tokens_b) / min(len(tokens_a), len(tokens_b))
        jaccard = len(tokens_a & tokens_b) / len(tokens_a | tokens_b)
    else:
        overlap = 0.0
        jaccard = 0.0

    seq_ratio = SequenceMatcher(None, key_a, key_b).ratio()
    token_order_a = " ".join(sorted(tokens_a))
    token_order_b = " ".join(sorted(tokens_b))
    token_seq = SequenceMatcher(None, token_order_a, token_order_b).ratio() if token_order_a and token_order_b else 0.0

    return max(
        seq_ratio,
        overlap,
        token_seq,
        (0.55 * overlap) + (0.25 * token_seq) + (0.20 * jaccard),
    )


def find_similar_concept_in_bank(concept, theme, concept_bank, min_score=CONCEPT_MERGE_MIN_SCORE):
    if not concept_bank or not theme:
        return None

    bucket = concept_bank.get(theme)
    if not bucket or not bucket.get("concepts"):
        return None

    candidate_tokens = tokenize_concept(concept)
    candidates = set()
    for token in candidate_tokens:
        candidates.update(bucket["token_index"].get(token, set()))

    if not candidates:
        return None

    best_concept = None
    best_score = 0.0
    for existing in candidates:
        score = concept_similarity_score(concept, existing)
        if score > best_score:
            best_score = score
            best_concept = existing

    if best_concept and best_score >= min_score:
        return best_concept
    return None


def reconcile_batch_concepts_with_memory(df_batch, concept_bank):
    if df_batch.empty or not concept_bank:
        return df_batch

    updated = df_batch.copy()
    for index in updated.index:
        theme = normalize_theme_name(updated.at[index, "Theme"])
        concept = clean_merged_concept(updated.at[index, "Merged_Concept"])
        if not theme or not concept or concept == "Uncategorized":
            continue

        matched = find_similar_concept_in_bank(concept, theme, concept_bank)
        if matched:
            updated.at[index, "Merged_Concept"] = matched

    return updated


def build_concept_memory_prompt(concept_bank, max_per_theme=10, max_total_chars=1800):
    if not concept_bank:
        return ""

    sections = []
    total_chars = 0
    for theme in CANONICAL_THEMES:
        bucket = concept_bank.get(theme)
        if not bucket:
            continue

        concepts = sorted(bucket["concepts"], key=lambda value: len(value))[:max_per_theme]
        if not concepts:
            continue

        section = f"- {theme}: " + "; ".join(concepts)
        if total_chars + len(section) > max_total_chars:
            break
        sections.append(section)
        total_chars += len(section)

    if not sections:
        return ""

    return (
        "CANONICAL CONCEPT MEMORY (REUSE IF SAME MEANING):\n"
        + "\n".join(sections)
        + "\nReuse exact concept text when equivalent; create new only if truly new."
    )


def is_high_quality_merged_concept(concept_text, type_label):
    """Validate concept quality without replacing with keyword-based headings.

    This is a quality gate only; concept generation remains LLM-driven.
    """
    if pd.isna(concept_text):
        return False

    concept = clean_merged_concept(concept_text)
    if not concept or concept == "Uncategorized":
        return False

    words = concept.split()
    lower = concept.lower()

    if len(words) < 2 or len(words) > 12:
        return False

    if any(char in concept for char in ['"', '“', '”', '?', '!']):
        return False

    noisy_prefixes = (
        'we ', 'i ', 'they ', 'when ', 'all the ', 'to resolve ',
        'the didi ', 'this is ', 'that is '
    )
    if lower.startswith(noisy_prefixes):
        return False

    if str(type_label).strip().lower() == 'solution':
        conversational_markers = [
            'didi', 'we told', 'we organised', 'oath', 'there we',
            'ham ', 'hum ', 'mera ', 'hamne ', 'our village'
        ]
        if any(marker in lower for marker in conversational_markers):
            return False

    return True


def validate_llm_batch_quality(df_batch, text_batch, type_label):
    """Raise on poor LLM formatting so caller can retry same batch."""
    if df_batch.empty:
        raise LLMEmptyBatchError("LLM returned empty batch")

    expected = len(text_batch)

    def _min_coverage_threshold(size):
        if size >= 150:
            return 0.86
        if size >= 100:
            return 0.90
        if size >= 60:
            return 0.93
        if size >= 30:
            return 0.96
        return 0.98

    mapped = df_batch['Original'].nunique()
    coverage_ratio = mapped / expected if expected else 1.0
    min_coverage = _min_coverage_threshold(expected)
    if coverage_ratio < min_coverage:
        raise LLMLowCoverageError(
            f"Low LLM coverage: {mapped}/{expected} ({coverage_ratio:.2%} < {min_coverage:.0%})"
        )

    quality_ratio = df_batch['Merged_Concept'].apply(
        lambda value: is_high_quality_merged_concept(value, type_label)
    ).mean()
    min_quality = 0.88 if str(type_label).strip().lower() == 'solution' else 0.80
    if quality_ratio < min_quality:
        raise LLMLowQualityError(
            f"Low concept quality ratio: {quality_ratio:.2f} < {min_quality:.2f}"
        )


def run_llm_chunk_with_retries(text_batch, type_label, concept_bank, batch_label):
    """Run one chunk with bounded retries for transient + validation errors."""
    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            return get_ai_mapping(text_batch, type_label, concept_bank=concept_bank)
        except Exception as error:
            last_error = error
            retryable = is_retryable_error(error)
            should_retry = retryable and attempt < MAX_RETRIES
            print(f"      ⚠️ Batch {batch_label} attempt {attempt}/{MAX_RETRIES} failed: {error}")
            if not should_retry:
                break

            delay = min(MAX_RETRY_SECONDS, BASE_RETRY_SECONDS * (2 ** (attempt - 1)))
            delay += random.uniform(0, 0.5)
            print(f"      🔁 Retrying batch {batch_label} in {delay:.1f}s...")
            time.sleep(delay)

    raise last_error if last_error else RuntimeError(f"Batch {batch_label} failed unexpectedly")


def map_with_adaptive_fallback(text_batch, type_label, concept_bank, batch_label, depth=0):
    """Map a batch with adaptive split fallback for low-coverage/quality failures."""
    if not text_batch:
        return pd.DataFrame(columns=["Original", "Theme", "Merged_Concept", "Confidence", "Mapping_Source"])

    try:
        mapped_df = run_llm_chunk_with_retries(text_batch, type_label, concept_bank, batch_label)
        return mapped_df.drop_duplicates(subset=["Original"], keep="last")
    except Exception as error:
        can_split = (
            len(text_batch) > ADAPTIVE_MIN_BATCH_SIZE
            and depth < ADAPTIVE_MAX_SPLIT_DEPTH
            and is_retryable_error(error)
        )
        if not can_split:
            print(f"      ❌ Batch {batch_label} failed after retries: {error}")
            return pd.DataFrame(columns=["Original", "Theme", "Merged_Concept", "Confidence", "Mapping_Source"])

        midpoint = len(text_batch) // 2
        left_items = text_batch[:midpoint]
        right_items = text_batch[midpoint:]
        print(
            f"      🧩 Splitting batch {batch_label} at depth {depth + 1} "
            f"({len(text_batch)} -> {len(left_items)} + {len(right_items)})"
        )

        left_df = map_with_adaptive_fallback(
            left_items,
            type_label,
            concept_bank,
            batch_label=f"{batch_label}.L",
            depth=depth + 1,
        )
        right_df = map_with_adaptive_fallback(
            right_items,
            type_label,
            concept_bank,
            batch_label=f"{batch_label}.R",
            depth=depth + 1,
        )

        combined = pd.concat([left_df, right_df], ignore_index=True)
        return combined.drop_duplicates(subset=["Original"], keep="last")


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

def get_ai_mapping(text_batch, type_label, concept_bank=None):
    if not llm_provider:
        raise RuntimeError("LLM provider is not configured")

    indexed_rows = "\n".join([f"{index + 1}\t{text}" for index, text in enumerate(text_batch)])
    concept_memory_block = build_concept_memory_prompt(concept_bank)

    solution_quality_block = """
SOLUTION QUALITY:
- Write a professional intervention heading, not narration.
- No direct speech, no person names/titles, no story-like sentence.
- Prefer action-oriented 3-8 word phrasing.
""" if str(type_label).strip().lower() == "solution" else ""

    prompt_content = f"""You are classifying education {type_label} statements.

THEMES (use exactly one):
{THEME_KNOWLEDGE_BASE}

PRIMARY OBJECTIVE:
For each input line, output ONE Theme and ONE canonical Merged_Concept.

DEDUP RULES (MANDATORY):
1) Merge semantically equivalent lines to the same exact Merged_Concept wording.
2) Reuse the same label for paraphrases; do not create near-duplicate labels.
3) Merged_Concept must be short, stable, and report-ready (2-10 words).
4) Avoid sentence-like wording, local anecdotes, or person-specific references.
5) Do not use '+' or multiple themes.

{concept_memory_block}
{solution_quality_block}

OUTPUT CONTRACT (STRICT):
- Return plain text only, no markdown, no explanation.
- Exactly {len(text_batch)} lines, one per input ID.
- IDs must be 1..{len(text_batch)} and each ID appears exactly once.
- Format each line exactly: ID|Theme|Merged_Concept
- Merged_Concept must not contain '|'.
- Keep rows ordered by ID ascending.

SELF-CHECK BEFORE FINALIZING:
- Confirm all IDs 1..{len(text_batch)} are present exactly once.
- Confirm all themes are from the allowed list.
- Confirm equivalent meanings share identical Merged_Concept text.

INPUT:
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
    df_batch = postprocess_mapping_batch(df_batch)
    df_batch = reconcile_batch_concepts_with_memory(df_batch, concept_bank)
    validate_llm_batch_quality(df_batch, text_batch, type_label)
    return df_batch


def enforce_batch_coverage(current_items, mapped_df, reference_lookup, concept_bank):
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

        matched_concept = find_similar_concept_in_bank(concept, theme, concept_bank)
        if matched_concept:
            concept = matched_concept

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
    filled_df = reconcile_batch_concepts_with_memory(filled_df, concept_bank)
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
    concept_bank = {}
    if os.path.exists(output_csv):
        try:
            existing_output = pd.read_csv(output_csv)
            if 'Original' in existing_output.columns:
                existing_clean = postprocess_mapping_batch(existing_output)
                already_processed = set(existing_output['Original'].dropna().astype(str).tolist())
                reference_lookup = build_reference_lookup(existing_clean)
                concept_bank = build_theme_concept_bank(existing_clean)
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
        mapped_df = map_with_adaptive_fallback(
            current_items,
            type_label,
            concept_bank,
            batch_label=str(current_batch),
        )

        llm_coverage = (mapped_df['Original'].nunique() / len(current_items)) if len(current_items) else 1.0
        print(f"      📊 LLM mapped {mapped_df['Original'].nunique()}/{len(current_items)} rows ({llm_coverage:.1%}) before final coverage enforcement")

        mapped_df = enforce_batch_coverage(current_items, mapped_df, reference_lookup, concept_bank)

        if not mapped_df.empty:
            save_progress(output_csv, mapped_df)
            for _, row in mapped_df.iterrows():
                orig = str(row.get("Original", "")).strip()
                norm = normalize_for_match(orig)
                theme = normalize_theme_name(row.get("Theme"))
                concept = row.get("Merged_Concept", "")
                conf = row.get("Confidence", 0.9)
                if norm and theme:
                    clean_concept = clean_merged_concept(str(concept))
                    reference_lookup[norm] = {
                        "Theme": theme,
                        "Merged_Concept": clean_concept,
                        "Confidence": float(conf) if str(conf).strip() else 0.9,
                        "Mapping_Source": "exact_reference",
                    }
                    add_concept_to_bank(concept_bank, theme, clean_concept)
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
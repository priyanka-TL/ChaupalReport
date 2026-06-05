import pandas as pd
import json
import re

# Education context keywords — texts with NONE of these are likely off-topic
_EDU_RE = re.compile(
    r'\b(school|education|student|child|girl|boy|study|class|grade|learn|dropout|'
    r'admission|teacher|attend|enroll|literacy|read|write|book|uniform|fee|scholarship|'
    r'marriage|labour|labor|poverty|document|aadhaar|certificate|distance|transport)\b',
    re.IGNORECASE,
)
# Patterns that strongly indicate garbled/irrelevant translation with no education link
_NOISE_RE = re.compile(
    r'\bcurd\b|\bghee\b|\bkhana\b|\bbirthday.*created\b|\bcreated.*birthday\b',
    re.IGNORECASE,
)

def _is_noise_text(text):
    """Return True for texts that are clearly noise — too short or garbled translations."""
    t = str(text).strip()
    words = t.split()
    if len(words) < 3:
        return True
    # Garbled food/personal metaphors with zero education context
    if _NOISE_RE.search(t) and not _EDU_RE.search(t):
        return True
    return False


def run_phase1(input_file):
    print("🚀 Starting Phase 1: Data Cleaning & Explosion...")
    df = pd.read_csv(input_file)
    df = df.loc[:, ~df.columns.str.match(r'^Unnamed')]
    TOTAL_CHAUPALS = len(df)

    # --- PART A: STANDARDIZE COUNTS (Same as before) ---
    def parse_counts(row):
        total, men, women, children = 0, 0, 0, 0
        pc_value = str(row['Participant Count']).strip()
        if pc_value.startswith('{'):
            try:
                data = json.loads(pc_value.replace("'", '"'))
                total = int(data.get('total') or 0)
                men = int(data.get('men') or 0)
                women = int(data.get('women') or 0)
                children = int(data.get('children') or 0)
            except: pass
        elif pc_value.replace('.','',1).isdigit():
            total = int(float(pc_value))
        
        # Fallback to individual columns
        if total == 0:
            def _to_int(v):
                try:
                    return int(float(v)) if pd.notna(v) else 0
                except (ValueError, TypeError):
                    return 0
            m = _to_int(row['Men'])
            w = _to_int(row['Women'])
            c = _to_int(row['Children'])
            total = m + w + c
        return pd.Series([total, men, women, children])

    df[['Participant Count', 'Men', 'Women', 'Children']] = df.apply(parse_counts, axis=1)

    # --- PART B: EXPLODE CHALLENGES & SOLUTIONS ---
    def explode_col(dataframe, col_name):
        temp_df = dataframe.copy()
        temp_df[col_name] = temp_df[col_name].fillna("None").str.split('|')
        exploded = temp_df.explode(col_name)
        exploded[col_name] = exploded[col_name].str.strip().apply(lambda x: re.sub(r'^\d+\.\s*', '', str(x)))
        exploded = exploded[exploded[col_name].str.len() > 2]
        exploded = exploded.loc[:, ~exploded.columns.str.match(r'^Unnamed')]
        return exploded

    df_chal = explode_col(df, 'Challenges')
    df_sol = explode_col(df, 'Solutions')

    # --- PART C: EXPORT UNIQUE LISTS FOR AI ---
    all_chal_texts = df_chal['Challenges'].unique()
    all_sol_texts  = df_sol['Solutions'].unique()

    # Drop obvious noise (garbled translations, too-short fragments)
    clean_chal = [t for t in all_chal_texts if not _is_noise_text(t)]
    clean_sol  = [t for t in all_sol_texts  if not _is_noise_text(t)]
    noise_c = len(all_chal_texts) - len(clean_chal)
    noise_s = len(all_sol_texts)  - len(clean_sol)
    if noise_c or noise_s:
        print(f"   🧹 Noise filter removed {noise_c} challenge texts, {noise_s} solution texts before AI tagging.")

    pd.DataFrame(clean_chal, columns=['text']).to_csv('unique_challenges.csv', index=False)
    pd.DataFrame(clean_sol,  columns=['text']).to_csv('unique_solutions.csv', index=False)
    
    # Save the exploded masters for the final merger
    df_chal.to_csv('exploded_challenges.csv', index=False)
    df_sol.to_csv('exploded_solutions.csv', index=False)

    print(f"✅ Phase 1 Complete. Unique Challenges (clean): {len(clean_chal)}, Unique Solutions (clean): {len(clean_sol)}")

if __name__ == "__main__":
    run_phase1('cleaned_data.csv')
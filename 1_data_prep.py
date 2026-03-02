import pandas as pd
import json
import re

def run_phase1(input_file):
    print("🚀 Starting Phase 1: Data Cleaning & Explosion...")
    df = pd.read_csv(input_file)
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
            m = int(float(row['Men'])) if pd.notna(row['Men']) else 0
            w = int(float(row['Women'])) if pd.notna(row['Women']) else 0
            c = int(float(row['Children'])) if pd.notna(row['Children']) else 0
            total = m + w + c
        return pd.Series([total, men, women, children])

    df[['Participant Count', 'Men', 'Women', 'Children']] = df.apply(parse_counts, axis=1)

    # --- PART B: EXPLODE CHALLENGES & SOLUTIONS ---
    def explode_col(dataframe, col_name):
        temp_df = dataframe.copy()
        temp_df[col_name] = temp_df[col_name].fillna("None").str.split('|')
        exploded = temp_df.explode(col_name)
        exploded[col_name] = exploded[col_name].str.strip().apply(lambda x: re.sub(r'^\d+\.\s*', '', str(x)))
        return exploded[exploded[col_name].str.len() > 2]

    df_chal = explode_col(df, 'Challenges')
    df_sol = explode_col(df, 'Solutions')

    # --- PART C: EXPORT UNIQUE LISTS FOR AI ---
    # Challenges
    pd.DataFrame(df_chal['Challenges'].unique(), columns=['text']).to_csv('unique_challenges.csv', index=False)
    # Solutions
    pd.DataFrame(df_sol['Solutions'].unique(), columns=['text']).to_csv('unique_solutions.csv', index=False)
    
    # Save the exploded masters for the final merger
    df_chal.to_csv('exploded_challenges.csv', index=False)
    df_sol.to_csv('exploded_solutions.csv', index=False)

    print(f"✅ Phase 1 Complete. Unique Challenges: {len(df_chal['Challenges'].unique())}, Unique Solutions: {len(df_sol['Solutions'].unique())}")

if __name__ == "__main__":
    run_phase1('cleaned_data.csv')
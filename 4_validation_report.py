import pandas as pd
import numpy as np
import os
import re

def clean_theme_name(text):
    """Cleans theme names, handling combined themes and empty values."""
    if pd.isna(text) or str(text).strip() == "" or str(text).lower() == "nan":
        return "Other Factors"
    text = str(text).strip()
    if '+' in text:
        text = text.split('+')[0].strip()
    text = re.sub(r'^\d+[\.\)\s-]*', '', text).strip()
    return text

def generate_validation_reports():
    print("🚀 Starting Validation Report Generation...")

    # 1. Load Data
    print("   📂 Loading datasets...")
    try:
        df_raw = pd.read_csv('cleaned_data.csv')
        chal_exploded = pd.read_csv('exploded_challenges.csv')
        sol_exploded = pd.read_csv('exploded_solutions.csv')
        chal_map = pd.read_csv('challenge_mapping.csv')
        sol_map = pd.read_csv('solution_mapping.csv')
    except Exception as e:
        print(f"❌ Error: Required CSV files missing. {e}")
        return

    # 2. Prepare Mappings
    print("   ⚙️  Preparing mappings...")
    chal_map['Theme'] = chal_map['Theme'].apply(clean_theme_name)
    sol_map['Theme'] = sol_map['Theme'].apply(clean_theme_name)

    # 3. Merge Mappings to Exploded Data
    # Challenges
    df_c = chal_exploded.merge(chal_map, left_on='Challenges', right_on='Original', how='left')
    df_c['Theme'] = df_c['Theme'].fillna("Other Factors")
    
    # Solutions
    df_s = sol_exploded.merge(sol_map, left_on='Solutions', right_on='Original', how='left')
    df_s['Theme'] = df_s['Theme'].fillna("Other Factors")

    # 4. Aggregate at Chaupal Level
    print("   📊 Aggregating data at Chaupal level...")
    
    # Group Challenges by ID (assuming 'id' in exploded matches 'id' in raw)
    # Note: Check if 'id' exists in exploded files. Usually exploded files inherit the ID.
    # Let's verify column names in a moment, but assuming standard structure:
    
    # Helper to aggregate themes and texts
    def agg_texts(series):
        return " | ".join([str(x) for x in series if pd.notna(x) and str(x) != ""])
    
    def agg_unique_themes(series):
        return ", ".join(sorted(list(set([str(x) for x in series if pd.notna(x) and str(x) != ""]))))

    # Challenges Aggregation
    chal_agg = df_c.groupby('id').agg({
        'Challenges': 'count',
        'Merged_Concept': agg_texts,
        'Theme': agg_unique_themes
    }).rename(columns={
        'Challenges': 'Challenge_Count',
        'Merged_Concept': 'All_Challenges_Listed',
        'Theme': 'Challenge_Themes_Identified'
    })

    # Solutions Aggregation
    sol_agg = df_s.groupby('id').agg({
        'Solutions': 'count',
        'Merged_Concept': agg_texts,
        'Theme': agg_unique_themes
    }).rename(columns={
        'Solutions': 'Solution_Count',
        'Merged_Concept': 'All_Solutions_Listed',
        'Theme': 'Solution_Themes_Identified'
    })

    # 5. Merge with Raw Chaupal Data
    print("   🔗 Merging with Chaupal demographics...")
    
    # Ensure numeric columns
    numeric_cols = ['Participant Count', 'Men', 'Women', 'Children']
    for col in numeric_cols:
        df_raw[col] = pd.to_numeric(df_raw[col], errors='coerce').fillna(0)

    # Calculate Percentages
    df_raw['Men_%'] = (df_raw['Men'] / df_raw['Participant Count'] * 100).round(1)
    df_raw['Women_%'] = (df_raw['Women'] / df_raw['Participant Count'] * 100).round(1)
    df_raw['Children_%'] = (df_raw['Children'] / df_raw['Participant Count'] * 100).round(1)

    # Merge
    master_df = df_raw.merge(chal_agg, on='id', how='left')
    master_df = master_df.merge(sol_agg, on='id', how='left')

    # Fill NaNs for counts
    master_df['Challenge_Count'] = master_df['Challenge_Count'].fillna(0).astype(int)
    master_df['Solution_Count'] = master_df['Solution_Count'].fillna(0).astype(int)
    master_df['All_Challenges_Listed'] = master_df['All_Challenges_Listed'].fillna("")
    master_df['Challenge_Themes_Identified'] = master_df['Challenge_Themes_Identified'].fillna("None")
    master_df['All_Solutions_Listed'] = master_df['All_Solutions_Listed'].fillna("")
    master_df['Solution_Themes_Identified'] = master_df['Solution_Themes_Identified'].fillna("None")

    # 6. Export
    output_file = 'Chaupal_Validation_Report.csv'
    print(f"   💾 Saving {output_file}...")
    
    # Select and Reorder columns for readability
    cols = [
        'id', 'District', 'Block', 'Village', 
        'Participant Count', 'Men', 'Women', 'Children',
        'Men_%', 'Women_%', 'Children_%',
        'Challenge_Count', 'Challenge_Themes_Identified', 'All_Challenges_Listed',
        'Solution_Count', 'Solution_Themes_Identified', 'All_Solutions_Listed'
    ]
    
    # Only keep columns that exist
    final_cols = [c for c in cols if c in master_df.columns]
    
    master_df[final_cols].to_csv(output_file, index=False)
    
    print("\n✅ Validation Report Generated Successfully!")
    print(f"   File: {output_file}")
    print("   Columns included: ID, Location, Demographics (Counts & %), Themes, and Full Text.")

if __name__ == "__main__":
    generate_validation_reports()

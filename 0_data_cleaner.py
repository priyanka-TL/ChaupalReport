import pandas as pd
import json
import numpy as np

def clean_participant_data(file_path, output_path):
    df = pd.read_csv(file_path)
    
    def parse_counts(row):
        # Initialize default values
        total, men, women, children = 0, 0, 0, 0
        
        # 1. Try to get values from existing individual columns first
        try:
            men = int(float(row['Men'])) if pd.notna(row['Men']) else 0
            women = int(float(row['Women'])) if pd.notna(row['Women']) else 0
            children = int(float(row['Children'])) if pd.notna(row['Children']) else 0
        except:
            pass

        pc_value = str(row['Participant Count']).strip()

        # 2. Check if Participant Count is a JSON string
        if pc_value.startswith('{'):
            try:
                # Fix common JSON formatting issues in CSVs
                json_str = pc_value.replace("'", '"')
                data = json.loads(json_str)
                
                # Update values if they exist in JSON (handling empty strings "")
                total = int(data.get('total')) if data.get('total') not in ["", None] else total
                men = int(data.get('men')) if data.get('men') not in ["", None] else men
                women = int(data.get('women')) if data.get('women') not in ["", None] else women
                children = int(data.get('children')) if data.get('children') not in ["", None] else children
            except:
                pass
        
        # 3. If Participant Count is just a plain number
        elif pc_value.replace('.','',1).isdigit():
            total = int(float(pc_value))

        # 4. Final Logic: If total is 0 but components exist, sum them up
        if total == 0 or total < (men + women + children):
            total = men + women + children
            
        return pd.Series([total, men, women, children])

    # Apply the cleaning logic
    df[['Participant Count', 'Men', 'Women', 'Children']] = df.apply(parse_counts, axis=1)
    
    # Save the cleaned data
    df.to_csv(output_path, index=False)
    print(f"✅ Data cleaned and saved to: {output_path}")
    print(f"Sample Totals: {df['Participant Count'].head().tolist()}")

if __name__ == "__main__":
    # Input: your raw file | Output: the file for Step 1
    clean_participant_data('raw_data.csv', 'cleaned_data.csv')
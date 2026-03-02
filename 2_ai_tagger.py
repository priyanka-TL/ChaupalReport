import pandas as pd
import os
import time
import io
import json
import boto3
from tqdm import tqdm
from dotenv import load_dotenv

load_dotenv()

# --- YOUR SPECIFIC AWS CONFIGURATION ---
claude_beadrock_client = boto3.client(
    "bedrock-runtime",
    region_name="ap-south-1",  # As per your config
    aws_access_key_id=os.getenv("AWS_ACCESS_KEY_ID"),
    aws_secret_access_key=os.getenv("AWS_SECRET_ACCESS_KEY"),
)


MODEL_ID = "global.anthropic.claude-sonnet-4-5-20250929-v1:0"
MODEL_VERSION = "bedrock-2023-05-31"

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

def get_ai_mapping_bedrock(text_batch, type_label):
    prompt_content = f"""Act as an expert Social Data Analyst. Use these THEMES:
    {THEME_KNOWLEDGE_BASE}
    
    SEMANTIC DEDUPLICATION PROTOCOL (MANDATORY):
    You must merge semantically similar items into a single "Merged_Concept".
    
    Rules:
    1. Group all variants expressing the same core issue.
    2. Select the most complete and descriptive version as the canonical form for 'Merged_Concept'.
    3. Examples of merging:
       - "Due to poverty" = "Due to poor financial condition" = "Lack of money" -> Merged_Concept: "Poverty preventing education"
       - "No Aadhaar card" = "Lack of Aadhaar" = "Aadhaar not made" -> Merged_Concept: "Lack of legal documentation (Aadhaar)"
       - "School is far" = "School is very far" = "Distance of school" -> Merged_Concept: "School distance and accessibility issues"
    4. CRITICAL: Assign EXACTLY ONE theme from the list. Do not combine themes with '+' or 'and'. If multiple apply, choose the most dominant one.
    
    TASK: Categorize these unique {type_label} statements.
    OUTPUT: Return ONLY a CSV-style format with three columns: Original|Theme|Merged_Concept
    Use the | character as the delimiter. Do not include headers, preamble, or markdown backticks.
    
    DATA:
    {text_batch}"""

    native_request = {
        "anthropic_version": MODEL_VERSION,
        "max_tokens": 4000,
        "temperature": 0,
        "messages": [
            {
                "role": "user",
                "content": [{"type": "text", "text": prompt_content}]
            }
        ]
    }

    try:
        response = claude_beadrock_client.invoke_model(
            modelId=MODEL_ID,
            body=json.dumps(native_request)
        )
        response_body = json.loads(response.get('body').read())
        raw_output = response_body['content'][0]['text'].strip()
        
        # Strip potential garbage
        raw_output = raw_output.replace('```csv', '').replace('```', '').strip()
        
        # Load into DF (Expects: Original|Theme|Merged_Concept)
        df_batch = pd.read_csv(io.StringIO(raw_output), sep='|', names=['Original', 'Theme', 'Merged_Concept'], header=None)
        return df_batch
    except Exception as e:
        print(f"Error in batch: {e}")
        return pd.DataFrame()

def process_file(input_csv, output_csv, type_label):
    if not os.path.exists(input_csv):
        print(f"File {input_csv} not found. Skipping.")
        return

    df_unique = pd.read_csv(input_csv)
    unique_list = df_unique['text'].dropna().unique().tolist()
    
    final_dfs = []
    batch_size = 50  # Set to 50 to avoid output token limits with large datasets
    
    total_batches = (len(unique_list) + batch_size - 1) // batch_size
    print(f"🔍 Analyzing {len(unique_list)} Unique {type_label}s via Claude 3.7 (ap-south-1)...")
    print(f"   Total Batches: {total_batches} | Batch Size: {batch_size}")

    for i in tqdm(range(0, len(unique_list), batch_size)):
        current_batch = (i // batch_size) + 1
        print(f"   ⏳ Processing Batch {current_batch}/{total_batches}...")
        
        batch = "\n".join(unique_list[i : i + batch_size])
        mapped_df = get_ai_mapping_bedrock(batch, type_label)
        if not mapped_df.empty:
            final_dfs.append(mapped_df)
            print(f"      ✅ Batch {current_batch} done. Got {len(mapped_df)} items.")
        else:
            print(f"      ⚠️ Batch {current_batch} returned empty or failed.")
            
        time.sleep(0.5) 
        
    if final_dfs:
        result_df = pd.concat(final_dfs, ignore_index=True)
        result_df.to_csv(output_csv, index=False)
        print(f"✅ Mapping successfully saved to {output_csv}")

if __name__ == "__main__":
    # Ensure these files exist from Phase 1
    process_file('unique_challenges.csv', 'challenge_mapping.csv', 'Challenge')
    process_file('unique_solutions.csv', 'solution_mapping.csv', 'Solution')
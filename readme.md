# Chaupal Report Generation Pipeline

This project generates two outputs from raw Shiksha Chaupal data:

- A narrative Word report: `Final_Shiksha_Report.docx`
- A validation CSV report: `Chaupal_Validation_Report.csv`

The pipeline cleans participant demographics, explodes text responses, uses AI for theme mapping and semantic deduplication, then builds section-wise analytics and insights.

Current branch for the chaupal6: Chaupal6WithOpenRouter

## Project Flow

1. **`0_data_cleaner.py`**
	- Cleans participant counts.
	- Resolves mixed formats in `Participant Count` (JSON-like strings or plain numbers).
	- Reconciles `Men`, `Women`, `Children` with total count.
	- Output: `cleaned_data.csv`

2. **`1_data_prep.py`**
	- Re-validates participant count fields.
	- Splits and explodes `Challenges` and `Solutions` by `|` into row-level records.
	- Creates unique text pools for AI mapping.
	- Outputs:
	  - `exploded_challenges.csv`
	  - `exploded_solutions.csv`
	  - `unique_challenges.csv`
	  - `unique_solutions.csv`

3. **`2_ai_tagger.py`**
	- Sends unique challenge/solution texts to Bedrock Claude in batches.
	- Returns exactly one theme per item + one canonical `Merged_Concept`.
	- Outputs:
	  - `challenge_mapping.csv`
	  - `solution_mapping.csv`

4. **`3_final_processor.py`**
	- Merges exploded data with AI mappings.
	- Applies:
	  - Challenge environment classification (`School`, `Home`, `Community`)
	  - Solution agency classification (`Individual-led`, `Community-led`, `Institutional`)
	- Runs optional AI refinement on top concepts for stronger deduplication/theme correction.
	- Computes statewide and district-level metrics.
	- Builds the final Word report with 7 sections.

5. **`4_validation_report.py`**
	- Creates a Chaupal-level validation table by aggregating mapped challenges/solutions.
	- Adds participant composition percentages.
	- Output: `Chaupal_Validation_Report.csv`

---

## Setup

1. Create virtual environment:

```bash
python3 -m venv sc_report_env
```

2. Activate environment:

```bash
source sc_report_env/bin/activate
```

3. Install dependencies:

```bash
pip install -r requirements.txt
```

If needed, install key libraries directly:

```bash
pip install pandas numpy python-docx boto3 tqdm python-dotenv
```

---

## AWS / AI Configuration

Set the following before running AI scripts:

- `LLM_PROVIDER` (`claude` or `gemini`)

If `LLM_PROVIDER=claude`:
- `AWS_ACCESS_KEY_ID`
- `AWS_SECRET_ACCESS_KEY`
- `AWS_REGION` (default: `ap-south-1`)
- `CLAUDE_MODEL_ID` (default: `global.anthropic.claude-sonnet-4-5-20250929-v1:0`)
- `CLAUDE_MODEL_VERSION` (default: `bedrock-2023-05-31`)

If `LLM_PROVIDER=gemini`:
- `GEMINI_API_KEY`
- `GEMINI_MODEL` (default: `gemini-2.0-flash`)

You can place secrets in a `.env` file (loaded via `python-dotenv`).

---

## Run Order

Run scripts in this exact order:

```bash
python 0_data_cleaner.py
python 1_data_prep.py
python 2_ai_tagger.py
python 3_final_processor.py
python 4_validation_report.py
```

---

## Report Structure (Word Report)

`3_final_processor.py` generates:

1. **Executive Summary**
2. **General Participation Overview**
3. **Core Content Analysis**
4. **Thematic Analysis**
5. **District Profiles**
6. **Unique Insights**
7. **Conclusion**

Each section is built from computed metrics and mapped text data (not static templates), so output changes with input data.

---

## Main Data Artifacts

- `cleaned_data.csv`
- `exploded_challenges.csv`
- `exploded_solutions.csv`
- `unique_challenges.csv`
- `unique_solutions.csv`
- `challenge_mapping.csv`
- `solution_mapping.csv`
- `Final_Shiksha_Report.docx`
- `Chaupal_Validation_Report.csv`
# Detailed Explanation: How the Chaupal Report Is Generated

This document explains exactly how data flows through the pipeline, how each report section is computed, and what AI prompts are used for tagging and refinement.

We run the report in 4 stages. 
First, we clean the raw discussion data so participant numbers are correct and consistent. 
Second, we split long responses into individual challenge and solution points so each idea is counted properly. 
Third, we use AI to group similar statements (for example, many versions of “no money” become one common issue) and tag each under one theme. 
Fourth, we calculate district and state-level insights (participation, top problems, common solutions, community-led vs system-led actions) and automatically build the final Word report section by section.

## 1) End-to-End Pipeline Logic

The project follows a deterministic sequence:

1. **Clean and standardize participant data** (`0_data_cleaner.py`)
2. **Explode challenge/solution text into analyzable rows** (`1_data_prep.py`)
3. **Use AI to assign theme + merged concept** (`2_ai_tagger.py`)
4. **Assemble narrative Word report with metrics and insights** (`3_final_processor.py`)
5. **Generate Chaupal-level validation CSV** (`4_validation_report.py`)

---

## 2) Script-by-Script Detail

### `0_data_cleaner.py` — Participant Standardization

Purpose: make `Participant Count`, `Men`, `Women`, `Children` internally consistent.

Core logic:
- Reads `raw_data.csv`.
- For each row:
	- Tries `Men/Women/Children` columns first.
	- If `Participant Count` is JSON-like, parses keys: `total`, `men`, `women`, `children`.
	- If `Participant Count` is numeric text, converts to integer.
	- If total is missing/zero or less than sum of components, replaces total with `men + women + children`.
- Writes cleaned output to `cleaned_data.csv`.

Why this matters: every downstream percentage and participation metric depends on these corrected counts.

---

### `1_data_prep.py` — Explosion + Unique Pools

Purpose: convert multi-response text into row-level records and prepare unique text lists for AI tagging.

Core logic:
- Reads `cleaned_data.csv`.
- Re-runs count normalization safeguard.
- Splits `Challenges` and `Solutions` on `|`.
- Explodes into separate rows (`explode_col`).
- Cleans numbering prefixes (`1.`, `2.` etc.).
- Filters trivial/short entries.
- Exports:
	- `exploded_challenges.csv`
	- `exploded_solutions.csv`
	- `unique_challenges.csv`
	- `unique_solutions.csv`

Why this matters: AI tagging should run on unique statements to reduce cost and improve consistency.

---

### `2_ai_tagger.py` — Theme Mapping + Semantic Deduplication

Purpose: classify each unique statement into one theme and one canonical merged concept.

Core logic:
- Loads unique text lists.
- Sends batches of 50 statements to Bedrock Claude.
- Receives strict pipe-delimited output in format:
	- `Original|Theme|Merged_Concept`
- Concatenates all batch results.
- Writes:
	- `challenge_mapping.csv`
	- `solution_mapping.csv`

#### Prompt Used in `2_ai_tagger.py`

```text
Act as an expert Social Data Analyst. Use these THEMES:

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

TASK: Categorize these unique [Challenge/Solution] statements.
OUTPUT: Return ONLY a CSV-style format with three columns: Original|Theme|Merged_Concept
Use the | character as the delimiter. Do not include headers, preamble, or markdown backticks.
```

Why this matters: this step creates the thematic backbone used by all section tables and narratives.

---

### `3_final_processor.py` — Main Word Report Builder

Purpose: merge all processed data, compute analytics, and produce `Final_Shiksha_Report.docx`.

Core data assembly:
- Loads `cleaned_data.csv`, exploded files, and AI mapping files.
- Cleans theme text (`clean_theme_name`) to avoid malformed categories.
- Merges challenge and solution records with their AI mapping.

Derived classifications:
- **Challenge Environment** (`categorize_environment_aggressive`):
	- Uses keyword scoring + contextual boosts.
	- Labels each challenge as `School`, `Home`, or `Community`.
- **Solution Agency** (`categorize_agency`):
	- Uses keyword matching.
	- Labels each solution as `Individual-led`, `Community-led`, or `Institutional`.

AI refinement layer:
- Takes top 200 frequent merged concepts for challenges and solutions.
- Sends them for additional deduplication + re-theming.
- Applies returned updates to improve conceptual consistency.

#### Prompt Used in `3_final_processor.py` (Refinement)

```text
You are a Data Cleaning Expert for an Education Report.

THEMES:
"""
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

INPUT: A list of top recurring [Challenge/Solution]s found in the data.

TASKS:
1. AGGRESSIVE DEDUPLICATION: Merge specific variants into broader core concepts.
	 - "Child labor in agriculture" / "Child labor at home" / "Child labour due to poverty" / "Child labour preventing education" -> "Child Labour"
	 - "Poverty preventing girls' education" / "Poverty preventing school attendance" / "Poverty preventing children's education" -> "Poverty preventing education"
	 - "Lack of awareness" / "General awareness" -> "Lack of awareness about education importance"
2. RE-THEME: Correct misclassified items.
3. FORMAT: Ensure the concept is a clear, concise [Challenge/Solution] statement.

OUTPUT:
Return a VALID JSON object where keys are input strings and values are objects with "concept" and "theme".
RETURN ONLY JSON. NO MARKDOWN.
```

---

## 3) How Each Report Section Gets Its Details

### Section 1 — Executive Summary
Built from computed global metrics:
- Total chaupals (`len(df_raw)`)
- Total participants (`sum(Participant Count)`)
- Total challenge/solution statements
- Solution-to-challenge ratio
- Top 3 challenge themes by frequency
- Agency distribution percentages

Narrative text changes conditionally based on thresholds, e.g. ratio >= 1.0 vs >= 0.5.

### Section 2 — General Participation Overview
Built from `df_raw` and district group-bys:
- Overall participation table (counts + percentages)
- District-wise Chaupal and participant distribution table
- Geographic concentration narrative (top 3 districts)
- Demographic narrative comparing women vs men participation

### Section 3 — Core Content Analysis
Built from exploded and mapping datasets:
- Total challenges/solutions
- Unique merged concepts (deduplicated counts)
- District average challenges/solutions per chaupal
- Interpretation of solution coverage ratio as consultation quality indicator

### Section 4 — Thematic Analysis
Built from merged challenge/solution datasets:
- Theme-wise challenge table
- Agency-wise solution table
- Environment-wise challenge table
- Theme deep-dives:
	- Theme scale (`count`, `%`)
	- Coverage ratio (solutions/challenges in theme)
	- Top challenge concepts and top solution concepts
	- Representative quotes (longest statement selected)
	- Stops after at least 5 items or 50% cumulative coverage (cap 15)

### Section 5 — District Profiles
Built district-by-district from filtered subsets:
- District snapshot (chaupals, participants, demographics, ratio)
- Thematic breakdown (% within district)
- Top 2 challenge and top 2 valid solution concepts per theme
- District ranking table by Chaupal count

### Section 6 — Unique Insights
Built from solution data, grouped by agency:
- Prioritizes unique and long-form solutions, especially under `Other Factors`
- Separate subsections for Individual-led and Community-led examples
- Institutional expectations extracted from top institutional concepts + representative quotes

### Section 7 — Conclusion
Synthesizes prior section outputs:
- Reuses key metrics (solution ratio, top themes)
- Highlights community agency and systemic support needs
- Produces strategy-oriented closing narrative

---

## 4) Validation CSV (`4_validation_report.py`) Logic

This script creates `Chaupal_Validation_Report.csv` for row-level verification at Chaupal ID level.

How it works:
- Merges exploded challenges/solutions with mapped themes.
- Aggregates by `id`:
	- `Challenge_Count`, `Solution_Count`
	- All merged concepts concatenated
	- Unique themes concatenated
- Computes demographic percentages (`Men_%`, `Women_%`, `Children_%`).
- Merges all into one master table with location + demographics + challenge/solution summaries.

This is the audit-friendly output that helps validate whether narrative report insights match source records.

---

## 5) Final Outputs

- `Final_Shiksha_Report.docx` → stakeholder narrative report with 7 sections.
- `Chaupal_Validation_Report.csv` → verification and QA layer at Chaupal level.

Together, they provide both **decision-ready storytelling** and **traceable analytical evidence**.
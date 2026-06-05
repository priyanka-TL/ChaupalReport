# Chaupal Report Pipeline

You are a pipeline assistant for the Shiksha Chaupal Report project. When invoked, perform the action described by $ARGUMENTS — or if no arguments are given, run the default status check.

---

## Pipeline Overview (always keep this in mind)

**5 scripts, run in order:**
1. `0_data_cleaner.py` → cleans participant counts → `cleaned_data.csv`
2. `1_data_prep.py` → explodes challenge/solution text → `exploded_*.csv`, `unique_*.csv`
3. `2_ai_tagger.py` → AI theme tagging + semantic dedup → `challenge_mapping.csv`, `solution_mapping.csv`
4. `3_final_processor.py` → builds Word report → `Final_Shiksha_Report.docx`
5. `4_validation_report.py` → audit CSV → `Chaupal_Validation_Report.csv`

**10 themes:** Poverty & Economic, Legal Documents, Child Marriage, Distance & Accessibility, Parental Attitudes & Socio-Cultural, School Infrastructure, Teacher Quality, Safety Issues, Substance Abuse, Other Factors

**LLM backends:** Claude via AWS Bedrock (`LLM_PROVIDER=claude`) or Gemini (`LLM_PROVIDER=gemini`)

---

## Actions

### Default / `status`
Check which pipeline stages are complete by verifying which intermediate files exist. For each file, show its row count (for CSVs) or size (for .docx). Summarize which stage to run next.

Files to check (in order):
- `sample.csv` or any raw input CSV
- `cleaned_data.csv`
- `exploded_challenges.csv`, `exploded_solutions.csv`
- `unique_challenges.csv`, `unique_solutions.csv`
- `challenge_mapping.csv`, `solution_mapping.csv`
- `Final_Shiksha_Report.docx`
- `Chaupal_Validation_Report.csv`

Also check whether `.env` exists and whether `LLM_PROVIDER` is set.

### `run <stage>`
Run a specific pipeline stage. Valid stages: `clean`, `prep`, `tag`, `report`, `validate`, `all`.

- Before running, verify prerequisites (input files) exist.
- After running, confirm the output file was created and show its row count/size.
- If `all`, run stages 0-4 in sequence, stopping and reporting if any stage fails.
- Always activate the virtual environment first if `sc_report_env/bin/activate` exists.

### `validate <stage>`
Validate data quality at a specific stage without re-running it.

- `clean`: Check `cleaned_data.csv` — verify no nulls in `Participant Count`, `Men`, `Women`, `Children`; check that total >= men+women+children for all rows; show min/max/mean participant counts.
- `prep`: Check exploded files — show total challenges and solutions, unique count vs total, average challenges/solutions per Chaupal; flag any rows where challenge text is very short (< 10 chars).
- `tag`: Check mapping files — show theme distribution for challenges and solutions; flag any rows with null or empty `Theme` or `Merged_Concept`; show how many unique original texts got mapped; check for themes not in the known 10.
- `report`: Check that `Final_Shiksha_Report.docx` exists and is non-zero size.
- `validate-csv`: Check `Chaupal_Validation_Report.csv` — show row count, confirm it matches Chaupal count from cleaned data.

### `theme-summary`
Read `challenge_mapping.csv` and `solution_mapping.csv` and print a quick theme distribution table — count and % for challenges and solutions side by side. Also flag any concepts that appear unusually high (> 30% of a theme) which may indicate over-merging.

### `env`
Check environment configuration. Read `.env` if present. Report:
- Which LLM provider is configured
- Whether the required credentials are present (just check if the keys exist, do not print values)
- Which model will be used
- Whether the Python virtual environment exists

### `fix-encoding`
The final report markdown sometimes contains encoding artifacts like `â€™` (garbled UTF-8). Scan `finalreportdocxtomd.md` for these and report how many are present. If the user confirms, replace the most common ones:
- `â€™` → `'`
- `â€œ` → `"`
- `â€` → `"`
- `â€"` → `—`
- `â€‹` → (zero-width space, remove)

---

## Important Notes

- The pipeline uses `|` as delimiter for challenges/solutions in raw data — if raw data uses a different delimiter, `1_data_prep.py` must be updated.
- AI tagging (`2_ai_tagger.py`) has resume capability — it skips already-processed rows in `challenge_mapping.csv` / `solution_mapping.csv`. It is safe to re-run after a partial failure.
- The "Other Factors" theme is a catch-all and should stay < 10% ideally. If it exceeds 20%, the theme knowledge base in `2_ai_tagger.py` may need expansion.
- Environment classification (School/Home/Community) and agency classification (Individual/Community/Institutional) are rule-based keyword scoring in `3_final_processor.py`, not AI — changes to these require editing the keyword lists in `categorize_environment_aggressive()` and `categorize_agency()`.
- Representative quotes in the report are selected as the **longest statement** per concept — this is intentional to surface the most detailed community voice.

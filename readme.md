pip install -r requirements. txt

1. Create a virtual environment
python3 -m venv sc_report_env

2. Activate the virtual environment
source sc_report_env/bin/activate

3. Install dependencies
pip install pandas numpy python-docx

4. Run the script - Data prep, fix counts
python 0_data_prep.py

5. Run the script - Grouping the challenges and solutions to theme
python 1_ai_tagger.py

6. Run the script - Doc report generation
python 3_final_processor.py

7. Run the script - CSV report generation
python 4_validation_report.py
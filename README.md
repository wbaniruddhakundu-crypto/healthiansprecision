# Healthians Inter-Run Precision Populator

A Streamlit application that reads Siemens assay PDF reports and populates a supplied Healthians Excel precision template day-wise.

## What this version does

- Accepts one Excel `.xlsx` precision template.
- Accepts multiple PDF reports for Day 1 through Day 5 (for example, separate routine, tumor-marker, and vitamin-D reports).
- Combines all uploaded PDFs assigned to the same day.
- Detects analyte headings and numeric result values with recognized units.
- Populates up to five replicates per analyte for each day in columns B:F and rows 11:15.
- Leaves missing days and missing replicate cells blank.
- Does not change the date cells already present in the template.
- Preserves the template's merged-cell layout and writes the analyte name in the correct merged header anchor (D4), avoiding the `MergedCell value is read-only` error.
- Ignores RLU/MAU values and populates only result values accompanied by units.
- Adds day-wise N, Mean, SD, CV and overall inter-run summary formulas.
- Creates one worksheet per detected analyte.

## Run locally on Windows

1. Install Python 3.10 or later.
2. Extract this ZIP file.
3. Open Command Prompt in the extracted `streamlit_precision_app` folder.
4. Run:

```bash
python -m venv .venv
.venv\\Scripts\\activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud using GitHub

1. Create a new GitHub repository, for example `HEALTHIANS-PRECISION`.
2. Upload the three files from this folder (`app.py`, `requirements.txt`, and `README.md`) to the repository root.
3. Open Streamlit Community Cloud and choose **Create app**.
4. Select your repository and branch `main`.
5. Set **Main file path** to exactly `app.py`.
6. Deploy the app.

## PDF filename requirement

Each PDF filename must include a day number from 1 to 5. Examples:

- `DAY 1 INTER ALL.pdf`
- `DAY 2 ROUTINE THCG AND EE2.pdf`
- `DAY 3 TUMOR MARKERS.pdf`
- `DAY 4 VIT D.pdf`
- `DAY 5 ROUTINE INTER.pdf`

You may upload several PDFs for the same day. Their analyte results are combined automatically.

## Important limitations and validation

- This version supports text-based/selectable-text PDFs. Scanned image-only PDFs need OCR and are not supported here.
- For each analyte/day, the first five detected result values are used. Additional values are not populated.
- If the same analyte is repeated across multiple PDFs for the same day, values from those files are concatenated in upload order; check for duplicate reports before using the output.
- The app keeps the template date cells unchanged. It does not infer or insert dates.
- Review the extracted results and Excel formulas against the source reports before using the workbook for formal precision validation.

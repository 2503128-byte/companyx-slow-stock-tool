# Company X Slow-Stock Review

A Streamlit app for the Company X Technology & Data Management case study. It reads the four required data tables, normalizes SKU keys, checks data quality, consolidates only explicit `SWITCH` pack-size replacements, classifies active items, recommends customer targets, and exports an Excel workbook.

## Run locally

```bash
python -m venv .venv
# Windows: .venv\\Scripts\\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Open the local URL printed in the terminal (usually http://localhost:8501). The sample data is bundled in `sample_data/`; uncheck **Load included sample data** to upload a new monthly workbook or four separate exports.

## Deploy for free with Streamlit Community Cloud

1. Create a GitHub repository and upload `app.py`, `engine.py`, `requirements.txt`, `.gitignore`, `README.md`, and `sample_data/Case_Data_Technology_and_Data_Management_All_Sheets.xlsx`.
2. Sign in to Streamlit Community Cloud with GitHub and choose **Create app**.
3. Select the repository, branch, and `app.py` as the main file.
4. Deploy and open the generated public URL. Test sample mode and upload mode.
5. Keep the repository and app available for at least 30 days after submission. Do not commit passwords, API keys, or private data.

The public hosting URL and GitHub repository must be created from the candidate's own accounts; this package does not claim to have published either link.

## Calculation rules

- Only `Alternate Flag = SWITCH` maps the old SKU to the alternate item's SKU. Stock and sales are aggregated to the replacement SKU and reported using the replacement master pack and unit cost.
- Inactive items are excluded from classification and targets, but inactive stock is reported as an alert.
- Aged quantity is the sum of the 91–180 and 180+ day buckets; aged value is aged quantity multiplied by unit cost.
- Average monthly sales is six-month quantity divided by 6; months to clear is total stock divided by average monthly sales.
- Classification order: Red List if six-month quantity is below one master pack; otherwise No Action if aged quantity is zero or months to clear is below 3; otherwise Push.
- Customer targets require more than ₹1,00,000 total purchase value in the six-month period. Buckets are exclusive and prioritized 1 (bought item/switch partner), 2 (same industry as an item buyer, only when industry exists), then 3 (bought another item in the same category).
- Exact duplicate sales rows are removed and reported as a data-quality issue. Negative aging-bucket quantities are floored at zero for calculations and flagged for source reconciliation. These are explicit handling choices, not proof that the source data is correct.

## Test results on the supplied case data

- 38 active canonical items analysed after consolidating 3 old switch SKUs into their replacement SKUs.
- Red List: 7; No Action: 13; Push: 18.
- Total aged value in Push items: ₹163,841.
- SKU1001: No Action; SKU1006: Push; SKU1019: Push.
- First three targets for SKU1005: Customer 220 (Bucket 1), Customer 224 (Bucket 1), Customer 214 (Bucket 2).

## AI use disclosure (adapt before submitting)

ChatGPT was used as a coding and analysis assistant to structure the Streamlit app, draft data-validation logic, review the case rules, and generate documentation. The candidate should personally run the app, inspect the outputs, understand the functions, and be prepared to explain and modify the code. Do not claim that the public deployment is live until it has been published and tested.

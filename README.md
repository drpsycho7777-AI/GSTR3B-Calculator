# GSTR-3B Online Turnover Calculator

A Streamlit Community Cloud-ready version of the GSTR-3B automatic turnover tool.

## Core rules

- Reads original text-based GST portal GSTR-3B PDFs.
- Turnover is **Table 3.1 (A) + (B) + (C) only**.
- **(D) Inward RCM** and **(E) Non-GST outward** are excluded from turnover.
- Checks missing returns using the business rule that a tax month's GSTR-3B is expected by the **18th of the next month**.
- Supports **Last 12 Months** and **Last 24 Months**.
- Detects duplicate periods and uses the latest filed copy.
- Supports multiple GSTINs and creates separate reports in a ZIP.
- Generates an Excel report with GST calculation, source audit trail, and FY summary.

## Deploy on Streamlit Community Cloud

### 1. Create a new GitHub repository
Recommended name: `GSTR3B-Calculator`

Make it **Public** for the simplest Streamlit Community Cloud deployment.

### 2. Upload the files directly to the repository root
After extracting the deployment ZIP, your GitHub repository must look like this:

```text
GSTR3B-Calculator/
├── app.py
├── gstr3b_core.py
├── requirements.txt
├── README.md
├── VERSION.txt
├── .gitignore
└── .streamlit/
    └── config.toml
```

**Important:** Do not upload the outer folder as another subfolder. `app.py` must be visible on the first page of the repository.

### 3. Deploy in Streamlit Community Cloud

Choose:

- Repository: `YOUR-GITHUB-USERNAME/GSTR3B-Calculator`
- Branch: `main`
- Main file path: `app.py`

Then click **Deploy**.

### 4. Confirm deployment

The first visible message should be:

> App loaded successfully — Online Deployment 3.1

If this message appears, the Streamlit frontend and Python app have loaded correctly.

## Usage

1. Select **Last 12 months** or **Last 24 months**.
2. Select the calculation date.
3. Upload the original GSTR-3B PDFs.
4. Check the required window, missing returns, and turnover.
5. Click **Generate Excel Report**.
6. Download the Excel file.

## Notes

- Image-only/scanned PDFs are intentionally rejected because the tool does not guess GST values.
- For quarterly GSTR-3B returns crossing the edge of a selected period, the tool warns that the quarterly turnover cannot be split exactly across the boundary.

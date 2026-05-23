# Project Tracking Dashboard - Streamlit

This Streamlit app reads your existing Excel workbook and recreates the main dashboard sections from your HTML dashboard:

- KPI cards
- Quick status views
- Phase progress cards
- Project Status, Material Status, Priority charts
- Region Breakdown and Top Customers
- Delivery Items Status
- Overall Progress Buckets
- Phase Progress by Project
- Project Details table with filters

## 1. Local setup

Put your Excel file next to `app.py` and name it:

```text
Project_Tracking_v7.xlsx
```

Then run:

```bash
pip install -r requirements.txt
streamlit run app.py
```

You can also test by uploading the Excel file inside the app.

## 2. OneDrive live setup - easiest method

Use this if the file can be shared by link.

1. Open your Excel file in OneDrive.
2. Click **Share**.
3. Create a link.
4. Copy the link.
5. In Streamlit Cloud, open your app settings and add this in **Secrets**:

```toml
EXCEL_FILE_URL = "PASTE_YOUR_ONEDRIVE_LINK_HERE"
REFRESH_SECONDS = 60
```

The app automatically adds `download=1` to the URL.

If the app says the URL returned HTML instead of Excel, your link is not a direct download link. In that case, create a new sharing link or use Microsoft Graph mode.

## 3. OneDrive private setup - Microsoft Graph mode

Use this for company Microsoft 365 OneDrive or SharePoint if you need private access.

Required Streamlit secrets:

```toml
GRAPH_TENANT_ID = "your-tenant-id"
GRAPH_CLIENT_ID = "your-app-client-id"
GRAPH_CLIENT_SECRET = "your-client-secret"
GRAPH_USER_ID = "your.email@company.com"
ONEDRIVE_FILE_PATH = "/Documents/Project_Tracking_v7.xlsx"
REFRESH_SECONDS = 60
```

The Azure app should have Microsoft Graph **Application** permission:

```text
Files.Read.All
```

Admin consent may be required.

## 4. Deploy on Streamlit Community Cloud

1. Create a GitHub repo.
2. Upload:
   - `app.py`
   - `requirements.txt`
3. Deploy the repo in Streamlit Community Cloud.
4. Add your OneDrive secrets in Streamlit Cloud.

Do not upload `.streamlit/secrets.toml` to GitHub.

## Notes

- The app refreshes every 60 seconds by default.
- Excel edits must be saved/synced to OneDrive before Streamlit can read the new data.
- The parser expects the workbook structure from `Project_Tracking_v7.xlsx`, especially the `Project_Master` sheet.

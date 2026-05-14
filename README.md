# Nooks Financial Dashboard - SharePoint Auto-Sync

Automatically syncs your Excel financial model from SharePoint to your Vercel-hosted dashboard every 15 minutes.

## 🎯 How It Works

```
SharePoint (Excel lives here)
    ↓ (every 15 min)
GitHub Actions downloads Excel
    ↓
Extracts scenario data
    ↓
Updates HTML with new data
    ↓
Commits to GitHub
    ↓
Vercel auto-deploys
```

## 📋 Prerequisites

1. **SharePoint** - Excel file stored in a SharePoint folder
2. **GitHub** - This repository
3. **Vercel** - Connected to your GitHub repo

## 🚀 Quick Start

### Step 1: Create Test Excel File

Create a simple Excel file to test the pipeline. Save as `nooks_q2_model.xlsx`:

**Sheet name:** `Scenarios`

**Structure:**
```
              FY25A   FY26E   FY27E   FY28E   FY29E
Ops - Revenue    19648   41699   58728   80029   103046
Ops - EBITDA      5388   10299   26925   39706    51355
Ops - Payroll     6790   19410   16888   22868    27407
Ops - Margin      0.274   0.238   0.458   0.496   0.498

Bridge - Revenue  19648   41699   67538  106290   148180
Bridge - EBITDA   5388   10299   33108   58014    86216
Bridge - Payroll  6790   19410   20265   29382    36482
Bridge - Margin   0.274   0.247   0.490   0.546   0.582

Full - Revenue    19648   41699   86561  160200   177653
Full - EBITDA     5680   12491   47310   97443   115663
Full - Payroll    6790   19779   26420   48200    56800
Full - Margin     0.289   0.299   0.547   0.608   0.651
```

**Important:**
- First column = row labels (e.g., "Ops - Revenue")
- Columns 2-6 = year values
- Margin values as decimals (0.274 = 27.4%)

### Step 2: Upload to SharePoint

1. Upload your test Excel to SharePoint
2. Note the **folder path**, e.g.:
   - `Shared Documents/Financial Models`
   - `Documents/Board Materials`

### Step 3: Set Up GitHub Secrets

Go to your GitHub repo → Settings → Secrets and variables → Actions

Add these secrets:

| Secret Name | Example Value | Description |
|------------|---------------|-------------|
| `SHAREPOINT_SITE_URL` | `https://yourcompany.sharepoint.com/sites/Finance` | Your SharePoint site URL |
| `SHAREPOINT_FOLDER` | `Shared Documents/Financial Models` | Folder path where Excel lives |
| `EXCEL_FILENAME` | `nooks_q2_model.xlsx` | Exact Excel filename |
| `SHAREPOINT_USERNAME` | `your.email@company.com` | Your Microsoft 365 email |
| `SHAREPOINT_PASSWORD` | `yourpassword` | Your Microsoft 365 password |

**Note:** For production, use App Registration (see Advanced Setup below)

### Step 4: Test Locally First

```bash
# Install dependencies
pip install -r requirements.txt

# Set environment variables
export SHAREPOINT_SITE_URL="https://yourcompany.sharepoint.com/sites/Finance"
export SHAREPOINT_FOLDER="Shared Documents/Financial Models"
export EXCEL_FILENAME="nooks_q2_model.xlsx"
export SHAREPOINT_USERNAME="your.email@company.com"
export SHAREPOINT_PASSWORD="yourpassword"

# Test the validation
python scripts/test-extraction.py
```

## 📝 What You Need to Provide

To get started, give me:

1. **SharePoint Site URL** - e.g., `https://nooks.sharepoint.com/sites/Finance`
2. **Folder Path** - e.g., `Shared Documents/Financial Models`
3. **Excel Filename** - e.g., `nooks_q2_model.xlsx`
4. **Excel File** - Re-upload it here so I can validate the structure

Then I'll customize the extraction logic for your specific Excel layout!

---

## 🔍 Monitoring

Check GitHub Actions tab for sync status every 15 minutes.

**Logs will show:**
```
📡 Connecting to SharePoint
✅ Downloaded to: data/nooks_q2_model.xlsx
📝 File changed - updating hash
✅ Data extracted
✅ HTML updated
```

---

**Ready to customize this for your specific Excel?** Re-upload your file and provide those SharePoint details! 🚀

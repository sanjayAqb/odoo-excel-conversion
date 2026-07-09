# Odoo Excel Conversion

Browser-based tool that converts supplier product Excel/CSV files into a predefined Odoo import format. All processing runs entirely in the browser — files never leave the user's device.

**Live repo:** [github.com/sanjayAqb/odoo-excel-conversion](https://github.com/sanjayAqb/odoo-excel-conversion)

## Features

- Upload supplier files: `.csv`, `.xlsx`, `.xls`
- Automatic file format detection
- Maps supplier columns to the Odoo import template
- Preserves the original Odoo template headers, column order, and structure
- Static defaults applied to every product row
- Preview of converted data before download
- Download as **Excel (.xlsx)** or **CSV (.csv)**
- Download filenames include a datetime stamp
- Progress bar with percentage during conversion
- Validates required supplier columns with clear error messages

## How it works

1. Open the app in a modern browser.
2. Browse or drag-and-drop a supplier product sheet.
3. Click **Convert to Odoo format**.
4. Review the preview.
5. Download Excel or CSV.

Processing uses **SheetJS** for file parsing and **Python (Pyodide + openpyxl)** for mapping and Excel generation — all client-side.

## Column mapping

| Supplier column | Odoo template column |
|-----------------|----------------------|
| Name | Name |
| Barcode | Barcode |
| Cost/Case | Cost |
| Selling Price | Sales Price |
| PLU | Internal Reference |
| Markup / Markup(%) | Markup % |
| VAT (%) / VAT(%) | VAT % |

### Static values (not from supplier sheet)

| Odoo column | Value |
|-------------|-------|
| Track Inventory | `1` |
| Product Category | `All` |

All other Odoo template columns are left unchanged in structure; only the mapped fields and static values above are populated.

## Project structure

```
odoo-excel-conversion/
├── index.html              # App UI
├── vercel.json             # Vercel static site config
├── static/
│   ├── app.js              # Upload, conversion flow, downloads
│   ├── convert.py          # Mapping and file generation (Pyodide)
│   ├── styles.css
│   └── favicon.svg
└── sample_sheets/
    ├── SampleOdooFormat.xlsx           # Odoo import template
    └── NICOTIN POUCH PRODUCT LIST.csv  # Sample supplier file
```

## Run locally

Serve the project over HTTP (required so the app can load the template and Python module):

```bash
cd odoo-excel-conversion
python3 -m http.server 8080
```

Then open: **http://localhost:8080/**

With XAMPP, place the folder under `htdocs` and open:

**http://localhost/odoo_excel_conversion/**

## Deploy to Vercel (Git CI/CD)

This is a static site — no build step or backend required. Vercel auto-deploys on every push to `main`.

### One-time setup

1. Push the repo to GitHub (already done if using `sanjayAqb/odoo-excel-conversion`).
2. Go to [vercel.com](https://vercel.com) and sign up with **GitHub**.
3. Click **Add New → Project**.
4. Import `sanjayAqb/odoo-excel-conversion`.
5. Use these settings:
   - **Framework Preset:** Other
   - **Build Command:** *(leave empty)*
   - **Output Directory:** *(leave empty)*
   - **Install Command:** *(leave empty)*
6. Click **Deploy**.

### CI/CD workflow (after setup)

Every push to `main` triggers a new deployment automatically:

```bash
git add .
git commit -m "Describe your change"
git push
```

Vercel builds and deploys within about a minute. Check status under **Deployments** in the Vercel dashboard.

### Custom domain (optional)

In Vercel: **Project → Settings → Domains** to add your own domain.

## Git setup (SSH)

Remote is configured for SSH:

```bash
git remote -v
# origin  git@github.com:sanjayAqb/odoo-excel-conversion.git
```

Push without a password:

```bash
git push
```

## Requirements

- Modern browser (Chrome, Firefox, Edge, Safari)
- Internet connection on first load (Pyodide and SheetJS load from CDN)

## Sample files

Use the files in `sample_sheets/` to test:

- **Input:** `NICOTIN POUCH PRODUCT LIST.csv`
- **Template:** `SampleOdooFormat.xlsx`

## License

Private / internal use unless otherwise specified.

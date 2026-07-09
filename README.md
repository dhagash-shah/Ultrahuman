# Ultrahuman Growth Huddle — dashboard

Static dashboard on Vercel with four modules: **MIS Dashboard**, **Capital & Cap
Table**, **Comps**, **Special Projects**. The MIS module is computed **directly from
the MIS + Blume Model workbooks on Google Drive** — no `/cashify-monthly-review`, no
copy-paste into a Growth-Huddle sheet.

## New process (automated)

```
Google Drive  ──drive_pull.py (service account, read-only)──►  local source files
   ├─ latest MIS + Blume Model ─ build_mis.py     ─► data.json.mis
   ├─ cap-table PDF + Invest.   ─ build_capital.py ─► data.json.capital
   │  Profile + Growth Huddle
   └─ Growth Huddle             ─ build_data.py    ─► data.json.comps / .specialProjects
                                                         │ fetch
                                                         ▼
                                                   index.html (static, on Vercel)
```

A **GitHub Action** (`.github/workflows/refresh.yml`) runs this monthly (and on a
manual "Run workflow" click): it pulls the latest files from Drive, rebuilds
`data.json`, and commits it. **Vercel auto-deploys on the push.** You only ever
update the workbook on Drive.

## Files to upload to GitHub

| File | Role |
|------|------|
| `index.html` | the dashboard (reads `data.json`) |
| `data.json` | generated data contract (committed; the site reads this) |
| `build_mis.py` | MIS engine — MIS + Model → `mis` block |
| `build_capital.py` | Capital block — cap-table PDF + Investment Profile + Growth Huddle → `capital` |
| `build_data.py` | Comps / Special Projects (+ legacy capital) from the Growth Huddle workbook |
| `drive_pull.py` | downloads all source files from Drive (service account) |
| `requirements.txt`, `.github/workflows/refresh.yml` | the automation |
| `.gitignore` | keeps all `*.xlsx` / `*.pdf` out of git |

**Never commit the `.xlsx` / `.pdf` source files** — confidential and git-ignored.

### Capital & Cap Table — sources (Part 2)

| Block | Source | Unit on sheet |
|-------|--------|---------------|
| Cap table (entity-level, diluted %) | Shareholding-pattern **PDF** ([1i2s…](https://drive.google.com/file/d/1i2s23yUKGI9Yiv3nhXJLhJeux0Jzk-6L/view)) | % |
| Valuation & **fund-wise** holding | Blume **Investment Profile** ([1iFb…](https://docs.google.com/spreadsheets/d/1iFbPSnTkZsjH6WTc8Lgp7F_BAMT3FS5g/edit)) | **₹ Mn** → shown ₹ Cr |
| Round-wise details | Growth Huddle *Exit Thinking* ([1vZV…](https://docs.google.com/spreadsheets/d/1vZVPXdimWi7yHwpbLfr4p2qtsywFX9qx/edit)) | ₹ Cr |

These links are embedded in `build_capital.py` (`LINKS`) and shown as "View source"
buttons in the Capital module. All money is normalised to **₹ Cr** and every figure
is unit-labelled. **Note:** the prior "₹8.9 Cr cost / ₹10.3 Cr MTM" were actually
**$ Mn** mislabelled — the Investment Profile (₹ Mn) corrects this to ₹72.7 Cr cost /
₹85.7 Cr MTM (stakes unchanged at 4.75%).

## One-time setup for the automated pipeline

1. **Service account:** in Google Cloud, create a service account + JSON key; enable
   the Drive API.
2. **Share the Ultrahuman Drive folders** with the service account's email (Viewer):
   the Ultrahuman root folder, MIS folder, Growth Huddle folder and Cap Tables folder.
3. **GitHub → Settings → Secrets and variables → Actions:**
   - Secret `GOOGLE_SERVICE_ACCOUNT_JSON` = the full key JSON.
   - Variable `CASHIFY_ROOT_FOLDER_ID` = `1TlP_wE_sZ15qi0qdod0NUssqXcC50e-6`.
   - Recommended variables:
     - `CASHIFY_MIS_FOLDER_ID`
     - `CASHIFY_GROWTH_HUDDLE_FOLDER_ID`
     - `CASHIFY_CAP_TABLES_FOLDER_ID`
4. Connect the repo to Vercel (static — no build command needed).

The Action picks the MIS file with the **latest month in the filename**. If two
MIS files have the same parsed month, the most recently modified file wins. Model,
Growth Huddle, cap table and investment profile files are pulled from their mapped
folders. The reporting month is auto-detected (latest complete actual). Sanity
checks run every build (incl. the Dec '25 control =
₹1,692,076,864) and fail loudly if a number drifts.

## Run it locally

```bash
pip install -r requirements.txt
# order matters — build_data writes a legacy capital block, build_capital overrides it
python3 build_data.py    growth_huddle.xlsx data.json                    # comps + special projects
python3 build_mis.py     mis.xlsx model.xlsx data.json                   # MIS block
python3 build_capital.py captable.pdf investment_profile.xlsx growth_huddle.xlsx data.json
python3 -m http.server 8766                               # http://localhost:8766
```
(`fetch` needs HTTP — opening `index.html` via `file://` won't load data.)

## Notes

- **Confidentiality:** `data.json` contains cap table, valuation, Blume's stake and
  MIS financials. Keep the **repo private** and turn on **Vercel Deployment
  Protection**. The "View MIS / View Blume Model" links point to Drive (item 11) —
  edit `meta.links` in `build_mis.py` to exact per-file share URLs if preferred.
- The **Capital / Comps / Special Projects** modules currently source from
  `growth_huddle.xlsx` via `build_data.py`. To automate those from Drive
  too, add them to the folder and extend `drive_pull.py` + the workflow the same way.
- **Store count (item 10)** was intentionally skipped — not a clean actual in the MIS.

# Returns Reconciliation

A small Streamlit tool for e-commerce ops teams that cross-checks marketplace return/refund
exports (Shopee, Lazada, TikTok) against an existing brand "Return Report" tracker workbook,
flags which return Order IDs are already tracked vs. brand new, and appends the new rows with
fields auto-filled from the marketplace files and/or a TC Order Report.

## Why this exists

Returns tracking for multi-marketplace e-commerce accounts usually lives in one big Excel
workbook per brand, updated by hand every time a new batch of marketplace return exports comes
in. The manual part is tedious and error-prone in one specific way: numeric order IDs (e.g.
Lazada's) get loaded as floats by Excel/openpyxl in some places and as strings elsewhere, which
causes already-tracked returns to look "new" if you're comparing by eye or with a naive
spreadsheet VLOOKUP. This tool normalizes IDs before comparing, so the new-vs-tracked split is
reliable.

## Features

- Upload a tracker workbook + any combination of Shopee / Lazada / TikTok return exports
- Optional TC Order Report upload to backfill Invoice Number and SKU when the marketplace
  export doesn't have them
- Automatically finds the tracker's real header row (skipping title/instruction rows above it)
- Normalizes numeric order IDs so float-vs-string formatting doesn't create false positives
- Copies cell formatting from existing rows so new rows look native to the sheet
- Leaves manual-confirmation columns (Return Confirmation, Fault Description, Status, Dispute,
  etc.) blank intentionally — those are for your ops/warehouse team to fill in later
- Per-marketplace summary: how many already tracked, how many new, any files that came back
  empty (TikTok exports sometimes arrive as header-only)
- Download the updated tracker as a ready-to-use `.xlsx`

## Project structure

```
returns-reconciliation/
├── app.py              # Streamlit UI
├── report_engine.py    # Core logic (no Streamlit dependency, importable/testable on its own)
├── requirements.txt
└── README.md
```

## Running locally

```bash
git clone <this-repo>
cd returns-reconciliation
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (usually `http://localhost:8501`).

## Deploying to Streamlit Community Cloud

1. Push this repo to GitHub.
2. Go to [share.streamlit.io](https://share.streamlit.io), connect your GitHub account.
3. Pick this repo, branch `main`, and set the main file path to `app.py`.
4. Deploy — no secrets or environment variables are required.

## Input file expectations

| File | Format | Required columns |
|---|---|---|
| Tracker | `.xlsx`, sheet name containing "Return report" | A column containing "Order Number"/"Order ID" somewhere in the header row |
| Shopee return export | `.xls` or `.xlsx` | `Order ID`, plus optionally `SKU`, `Return Quantity`, `Return Reason`, `Return Tracking Number`, `Order Creation Date`, `Return Creation Time` |
| Lazada return export | `.xlsx` | `Order ID`, plus optionally `Return Order Date`, `Seller SKU ID` |
| TikTok return export | `.xlsx` | A column containing "Order ID" (these exports sometimes arrive empty) |
| TC Order Report | `.csv` | `order_id`, plus optionally `invoice_number`, `sku` |

The tracker's header row doesn't have to be in any fixed position — the app scans the first 15
rows looking for an "Order Number"/"Order ID"-like header to handle the title/instruction rows
that brand trackers usually have above the real header.

## Known limitations

- Only one sheet (the main "Return report" sheet) gets updated. Other sheets in the same
  workbook (SKU↔code lookup tables, raw marketplace pulls, exchange logs) are left untouched.
- If a marketplace file belongs to a *different* brand than the tracker, nothing will match and
  every order will show as "new" — there's no automatic brand-mismatch detection in the engine
  itself, so sanity-check the new-row count before downloading if it looks unexpectedly high.
- Multi-SKU orders are combined into a single cell (`SKU A / SKU B`) rather than split into
  multiple rows. If your tracker convention splits these out, edit the downloaded file manually
  for those orders.

## License

Internal tool — adapt freely for your own ops workflows.

"""
app.py
------
Streamlit UI for the Returns Reconciliation tool.

Upload a brand Return Report tracker (xlsx) plus any combination of marketplace
return exports (Shopee, Lazada, TikTok) and an optional TC Order Report (csv).
Each marketplace upload accepts .xls / .xlsx / .zip (the zip is auto-extracted
to find the return/refund file inside).
"""

from __future__ import annotations

import io
import zipfile
from datetime import datetime
from pathlib import Path

import streamlit as st

from report_engine import run_reconciliation, ReconciliationResult


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

EXCEL_EXTENSIONS = {".xls", ".xlsx"}
# Keywords used to identify the return/refund file inside a zip
RETURN_FILE_KEYWORDS = ["return_refund", "return refund", "returns", "refund"]
CANCEL_FILE_KEYWORDS = ["cancelled", "cancel"]


def extract_from_zip(uploaded_zip) -> io.BytesIO | None:
    """
    Extract the most relevant return/refund file from a zip upload.
    Priority: file whose name contains return/refund keywords and has an Excel
    extension. Skips cancelled-order files. Returns a BytesIO with .name set.
    """
    raw = uploaded_zip.read()
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        members = z.namelist()
        excel_members = [
            m for m in members
            if Path(m).suffix.lower() in EXCEL_EXTENSIONS
        ]

        # Prefer files that look like return/refund exports
        return_members = [
            m for m in excel_members
            if any(kw in m.lower() for kw in RETURN_FILE_KEYWORDS)
            and not any(kw in m.lower() for kw in CANCEL_FILE_KEYWORDS)
        ]

        target = return_members[0] if return_members else (excel_members[0] if excel_members else None)

        if target is None:
            return None

        file_bytes = z.read(target)
        buf = io.BytesIO(file_bytes)
        buf.name = Path(target).name   # preserve filename so engine can detect format
        return buf


def resolve_upload(uploaded_file) -> io.BytesIO | None:
    """
    Accept either a direct Excel upload or a zip containing an Excel file.
    Returns a BytesIO ready for the engine, or None if nothing usable was found.
    """
    if uploaded_file is None:
        return None

    name = uploaded_file.name.lower()

    if name.endswith(".zip"):
        buf = extract_from_zip(uploaded_file)
        if buf is None:
            st.warning(f"No Excel file found inside {uploaded_file.name}.")
        return buf

    # Direct Excel / CSV upload — wrap in BytesIO with name preserved
    raw = uploaded_file.read()
    buf = io.BytesIO(raw)
    buf.name = uploaded_file.name
    return buf


# ---------------------------------------------------------------------------
# Page
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Returns Reconciliation", page_icon="🔄", layout="wide")

st.title("🔄 Returns Reconciliation")
st.caption(
    "Cross-check marketplace return files against your brand's Return Report tracker, "
    "and auto-fill new rows for anything that isn't tracked yet."
)

with st.expander("ℹ️ How this works", expanded=False):
    st.markdown(
        """
        1. **Upload your tracker** — the brand's `*Return_report*.xlsx` workbook
           (e.g. `OMRON Return report - New Format` or `GSK Return report - New Format`).
        2. **Upload marketplace return files** for whichever platforms you have —
           Shopee, Lazada, and/or TikTok. Each uploader accepts:
           - `.xls` / `.xlsx` — direct export file
           - `.zip` — the zip downloaded from Seller Centre (the return/refund file
             is auto-extracted; cancelled-order files inside the zip are ignored)
        3. **Optionally upload the TC Order Report** (`.csv`) — used to fill Invoice
           Number and SKU when the marketplace file doesn't carry them.
        4. Click **Run reconciliation**. The app shows per-marketplace counts of
           already-tracked vs. new, previews the new rows, and lets you download
           the updated tracker.

        **Manual columns are never auto-filled** — Return Confirmation, Receiving
        Date, Fault Description, Dispute, Status, etc. are always left blank for
        your ops/warehouse team.

        **TikTok:** only returns with status *In Process* or *To Process* are added.
        Completed and Refund Rejected returns are skipped automatically.
        """
    )

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Tracker workbook")
    tracker_file = st.file_uploader(
        "Brand Return Report (.xlsx)",
        type=["xlsx"],
        help="The existing multi-sheet tracker workbook you want to update.",
    )

with col2:
    st.subheader("2. TC Order Report (optional)")
    tc_file = st.file_uploader(
        "TC Order Report (.csv)",
        type=["csv"],
        help="Used to enrich new rows with Invoice Number and SKU.",
    )

st.subheader("3. Marketplace return files — upload any combination")
st.caption("Each uploader accepts .xls, .xlsx, or .zip")

mcol1, mcol2, mcol3 = st.columns(3)
with mcol1:
    shopee_raw = st.file_uploader(
        "🛍️ Shopee return export",
        type=["xls", "xlsx", "zip"],
        key="shopee",
        help="Direct .xls/.xlsx or a Seller Centre .zip download",
    )
with mcol2:
    lazada_raw = st.file_uploader(
        "🟠 Lazada return export",
        type=["xls", "xlsx", "zip"],
        key="lazada",
        help="Direct .xls/.xlsx or a Seller Centre .zip download",
    )
with mcol3:
    tiktok_raw = st.file_uploader(
        "🎵 TikTok return export",
        type=["xls", "xlsx", "zip"],
        key="tiktok",
        help="Direct .xls/.xlsx or a Seller Centre .zip download",
    )

st.divider()

run = st.button("▶️ Run reconciliation", type="primary", disabled=tracker_file is None)

if tracker_file is None:
    st.info("Upload the tracker workbook to get started.")

if run:
    if not any([shopee_raw, lazada_raw, tiktok_raw]):
        st.warning("Upload at least one marketplace return file before running.")
        st.stop()

    # Resolve uploads (extract from zip if needed)
    shopee_file  = resolve_upload(shopee_raw)
    lazada_file  = resolve_upload(lazada_raw)
    tiktok_file  = resolve_upload(tiktok_raw)
    tc_buf       = resolve_upload(tc_file) if tc_file else None

    tracker_raw = tracker_file.read()
    tracker_buf = io.BytesIO(tracker_raw)

    with st.spinner("Cross-checking order IDs and building new rows..."):
        try:
            workbook, results = run_reconciliation(
                tracker_file=tracker_buf,
                shopee_file=shopee_file,
                lazada_file=lazada_file,
                tiktok_file=tiktok_file,
                tc_file=tc_buf,
            )
        except Exception as e:
            st.error(f"Something went wrong: {e}")
            st.stop()

    st.success("Reconciliation complete!")

    total_new = sum(len(r.new_rows) for r in results)

    summary_cols = st.columns(max(len(results), 1))
    for col, res in zip(summary_cols, results):
        with col:
            st.metric(
                label=res.marketplace,
                value=f"{len(res.new_rows)} new",
                delta=f"{len(res.already_tracked)} already tracked",
                delta_color="off",
            )

    for res in results:
        for w in res.warnings:
            st.warning(f"**{res.marketplace}**: {w}")

    st.divider()
    st.subheader("New rows added")

    if total_new == 0:
        st.info("No new return orders — everything in the uploaded files is already tracked.")
    else:
        for res in results:
            if not res.new_rows:
                continue
            st.markdown(f"**{res.marketplace}** — {len(res.new_rows)} new row(s)")
            table_data = [
                {
                    "Order ID": r.order_id,
                    "Return Request Date": r.return_request_date,
                    "Ordered Date": r.ordered_date,
                    "Invoice Number": r.invoice_number,
                    "Tracking": r.tracking_number,
                    "SKU": r.sku,
                    "Qty": r.qty,
                    "Return Reason": r.return_reason,
                    "Notes": r.notes,
                }
                for r in res.new_rows
            ]
            st.dataframe(table_data, use_container_width=True)

    # Download
    buf = io.BytesIO()
    workbook.save(buf)
    buf.seek(0)

    orig_name = tracker_file.name.rsplit(".", 1)[0]
    out_name = f"{orig_name}_updated_{datetime.now().strftime('%Y%m%d')}.xlsx"

    st.divider()
    st.download_button(
        "⬇️ Download updated tracker",
        data=buf,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

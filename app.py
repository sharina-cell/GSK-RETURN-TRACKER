"""
app.py
------
Streamlit UI for the Returns Reconciliation tool.

Upload a brand Return Report tracker (xlsx) plus any combination of marketplace
return exports (Shopee, Lazada, TikTok) and an optional TC Order Report (csv).
The app cross-checks which return Order IDs are already tracked, builds the new
rows for anything that isn't, enriches them from the TC Order Report when
possible, and lets you download the updated tracker.
"""

from __future__ import annotations

import io
from datetime import datetime

import streamlit as st

from report_engine import run_reconciliation, ReconciliationResult


st.set_page_config(page_title="Returns Reconciliation", page_icon="🔄", layout="wide")

st.title("🔄 Returns Reconciliation")
st.caption(
    "Cross-check marketplace return files against your brand's Return Report tracker, "
    "and auto-fill new rows for anything that isn't tracked yet."
)

with st.expander("ℹ️ How this works", expanded=False):
    st.markdown(
        """
        1. **Upload your tracker** — the brand's `*Return_report*.xlsx` workbook (e.g. the
           `OMRON Return report - New Format` or `GSK Return report - New Format` sheet).
        2. **Upload marketplace return files** for whichever platforms you have this round —
           Shopee (`.xls`/`.xlsx`), Lazada (`.xlsx`), and/or TikTok (`.xlsx`).
        3. **Optionally upload the TC Order Report** (`.csv`) — used to fill in Invoice Number
           and SKU when the marketplace file itself doesn't have them.
        4. Click **Run reconciliation**. The app tells you, per marketplace, how many return
           Order IDs were already tracked vs. brand-new, then lets you download the updated
           tracker with new rows appended.

        **Note:** columns that require manual confirmation (Return confirmation, Return
        Receiving Date, Fault Description, Dispute, Status, etc.) are always left blank —
        those are filled in later by your ops/warehouse team, never auto-filled here.

        **Brand mismatch safety:** if your marketplace files belong to a different brand
        than the tracker (e.g. uploading GSK return files against an OMRON tracker), the
        Order IDs simply won't match anything and everything will show as "new" — review
        the new-row preview before downloading if counts look unexpectedly high.
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
        help="Used to enrich new rows with Invoice Number and SKU when missing.",
    )

st.subheader("3. Marketplace return files (upload any combination)")
mcol1, mcol2, mcol3 = st.columns(3)
with mcol1:
    shopee_file = st.file_uploader("Shopee return export", type=["xls", "xlsx"], key="shopee")
with mcol2:
    lazada_file = st.file_uploader("Lazada return export", type=["xls", "xlsx"], key="lazada")
with mcol3:
    tiktok_file = st.file_uploader("TikTok return export", type=["xls", "xlsx"], key="tiktok")

st.divider()

run = st.button("▶️ Run reconciliation", type="primary", disabled=tracker_file is None)

if tracker_file is None:
    st.info("Upload the tracker workbook to get started.")

if run:
    if not any([shopee_file, lazada_file, tiktok_file]):
        st.warning("Upload at least one marketplace return file before running.")
        st.stop()

    with st.spinner("Cross-checking order IDs and building new rows..."):
        try:
            workbook, results = run_reconciliation(
                tracker_file=tracker_file,
                shopee_file=shopee_file,
                lazada_file=lazada_file,
                tiktok_file=tiktok_file,
                tc_file=tc_file,
            )
        except Exception as e:
            st.error(f"Something went wrong: {e}")
            st.stop()

    st.success("Reconciliation complete!")

    total_new = sum(len(r.new_rows) for r in results)
    total_tracked = sum(len(r.already_tracked) for r in results)

    summary_cols = st.columns(len(results) if results else 1)
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
        st.info("No new return orders found — everything in the uploaded files was already tracked.")
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
                    "Tracking Number": r.tracking_number,
                    "SKU": r.sku,
                    "Qty": r.qty,
                    "Return Reason": r.return_reason,
                    "Notes": r.notes,
                }
                for r in res.new_rows
            ]
            st.dataframe(table_data, use_container_width=True)

    # Prepare the download
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

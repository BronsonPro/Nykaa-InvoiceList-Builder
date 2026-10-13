"""
Nykaa InvoiceList builder (Streamlit)

  Master file(s)  -> Invoice MRP (by EAN)
  PO file         -> columns A-J
  Sales invoice   -> columns K-V

Run locally:   streamlit run app.py
"""
import pandas as pd
import streamlit as st

from converter import (auto_match, build_invoice_list, invoice_values, load_master, norm_ean,
                       output_filename, parse_po, parse_sales_invoice, po_to_rows, suggest)

st.set_page_config(page_title="Nykaa InvoiceList", page_icon="🧾", layout="wide")
st.title("🧾 Nykaa InvoiceList builder")
st.caption("Upload the Master file, the Nykaa PO and the sales invoice, then download "
           "InvoiceList_<PO Number>.xlsx ready for upload.")

# ---- uploads ------------------------------------------------------------------
u1, u2, u3 = st.columns(3)
master_files = u1.file_uploader("1 · Master file(s) – EAN & MRP", type=["xlsx", "xls"],
                                accept_multiple_files=True)
po_file = u2.file_uploader("2 · Nykaa PO (PoAlpha_*.xls)", type=["xls", "xlsx", "xml"])
inv_file = u3.file_uploader("3 · Sales invoice", type=["xlsx", "xls"])

# ---- all three files are required -------------------------------------------
missing = [n for n, f in [("Master file", master_files), ("Nykaa PO", po_file),
                          ("Sales invoice", inv_file)] if not f]
if missing:
    st.info("Upload all three files to continue. Still needed: **" + ", ".join(missing) + "**")
    st.stop()

master_mrp, msgs = load_master([(f.name, f.getvalue()) for f in master_files])
with u1:
    for m in msgs:
        (st.caption if "loaded" in m else st.warning)(m)
if not master_mrp:
    st.error("No EAN/MRP data found in the Master file(s). It needs columns named 'EAN' and 'MRP'.")
    st.stop()

try:
    po = parse_po(po_file.getvalue())
except Exception as e:
    st.error(f"Couldn't read the PO: {e}")
    st.stop()
try:
    inv = parse_sales_invoice(inv_file.getvalue())
except Exception as e:
    st.error(f"Couldn't read the sales invoice: {e}")
    st.stop()

# ---- summary ------------------------------------------------------------------
st.divider()
c = st.columns(4)
c[0].metric("PO Number", po.po_number)
c[1].metric("PO lines / units", f"{len(po.lines)} / {sum(l.qty for l in po.lines)}")
c[2].metric(f"Invoice No · dated {inv.invoice_date}", inv.invoice_no)
c[3].metric("Invoice lines / units", f"{len(inv.lines)} / {int(sum(l.qty for l in inv.lines))}")
if inv.po_number and inv.po_number.upper() != po.po_number.upper():
    st.error(f"The sales invoice is for PO **{inv.po_number}**, but the PO file is "
             f"**{po.po_number}**. Upload the matching files.")
    st.stop()

missing_vsku = sum(1 for l in po.lines if not l.vendor_sku)
fallback = "blank"
if missing_vsku:
    choice = st.radio(
        f"Vendor SKU Code is empty for {missing_vsku} line(s) in the PO. Fill it with:",
        ["Leave blank", "SKU code (Item Code)", "EAN"], horizontal=True)
    fallback = {"Leave blank": "blank", "SKU code (Item Code)": "sku", "EAN": "ean"}[choice]
po_rows = po_to_rows(po, fallback)

# ---- match invoice lines to PO lines ------------------------------------------
pairs, un_inv, un_po = auto_match(po, inv)
st.subheader("Matching")
st.write(f"✅ **{len(pairs)}** of {len(inv.lines)} invoice lines matched to the PO by EAN.")

SKIP = "— Not in PO (leave out) —"
po_label = {i: f"Line {po.lines[i].line_no} · {po.lines[i].product_name[:60]} · EAN {po.lines[i].ean}"
            for i in un_po}
if un_inv:
    st.warning(f"{len(un_inv)} invoice line(s) have an EAN that isn't in the PO. "
               "Pick the PO line each one belongs to (best guess pre-selected):")
    for j in un_inv:
        il = inv.lines[j]
        best, score = suggest(il, po, un_po)
        options = [SKIP] + un_po
        default = options.index(best) if best is not None and score >= 0.6 else 0
        pick = st.selectbox(
            f"Invoice: **{il.name}** · EAN {il.ean} · Qty {int(il.qty)}",
            options, index=default, key=f"m_{j}",
            format_func=lambda o: o if o == SKIP else po_label[o])
        if pick != SKIP:
            pairs[j] = pick

chosen = list(pairs.values())
dupes = {i for i in chosen if chosen.count(i) > 1}
if dupes:
    st.error("The same PO line is picked for more than one invoice line: " +
             ", ".join(f"line {po.lines[i].line_no}" for i in dupes))
    st.stop()

left_po = [i for i in range(len(po.lines)) if i not in set(chosen)]
if left_po:
    with st.expander(f"{len(left_po)} PO line(s) not invoiced – they will be left out of the file"):
        st.dataframe(pd.DataFrame([{"Line": po.lines[i].line_no, "SKU": po.lines[i].sku_code,
                                    "Product": po.lines[i].product_name, "PO Qty": po.lines[i].qty}
                                   for i in left_po]), hide_index=True, use_container_width=True)

# ---- build rows -----------------------------------------------------------------
rows, no_mrp, over_qty = [], [], []
for j, i in sorted(pairs.items(), key=lambda kv: po.lines[kv[1]].line_no):
    il, pl = inv.lines[j], po.lines[i]
    mrp = master_mrp.get(il.ean, master_mrp.get(norm_ean(pl.ean)))
    if mrp is None:
        no_mrp.append(f"line {pl.line_no} (EAN {il.ean})")
    if il.qty > pl.qty:
        over_qty.append(f"line {pl.line_no} ({int(il.qty)} > {pl.qty})")
    rows.append({**po_rows[i], **invoice_values(il, inv.is_igst, mrp),
                 "invoice_no": inv.invoice_no, "invoice_date": inv.invoice_date})

if no_mrp:
    st.warning(f"MRP not found in the Master file for {len(no_mrp)} line(s), left blank: "
               + ", ".join(no_mrp))
if over_qty:
    st.warning("Invoice qty is more than PO qty on " + ", ".join(over_qty))

# ---- preview + download ------------------------------------------------------
df = pd.DataFrame(rows)
df["invoice_mrp"] = pd.to_numeric(df["invoice_mrp"], errors="coerce")  # blank -> empty cell in preview
st.subheader("Preview")
st.dataframe(
    df[["line_no", "sku_code", "product_name", "qty", "invoice_qty", "invoice_mrp",
        "invoice_cost", "tax_rate", "line_amount"]].rename(columns={
        "line_no": "Line", "sku_code": "SKU code", "product_name": "Product", "qty": "PO Qty",
        "invoice_qty": "Inv Qty", "invoice_mrp": "Inv MRP", "invoice_cost": "Inv Cost",
        "tax_rate": "Tax %", "line_amount": "Line Amount"}).round(2),
    hide_index=True, use_container_width=True, height=380)

t = st.columns(4)
t[0].metric("Lines / units in file", f"{len(df)} / {int(df['invoice_qty'].sum())}")
t[1].metric("Taxable", f"₹{(df['invoice_qty'] * df['invoice_cost']).sum():,.2f}")
t[2].metric("Tax", f"₹{df['line_tax'].sum():,.2f}")
t[3].metric("Total", f"₹{df['line_amount'].sum():,.2f}")

st.download_button(
    f"⬇️ Download {output_filename(po.po_number)}",
    data=build_invoice_list(rows),
    file_name=output_filename(po.po_number),
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

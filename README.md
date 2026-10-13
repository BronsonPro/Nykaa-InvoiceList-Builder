# Nykaa InvoiceList builder

Upload **all three** files and download `InvoiceList_<PONumber>.xlsx` in Nykaa's upload format:

1. **Master file(s)**: Excel with columns `EAN` and `MRP` (gives Invoice MRP). More than one file is allowed.
2. **Nykaa PO**: `PoAlpha_*.xls` (gives columns A–J)
3. **Sales invoice**: Tally GST invoice `.xlsx` (gives columns K–V)

Invoice lines are matched to PO lines by EAN. If an invoice EAN isn't in the PO (for example, a wrong
EAN on Nykaa's PO), the app suggests the right PO line by product name and lets you confirm it.
PO lines that weren't invoiced are left out, and the app lists them.
The download only appears once all three files are uploaded and they belong to the same PO.

## Files
| File | What it is |
|---|---|
| `app.py` | The web app (Streamlit) |
| `converter.py` | Reads the PO, sales invoice and master file, and writes the InvoiceList (no UI) |
| `invoice_template.xlsx` | Nykaa InvoiceList template (Instructions + InvoiceList sheets) |
| `requirements.txt` | Python packages |

## Share it as a link (free, about 5 minutes)
1. Create a new GitHub repository (it can be **private**), e.g. `Nykaa-InvoiceList-Builder`, and upload
   these files to it: `app.py`, `converter.py`, `invoice_template.xlsx`, `requirements.txt`, `README.md`, `.gitignore`.
2. Go to **https://share.streamlit.io** and sign in with GitHub.
3. Click **Create app → Deploy a public app from GitHub**, then pick the repo, branch `main`
   and main file `app.py`. Choose a URL such as `nykaa-invoicelist-builder.streamlit.app`, then click **Deploy**.
4. Send that URL to your teammate. To keep it private, open the app's **Settings → Sharing**,
   choose "Only specific people can view this app", and add your teammate's email.

Whenever you push a change to GitHub, the app updates automatically.

## Run on your own PC instead
```
pip install -r requirements.txt
streamlit run app.py
```

## How fields are mapped (columns A–J, from the PO)
| InvoiceList column | PO field |
|---|---|
| A PO code | PO No (top of PO) |
| B PO Line No | `#` |
| C SKU code | Item Code |
| D Product Name | SKU Name (the "Colour : / Size :" suffix is removed) |
| E Vendor SKU Code | Vendor Sku (if empty: blank, Item Code or EAN, chosen in the app) |
| F EAN Code | EAN |
| G HSN | HSN |
| H PO Product MRP | MRP |
| I PO Product Cost | Unit Price |
| J PO Product Qty | Qty |

## Columns K–V (from the sales invoice; same formulas as the Nykaa Invoice Filler)
| InvoiceList column | Source |
|---|---|
| K Invoice No | "Invoice No. :" on the sales invoice |
| L Invoice Date | "Date :" on the sales invoice, as DD/MM/YYYY |
| M Invoice Qty | QTY Nos |
| N Invoice MRP | MRP from the Master file, looked up by EAN (left blank if not found, with a warning) |
| O Invoice Cost | Rate × (1 − Disc%/100) ÷ (1 + GST%/100) |
| P Discount | Disc% ÷ 100 (shown as %) |
| Q Tax Rate | GST% |
| R Line Tax Amount | Invoice Qty × Invoice Cost × Tax Rate/100 |
| S, T CGST / SGST | Line Tax ÷ 2 each (0 if the invoice has an IGST Amount column) |
| U IGST | Line Tax if the invoice has an IGST Amount column, otherwise 0 |
| V Line Amount | Invoice Qty × Invoice Cost + Line Tax |

The app stops if the invoice's PO No doesn't match the PO file. It warns if an MRP is missing from the
Master file or if an invoice qty is above the PO qty.

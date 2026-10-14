# Nykaa InvoiceList builder

Upload **all three** files and download `InvoiceList_<PONumber>.xlsx` in Nykaa's upload format:

1. **Master file(s)**: Excel with an `EAN Code` (or `EAN`) column and an `MRP` column (gives Invoice MRP).
   More than one file is allowed, and every sheet is read. A `Nykaa Code` column is used to match lines whose EAN differs (see below).
2. **Nykaa PO**: `PoAlpha_*.xls` (gives columns A–J)
3. **Sales invoice**: Tally GST invoice `.xlsx` (gives columns K–V)

Invoice lines are matched to PO lines in this order:
1. **Same EAN** in the PO and the invoice.
2. **Via the Master file:** some Nykaa POs show the SKU code (Nykaa Code) in the EAN column, or a wrong
   EAN. The PO's Item Code is looked up as a `Nykaa Code` in the Master file, and the EAN listed there
   is matched to the invoice.
3. **Manual:** anything still unmatched is shown with a best-guess PO line, chosen by product name, for
   you to confirm. Invoiced items that end up left out are flagged in red.

PO lines that weren't invoiced are left out of the file, and the app lists them.
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
| N Invoice MRP | MRP from the Master file: matched by invoice EAN, then PO EAN, then PO Item Code = `Nykaa Code` (left blank with a warning if none match) |
| O Invoice Cost | Rate × (1 − Disc%/100) ÷ (1 + GST%/100) |
| P Discount | Disc% ÷ 100 (shown as %) |
| Q Tax Rate | GST% |
| R Line Tax Amount | Invoice Qty × Invoice Cost × Tax Rate/100 |
| S, T CGST / SGST | Line Tax ÷ 2 each (0 if the invoice has an IGST Amount column) |
| U IGST | Line Tax if the invoice has an IGST Amount column, otherwise 0 |
| V Line Amount | Invoice Qty × Invoice Cost + Line Tax |

The app stops if the invoice's PO No doesn't match the PO file. It warns if an MRP is missing from the
Master file or if an invoice qty is above the PO qty.

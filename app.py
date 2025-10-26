from flask import Flask, render_template, request, send_file
import pandas as pd
import xml.etree.ElementTree as ET
import os

app = Flask(__name__)

# === Home page ===
@app.route('/')
def index():
    return render_template('index.html')

# === Download template ===
@app.route('/download-template')
def download_template():
    return send_file("Template_S&P.xlsx", as_attachment=True)

# === Convert Excel to Tally XML ===
@app.route('/convert', methods=['POST'])
def convert():
    file = request.files.get('file')
    if not file:
        return "No file uploaded", 400

    df = pd.read_excel(file)
    df = df.fillna("")
    for col in ["Quantity", "Rate per Piece", "Final Amount", "CGST Amount", "SGST Amount", "IGST Amount"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # === XML structure setup ===
    envelope = ET.Element("ENVELOPE")
    header = ET.SubElement(envelope, "HEADER")
    ET.SubElement(header, "TALLYREQUEST").text = "Import Data"

    body = ET.SubElement(envelope, "BODY")
    importdata = ET.SubElement(body, "IMPORTDATA")
    requestdesc = ET.SubElement(importdata, "REQUESTDESC")
    ET.SubElement(requestdesc, "REPORTNAME").text = "Vouchers"

    staticvars = ET.SubElement(requestdesc, "STATICVARIABLES")
    ET.SubElement(staticvars, "SVCURRENTCOMPANY").text = df["Company Name"].iloc[0] if "Company Name" in df.columns else "Default Company"

    requestdata = ET.SubElement(importdata, "REQUESTDATA")

    # === Group by Invoice ===
    for inv, group in df.groupby("Invoice Number"):
        date_val = pd.to_datetime(group["Date"].iloc[0])
        date_str = date_val.strftime("%Y%m%d")
        vtype = str(group["Voucher Type"].iloc[0]).strip().lower()

        vch = ET.SubElement(requestdata, "TALLYMESSAGE")
        voucher = ET.SubElement(vch, "VOUCHER", {
            "VCHTYPE": str(group["Voucher Type"].iloc[0]),
            "ACTION": "Create",
            "OBJVIEW": "Invoice Voucher View"
        })

        ET.SubElement(voucher, "DATE").text = date_str
        ET.SubElement(voucher, "VOUCHERNUMBER").text = str(inv)
        ET.SubElement(voucher, "PARTYNAME").text = str(group["Party Name"].iloc[0])
        ET.SubElement(voucher, "VOUCHERTYPENAME").text = str(group["Voucher Type"].iloc[0])
        ET.SubElement(voucher, "ISINVOICE").text = "Yes"
        ET.SubElement(voucher, "NARRATION").text = str(group["Narration"].iloc[0]) if "Narration" in group.columns else ""

        # === Totals ===
        total_items = 0
        for _, row in group.iterrows():
            amt = row["Final Amount"] if row["Final Amount"] > 0 else row["Quantity"] * row["Rate per Piece"]
            total_items += amt
        total_cgst = group["CGST Amount"].sum()
        total_sgst = group["SGST Amount"].sum()
        total_igst = group["IGST Amount"].sum()
        total_amount = total_items + total_cgst + total_sgst + total_igst

        # === PARTY LEDGER ENTRY ===
        party_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
        ET.SubElement(party_entry, "LEDGERNAME").text = str(group["Party Name"].iloc[0])
        ET.SubElement(party_entry, "ISPARTYLEDGER").text = "Yes"

        if vtype == "sales":
            ET.SubElement(party_entry, "ISDEEMEDPOSITIVE").text = "Yes"   # Debit
            ET.SubElement(party_entry, "AMOUNT").text = f"{-total_amount:.2f}"
        else:  # Purchase
            ET.SubElement(party_entry, "ISDEEMEDPOSITIVE").text = "No"    # Credit
            ET.SubElement(party_entry, "AMOUNT").text = f"{total_amount:.2f}"

        # === STOCK ITEMS ===
        for _, row in group.iterrows():
            stock_amt = row["Final Amount"] if row["Final Amount"] > 0 else row["Quantity"] * row["Rate per Piece"]

            stock_entry = ET.SubElement(voucher, "ALLINVENTORYENTRIES.LIST")
            ET.SubElement(stock_entry, "STOCKITEMNAME").text = str(row["Stock Item Name"])
            ET.SubElement(stock_entry, "RATE").text = str(row["Rate per Piece"])
            ET.SubElement(stock_entry, "BILLEDQTY").text = str(row["Quantity"])
            ET.SubElement(stock_entry, "ACTUALQTY").text = str(row["Quantity"])

            if vtype == "sales":
                ET.SubElement(stock_entry, "ISDEEMEDPOSITIVE").text = "No"
                ET.SubElement(stock_entry, "AMOUNT").text = f"{stock_amt:.2f}"
            else:
                ET.SubElement(stock_entry, "ISDEEMEDPOSITIVE").text = "Yes"
                ET.SubElement(stock_entry, "AMOUNT").text = f"{-stock_amt:.2f}"

            acc_alloc = ET.SubElement(stock_entry, "ACCOUNTINGALLOCATIONS.LIST")
            ET.SubElement(acc_alloc, "LEDGERNAME").text = str(row["Ledger Name"])
            if vtype == "sales":
                ET.SubElement(acc_alloc, "ISDEEMEDPOSITIVE").text = "No"
                ET.SubElement(acc_alloc, "AMOUNT").text = f"{stock_amt:.2f}"
            else:
                ET.SubElement(acc_alloc, "ISDEEMEDPOSITIVE").text = "Yes"
                ET.SubElement(acc_alloc, "AMOUNT").text = f"{-stock_amt:.2f}"

        # === GST LEDGERS ===
        gst_ledgers = [
            ("CGST Amount", "CGST LEDGER NAME"),
            ("SGST Amount", "SGST LEDGER NAME"),
            ("IGST Amount", "IGST LEDGER NAME")
        ]
        for amt_col, name_col in gst_ledgers:
            gst_amt = group[amt_col].sum() if amt_col in group.columns else 0
            ledger_name = str(group[name_col].iloc[0]).strip() if name_col in group.columns else ""
            if gst_amt > 0 and ledger_name:
                gst_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
                ET.SubElement(gst_entry, "LEDGERNAME").text = ledger_name
                if vtype == "sales":
                    ET.SubElement(gst_entry, "ISDEEMEDPOSITIVE").text = "No"
                    ET.SubElement(gst_entry, "AMOUNT").text = f"{gst_amt:.2f}"
                else:
                    ET.SubElement(gst_entry, "ISDEEMEDPOSITIVE").text = "Yes"
                    ET.SubElement(gst_entry, "AMOUNT").text = f"{-gst_amt:.2f}"

    # === Write output ===
    output_path = "S&P.xml"
    tree = ET.ElementTree(envelope)
    tree.write(output_path, encoding="utf-8", xml_declaration=True)

    return send_file(output_path, as_attachment=True)

# === Render Compatibility ===
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)

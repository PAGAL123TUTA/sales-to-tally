from flask import Flask, render_template, request, send_file
import pandas as pd
from datetime import datetime
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

    # Clean numeric columns
    for col in ["Quantity", "Rate per Piece", "Final Amount", "CGST Amount", "SGST Amount", "IGST Amount"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df = df.fillna("")

    # === XML structure setup ===
    envelope = ET.Element("ENVELOPE")
    header = ET.SubElement(envelope, "HEADER")
    tallyrequest = ET.SubElement(header, "TALLYREQUEST")
    tallyrequest.text = "Import Data"

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

        # === Calculate totals ===
        total_item_amt = 0
        for _, row in group.iterrows():
            amt = row["Final Amount"] if row["Final Amount"] > 0 else row["Quantity"] * row["Rate per Piece"]
            total_item_amt += amt

        total_cgst = group["CGST Amount"].sum()
        total_sgst = group["SGST Amount"].sum()
        total_igst = group["IGST Amount"].sum()
        total_credit = total_item_amt + total_cgst + total_sgst + total_igst

        vtype = str(group["Voucher Type"].iloc[0]).lower()
        party_debit = True if vtype != "purchase" else False

        # === Party Ledger ===
        party_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
        ET.SubElement(party_entry, "LEDGERNAME").text = str(group["Party Name"].iloc[0])
        ET.SubElement(party_entry, "ISDEEMEDPOSITIVE").text = "Yes" if party_debit else "No"
        ET.SubElement(party_entry, "AMOUNT").text = f"{-total_credit:.2f}" if party_debit else f"{total_credit:.2f}"
        ET.SubElement(party_entry, "ISPARTYLEDGER").text = "Yes"

        # === GST Ledgers ===
        gst_columns = [
            ("CGST Amount", "CGST LEDGER NAME"),
            ("SGST Amount", "SGST LEDGER NAME"),
            ("IGST Amount", "IGST LEDGER NAME")
        ]
        for amt_col, name_col in gst_columns:
            total_amt = group[amt_col].sum()
            ledger_name = str(group[name_col].iloc[0]).strip() if name_col in group.columns else ""
            if ledger_name and ledger_name.lower() != "nan" and total_amt > 0:
                gst_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
                ET.SubElement(gst_entry, "LEDGERNAME").text = ledger_name
                if vtype == "sales":
                    ET.SubElement(gst_entry, "ISDEEMEDPOSITIVE").text = "No"
                    ET.SubElement(gst_entry, "AMOUNT").text = f"{total_amt:.2f}"
                else:  # purchase
                    ET.SubElement(gst_entry, "ISDEEMEDPOSITIVE").text = "Yes"
                    ET.SubElement(gst_entry, "AMOUNT").text = f"{-total_amt:.2f}"

        # === Stock Items ===
        for _, row in group.iterrows():
            item_amt = row["Final Amount"] if row["Final Amount"] > 0 else row["Quantity"] * row["Rate per Piece"]
            stock_entry = ET.SubElement(voucher, "ALLINVENTORYENTRIES.LIST")
            ET.SubElement(stock_entry, "STOCKITEMNAME").text = str(row["Stock Item Name"])
            ET.SubElement(stock_entry, "ISDEEMEDPOSITIVE").text = "No"
            ET.SubElement(stock_entry, "RATE").text = str(row["Rate per Piece"])
            ET.SubElement(stock_entry, "AMOUNT").text = f"{item_amt:.2f}"
            ET.SubElement(stock_entry, "BILLEDQTY").text = str(row["Quantity"])
            ET.SubElement(stock_entry, "ACTUALQTY").text = str(row["Quantity"])

            acc_alloc = ET.SubElement(stock_entry, "ACCOUNTINGALLOCATIONS.LIST")
            ET.SubElement(acc_alloc, "LEDGERNAME").text = str(row["Ledger Name"])
            if vtype == "sales":
                ET.SubElement(acc_alloc, "ISDEEMEDPOSITIVE").text = "No"
                ET.SubElement(acc_alloc, "AMOUNT").text = f"{item_amt:.2f}"
            else:
                ET.SubElement(acc_alloc, "ISDEEMEDPOSITIVE").text = "Yes"
                ET.SubElement(acc_alloc, "AMOUNT").text = f"{-item_amt:.2f}"

    output_path = "S&P.xml"
    tree = ET.ElementTree(envelope)
    tree.write(output_path, encoding="utf-8", xml_declaration=True)

    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)

from flask import Flask, render_template, request, send_file
import pandas as pd
import xml.etree.ElementTree as ET
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download-template')
def download_template():
    return send_file("Template_S&P.xlsx", as_attachment=True)

@app.route('/convert', methods=['POST'])
def convert():
    file = request.files.get('file')
    if not file:
        return "No file uploaded", 400

    df = pd.read_excel(file)
    df = df.fillna("")
    for col in ["Quantity", "Rate per Piece", "Final Amount", "CGST Amount", "SGST Amount", "IGST Amount"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    envelope = ET.Element("ENVELOPE")
    header = ET.SubElement(envelope, "HEADER")
    ET.SubElement(header, "TALLYREQUEST").text = "Import Data"

    body = ET.SubElement(envelope, "BODY")
    importdata = ET.SubElement(body, "IMPORTDATA")
    requestdesc = ET.SubElement(importdata, "REQUESTDESC")
    ET.SubElement(requestdesc, "REPORTNAME").text = "Vouchers"

    staticvars = ET.SubElement(requestdesc, "STATICVARIABLES")
    ET.SubElement(staticvars, "SVCURRENTCOMPANY").text = df.get("Company Name", pd.Series(["Default Company"])).iloc[0]

    requestdata = ET.SubElement(importdata, "REQUESTDATA")

    for inv, group in df.groupby("Invoice Number"):
        vtype = str(group["Voucher Type"].iloc[0]).strip().lower()
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
        ET.SubElement(voucher, "NARRATION").text = str(group.get("Narration", [""]).iloc[0])

        total_amount = 0

        # === Stock Items ===
        for _, row in group.iterrows():
            amt = row["Final Amount"] if row["Final Amount"] > 0 else row["Quantity"] * row["Rate per Piece"]
            total_amount += amt

            stock_entry = ET.SubElement(voucher, "ALLINVENTORYENTRIES.LIST")
            ET.SubElement(stock_entry, "STOCKITEMNAME").text = str(row["Stock Item Name"])
            ET.SubElement(stock_entry, "ISDEEMEDPOSITIVE").text = "No" if vtype == "sales" else "Yes"
            ET.SubElement(stock_entry, "RATE").text = str(row["Rate per Piece"])
            ET.SubElement(stock_entry, "AMOUNT").text = f"{amt if vtype == 'sales' else -amt:.2f}"
            ET.SubElement(stock_entry, "BILLEDQTY").text = str(row["Quantity"])
            ET.SubElement(stock_entry, "ACTUALQTY").text = str(row["Quantity"])

            acc_alloc = ET.SubElement(stock_entry, "ACCOUNTINGALLOCATIONS.LIST")
            ET.SubElement(acc_alloc, "LEDGERNAME").text = str(row["Ledger Name"])
            ET.SubElement(acc_alloc, "ISDEEMEDPOSITIVE").text = "No" if vtype == "sales" else "Yes"
            ET.SubElement(acc_alloc, "AMOUNT").text = f"{amt if vtype == 'sales' else -amt:.2f}"

        # === GST Entries ===
        total_gst = 0
        gst_map = [
            ("CGST Amount", "CGST LEDGER NAME"),
            ("SGST Amount", "SGST LEDGER NAME"),
            ("IGST Amount", "IGST LEDGER NAME")
        ]
        for amt_col, name_col in gst_map:
            gst_total = group[amt_col].sum()
            gst_name = str(group[name_col].iloc[0]) if name_col in group.columns else ""
            if gst_total and gst_name and gst_name.lower() != "nan":
                gst_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
                ET.SubElement(gst_entry, "LEDGERNAME").text = gst_name
                ET.SubElement(gst_entry, "ISDEEMEDPOSITIVE").text = "No" if vtype == "sales" else "Yes"
                ET.SubElement(gst_entry, "AMOUNT").text = f"{gst_total if vtype == 'sales' else -gst_total:.2f}"
                total_gst += gst_total

        total_credit = total_amount + total_gst

        # === Party Ledger Entry (auto-balanced)
        party_entry = ET.SubElement(voucher, "LEDGERENTRIES.LIST")
        ET.SubElement(party_entry, "LEDGERNAME").text = str(group["Party Name"].iloc[0])
        ET.SubElement(party_entry, "ISPARTYLEDGER").text = "Yes"

        if vtype == "sales":
            ET.SubElement(party_entry, "ISDEEMEDPOSITIVE").text = "Yes"
            ET.SubElement(party_entry, "AMOUNT").text = f"{-total_credit:.2f}"
        else:  # Purchase
            ET.SubElement(party_entry, "ISDEEMEDPOSITIVE").text = "No"
            ET.SubElement(party_entry, "AMOUNT").text = f"{total_credit:.2f}"

    tree = ET.ElementTree(envelope)
    output_path = "S&P.xml"
    tree.write(output_path, encoding="utf-8", xml_declaration=True)
    return send_file(output_path, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)

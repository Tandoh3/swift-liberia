import streamlit as st
import pandas as pd
import io
import docx  # pip install python-docx

# Function to read .docx
def read_docx(file) -> str:
    doc = docx.Document(file)
    return "\n".join([para.text for para in doc.paragraphs])

# SWIFT Parser
def parse_swift_message(text: str):
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    bank_name = lines[0]
    currency_code = lines[1]

    header_info = {}
    transactions = []
    opening_balance = {}
    closing_balance = {}

    current_txn = {}
    i = 0
    while i < len(lines):
        line = lines[i]

        if line.startswith("F20:"):
            header_info["Transaction Reference"] = lines[i + 1]
            i += 1

        elif line.startswith("F25:"):
            header_info["Account ID"] = lines[i + 1]
            i += 1

        elif line.startswith("F28C:"):
            header_info["Statement Number"] = lines[i + 1].split(":")[-1].strip()
            header_info["Sequence Number"] = lines[i + 2].split(":")[-1].replace("/", "").strip()
            i += 2

        elif line.startswith("F60F:"):
            opening_balance = {
                "Type": "Opening Balance",
                "Amount": lines[i + 4].split("#")[0].split(":")[-1].replace(",", ".").strip()
            }
            i += 4

        elif line.startswith("F61:"):
            current_txn = {
                "Value Date": lines[i + 1].split(":")[-1].strip()[2:10], 
                "Entry Date": lines[i + 2].split(":")[-1].strip()[:-3], 
                "Debit/Credit": lines[i + 3].split(":")[-1].strip(),
                "Funds Code": lines[i + 4].split(":")[-1].strip(),
                "Amount": lines[i + 5].split("#")[0].split(":")[-1].replace(",", ".").strip(),
                "Transaction Type": lines[i + 6].split(":")[-1].strip(),
                "ID Code": lines[i + 7].split(":")[-1].strip(),
                "Account Owner Ref": lines[i + 8].split(":")[-1].strip(),
                "Servicing Inst Ref": lines[i + 9].split(":")[-1].replace("//", "").strip(),
                "B/O": lines[i + 10].split(":")[-1].strip()
            }
            i += 10

        elif line.startswith("F86:") and current_txn:
            f86_data = []
            j = i + 1
            while j < len(lines) and not lines[j].startswith("F61:") and not lines[j].startswith("F62M:"):
                f86_data.append(lines[j])
                j += 1
            current_txn["Narrative"] = " ".join(f86_data)
            transactions.append(current_txn)
            current_txn = {}
            i = j - 1

        elif line.startswith("F62M:"):
            closing_balance = {
                "Type": "Closing Balance",
                "Amount": lines[i + 4].split("#")[0].split(":")[-1].replace(",", ".").strip()
            }
            i += 4

        i += 1

    return bank_name, currency_code, opening_balance, closing_balance, transactions

# Streamlit UI
st.set_page_config(page_title="SWIFT Extractor", layout="centered")
st.title("SWIFT Statement Extractor")

with st.expander("📄 Sample SWIFT Message Format (click to expand)", expanded=False):
    st.markdown("### ✍️ Format Your `.txt` and Upload Like This:")
    
    sample_text = """CITI NY
USD
F20: Transaction Reference Number
TTS2514301573952
F25: Account Identification - Account
36922859
F28C: Statement Number/Sequence Number
Statement Number:0909
Sequence Number:/00001
F60F: Opening Balance - D/C Mark - Date - Currency - Amount
DCMark: D/C Mark:C
Date:2505232025 May 23
Currency:USDUS DOLLAR
Amount:7483232,84#7483232,84#
"""

    st.code(sample_text, language="text")

    st.markdown("✅ You can download this as a `.txt` file to test the upload:")

    st.download_button(
        label="📄 Download Sample SWIFT File (.txt)",
        data=sample_text,
        file_name="sample_swift.txt",
        mime="text/plain"
    )



uploaded_file = st.file_uploader("📄 Upload a .txt file", type=["txt", "docx"])

if uploaded_file:
    if uploaded_file.name.endswith(".txt"):
        raw_text = uploaded_file.read().decode("utf-8", errors="ignore")
    else:
        raw_text = read_docx(uploaded_file)

    bank_name, currency_code, opening_balance, closing_balance, transactions = parse_swift_message(raw_text)

    st.success("SWIFT message parsed successfully.")

    output = io.BytesIO()

    balance_df = pd.DataFrame([opening_balance, closing_balance])
    txn_df = pd.DataFrame(transactions)

    # Write to Excel with formatting
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("SWIFT Report")
        writer.sheets["SWIFT Report"] = worksheet

        # Define bordered format
        bordered = workbook.add_format({
            "border": 1,
            "align": "center",
            "valign": "vcenter"
        })

        # Write Opening and Closing Balance with border
        for row_idx, (_, row) in enumerate(balance_df.iterrows()):
            worksheet.write(row_idx, 0, row["Type"], bordered)
            worksheet.write(row_idx, 1, float(row["Amount"]), bordered)

        # Start transactions after a blank row
        txn_start = len(balance_df) + 1
        txn_df.to_excel(writer, index=False, startrow=txn_start, sheet_name="SWIFT Report")

        # Optional: Autofit column widths
        for col_idx, col in enumerate(txn_df.columns):
            column_len = max(txn_df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(col_idx, col_idx, column_len)


    st.subheader("📋 Transactions Preview")
    st.dataframe(txn_df, use_container_width=True)

    st.download_button(
        label="⬇️ Download Excel",
        data=output.getvalue(),
        file_name="swift_statement.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Please upload a SWIFT `.txt` document.")

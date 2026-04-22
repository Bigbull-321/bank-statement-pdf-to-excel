import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import cv2
import pytesseract
import numpy as np
from pdf2image import convert_from_bytes
import os

# -------------------------------
# Config paths for OCR
# -------------------------------
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
poppler_path = r"C:\Users\Admin\Downloads\Release-25.11.0-0\poppler-25.11.0\Library\bin"

# -------------------------------
# Streamlit page config
# -------------------------------
st.set_page_config(page_title="Bank PDF to Excel Converter", layout="centered")
st.title("Bank Statement PDF → Excel Converter")
st.write("Upload your Canara, Union, Axis, HDFC, Kotak, BOI, ICICI, or SBI Bank PDF statement.")

uploaded_file = st.file_uploader("Drag and Drop PDF Here", type=["pdf"])
pdf_password = st.text_input("Enter PDF Password (Leave blank if not required)", type="password")

# =====================================================
# SAFE PDF LOADER
# =====================================================
def open_pdf(uploaded_file, password=None):
    try:
        pdf_bytes = uploaded_file.read()
        pdf_buffer = io.BytesIO(pdf_bytes)
        pdf = pdfplumber.open(pdf_buffer, password=password)
        return pdf
    except Exception as e:
        raise Exception(f"❌ Could not open PDF. Error: {e}")

# =====================================================
# TABLE EXTRACTION
# =====================================================
def extract_tables(pdf):
    all_rows = []
    for page in pdf.pages:
        tables = page.extract_tables()
        if tables:
            for table in tables:
                for row in table:
                    if any(cell is not None and str(cell).strip() != "" for cell in row):
                        all_rows.append(row)
    return all_rows

# =====================================================
# BANK DETECTION
# =====================================================
def detect_bank_type(rows, pdf):
    if rows:
        header_text = " ".join([str(cell).upper() for cell in rows[0] if cell]).strip()
        if "TRANS DATE" in header_text or "REF/CHQ.NO" in header_text:
            return "Canara Bank"
        elif "TRAN ID" in header_text or "UTR" in header_text:
            return "Union Bank"
        elif "TRAN DATE" in header_text or "INIT.BR" in header_text:
            return "Axis Bank"
        elif "SR NO" in header_text and "REMARKS" in header_text:
            return "BOI Bank"
        elif "TRANSACTION DATE" in header_text and "VALUE DATE" in header_text:
            return "Kotak Bank"
        elif "DATE" in header_text and "PARTICULARS" in header_text:
            return "ICICI Bank"

    first_page_text = pdf.pages[0].extract_text() if pdf.pages else ""
    if first_page_text:
        if re.search(r"\d{2}/\d{2}/\d{4}", first_page_text):
            return "HDFC Bank"
        elif re.search(r"\d{2}\s[A-Za-z]{3}\s\d{4}", first_page_text):
            return "Kotak Bank"
        elif "ICICI" in first_page_text.upper():
            return "ICICI Bank"
    return "SBI Bank (OCR)"  # Default to OCR if none detected

# =====================================================
# FIXED SBI OCR EXTRACTION - NO DATE DUPLICATION
# =====================================================
def extract_sbi_transactions(pdf_bytes):
    try:
        pages = convert_from_bytes(pdf_bytes, dpi=300, poppler_path=poppler_path)
    except Exception as e:
        raise Exception(f"❌ Could not convert PDF to images. Error: {e}")

    transactions = []
    date_pat = r"\d{2}-\d{2}-\d{4}"
    amount_pat = r"\d{1,3}(?:,\d{2,3})*\.\d{2}"
    
    for page in pages:
        img = np.array(page)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)[1]
        text = pytesseract.image_to_string(thresh, config='--oem 3 --psm 6')
        
        for line in text.split("\n"):
            line = " ".join(line.split())
            if len(line) < 15: 
                continue

            # Handle BROUGHT FORWARD
            if "BROUGHT FORWARD" in line.upper():
                amounts = re.findall(amount_pat, line)
                if amounts:
                    balance = amounts[-1]
                    if "CR" in line.upper():
                        balance += "CR"
                    elif "DR" in line.upper():
                        balance += "DR"
                    transactions.append(["", "", "BROUGHT FORWARD", "", "", "", balance])
                continue

            # Find dates in the line
            dates = re.findall(date_pat, line)
            if len(dates) >= 2:
                post_date = dates[0]
                value_date = dates[1]
                
                # Find all amounts
                amounts = re.findall(amount_pat, line)
                
                if len(amounts) >= 1:
                    # Last amount is balance
                    balance_amount = amounts[-1]
                    
                    # Check for CR/DR after balance
                    if "CR" in line[line.rfind(balance_amount):].upper():
                        balance = balance_amount + "CR"
                    elif "DR" in line[line.rfind(balance_amount):].upper():
                        balance = balance_amount + "DR"
                    else:
                        balance = balance_amount
                    
                    # Get description - remove all dates and the balance amount
                    description = line
                    
                    # Remove all dates from description
                    for d in dates:
                        description = description.replace(d, "", 1)
                    
                    # Remove balance amount from description
                    description = description.replace(balance_amount, "")
                    
                    # Remove any transaction amounts from description
                    transaction_amounts = amounts[:-1] if len(amounts) > 1 else []
                    for amt in transaction_amounts:
                        description = description.replace(amt, "")
                    
                    # Clean up description - remove extra spaces and any remaining CR/DR
                    description = re.sub(r'\s+', ' ', description).strip()
                    description = re.sub(r'\bCR\b', '', description, flags=re.IGNORECASE)
                    description = re.sub(r'\bDR\b', '', description, flags=re.IGNORECASE)
                    description = re.sub(r'\s+', ' ', description).strip()
                    
                    # Determine debit and credit
                    debit = ""
                    credit = ""
                    
                    if len(amounts) > 1:
                        transaction_amounts = amounts[:-1]
                        
                        if len(transaction_amounts) == 1:
                            # Single transaction amount
                            if transactions:
                                try:
                                    prev_balance = float(re.sub(r'[^\d.]', '', transactions[-1][6].replace(',', '')))
                                    curr_balance = float(balance_amount.replace(',', ''))
                                    
                                    if curr_balance > prev_balance:
                                        credit = transaction_amounts[0]
                                    else:
                                        debit = transaction_amounts[0]
                                except:
                                    # Default to credit if can't determine
                                    credit = transaction_amounts[0]
                            else:
                                credit = transaction_amounts[0]
                        
                        elif len(transaction_amounts) >= 2:
                            # Multiple amounts - first is debit, second is credit
                            debit = transaction_amounts[0]
                            credit = transaction_amounts[1]
                    
                    transactions.append([post_date, value_date, description, "", debit, credit, balance])

    df = pd.DataFrame(transactions, columns=[
        "Post Date", "Value Date", "Description", "Cheque No/Reference", "Debit", "Credit", "Balance"
    ])
    df = df.dropna(how="all").reset_index(drop=True)
    
    return df

# =====================================================
# MAIN PROCESS
# =====================================================
if uploaded_file is not None:
    pdf_name = os.path.splitext(uploaded_file.name)[0]  # Extract PDF name without extension
    with st.spinner("Processing PDF..."):
        try:
            # Read PDF bytes
            pdf_bytes = uploaded_file.read()
            uploaded_file.seek(0)  # Reset pointer for reuse

            # Try PDF plumbler first
            try:
                pdf = open_pdf(uploaded_file, password=pdf_password if pdf_password else None)
                rows = extract_tables(pdf)
                bank_type = detect_bank_type(rows, pdf)
            except:
                pdf = None
                rows = []
                bank_type = "SBI Bank (OCR)"

            df = pd.DataFrame()

            # -----------------------------
            # BOI BANK
            # -----------------------------
            if bank_type == "BOI Bank":
                boi_rows = []
                for row in rows:
                    cleaned = [cell.strip() if cell else "" for cell in row]
                    if "Sr No" in cleaned: continue
                    if len(cleaned) >= 6:
                        boi_rows.append(cleaned[:6])
                df = pd.DataFrame(boi_rows, columns=["Sr No", "Date", "Remarks", "Debit", "Credit", "Balance"])
                for col in ["Debit", "Credit", "Balance"]:
                    df[col] = df[col].str.replace("₹", "", regex=False).str.replace(",", "", regex=False)

            # -----------------------------
            # HDFC BANK
            # -----------------------------
            elif bank_type == "HDFC Bank":
                date_re = re.compile(r"^(\d{2}/\d{2}/\d{4})")
                amount_re = re.compile(r"([\d,]+\.\d{2})")
                parsed_rows = []
                current_txn = None
                for page in pdf.pages:
                    text = page.extract_text()
                    if not text: continue
                    for line in text.split("\n"):
                        line = line.strip()
                        if not line: continue
                        date_match = date_re.match(line)
                        amounts = amount_re.findall(line)
                        if date_match and len(amounts) >= 3:
                            if current_txn: parsed_rows.append(current_txn)
                            narration_part = line[10:].strip()
                            current_txn = {
                                "Txn Date": date_match.group(1),
                                "Narration": narration_part,
                                "Withdrawals": amounts[-3].replace(",", ""),
                                "Deposits": amounts[-2].replace(",", ""),
                                "Closing Balance": amounts[-1].replace(",", "")
                            }
                        else:
                            if current_txn:
                                current_txn["Narration"] += " " + line
                if current_txn: parsed_rows.append(current_txn)
                df = pd.DataFrame(parsed_rows)

            # -----------------------------
            # KOTAK BANK
            # -----------------------------
            elif bank_type == "Kotak Bank":
                date_re = re.compile(r"\d{2}\s[A-Za-z]{3}\s\d{4}")
                amount_re = re.compile(r"[+-]?\d+(?:,\d{2,3})*\.\d{2}")
                ref_re = re.compile(r"\b(?:MB|UPI|CHQ)[A-Z0-9\-]+\b")
                time_re = re.compile(r"\d{1,2}:\d{2}\s?(?:AM|PM)")
                junk_re = re.compile(
                    r"(Statement generated on|Page\s+\d+\s+of\s+\d+|"
                    r"Account Statement|TRANSACTION DATE|VALUE DATE|"
                    r"TRANSACTION DETAILS|CHQ / REF NO|DEBIT/CREDIT|BALANCE|"
                    r"Kotak)", re.IGNORECASE
                )
                parsed_rows = []
                current_txn = None
                row_no = 0
                for page in pdf.pages:
                    text = page.extract_text()
                    if not text: continue
                    for line in text.split("\n"):
                        line = line.strip()
                        if not line or junk_re.search(line): continue
                        dates = date_re.findall(line)
                        amounts = amount_re.findall(line)
                        if len(dates) >= 2 and len(amounts) >= 2:
                            if current_txn: parsed_rows.append(current_txn)
                            row_no += 1
                            txn_date, value_date = dates[0], dates[1]
                            cleaned = line.replace(txn_date, "").replace(value_date, "")
                            cleaned = time_re.sub("", cleaned)
                            ref_match = ref_re.search(cleaned)
                            ref_no = ref_match.group() if ref_match else ""
                            amt_match = re.search(amount_re, cleaned)
                            narration = cleaned[:amt_match.start()] if amt_match else cleaned
                            narration = narration.replace(ref_no, "").strip()
                            current_txn = {
                                "#": row_no,
                                "TRANSACTION DATE": txn_date,
                                "VALUE DATE": value_date,
                                "TRANSACTION DETAILS": narration,
                                "CHQ / REF NO.": ref_no,
                                "DEBIT/CREDIT (₹)": amounts[-2].replace(",", ""),
                                "BALANCE (₹)": amounts[-1].replace(",", "")
                            }
                        else:
                            if current_txn:
                                extra = time_re.sub("", line).strip()
                                if extra and not junk_re.search(extra):
                                    current_txn["TRANSACTION DETAILS"] += " " + extra
                if current_txn: parsed_rows.append(current_txn)
                df = pd.DataFrame(parsed_rows)

            # -----------------------------
            # ICICI BANK
            # -----------------------------
            elif bank_type == "ICICI Bank":
                all_lines = []
                for page in pdf.pages:
                    text = page.extract_text()
                    if text: all_lines.extend(text.split("\n"))
                clean_lines = [l.strip() for l in all_lines if l.strip() and "DATE MODE PARTICULARS" not in l and not l.startswith("Total:")]
                transactions = []
                i = 0
                while i < len(clean_lines):
                    line = clean_lines[i]
                    date_match = re.match(r"\d{2}-\d{2}-\d{4}", line)
                    if date_match:
                        date = date_match.group()
                        remaining = line[len(date):].strip()
                        if remaining.startswith("B/F"):
                            numbers = re.findall(r"\d[\d,]*\.\d{2}", remaining)
                            balance = numbers[0] if numbers else ""
                            transactions.append({"DATE": date, "MODE": "B/F", "PARTICULARS": "", "DEPOSITS": "", "WITHDRAWALS": "", "BALANCE": balance})
                            i += 1
                            continue
                        numbers = re.findall(r"\d[\d,]*\.\d{2}", line)
                        if len(numbers) == 2:
                            amount, balance = numbers
                            particulars = remaining.replace(amount, "").replace(balance, "").strip()
                            transactions.append({"DATE": date, "MODE": "", "PARTICULARS": particulars, "DEPOSITS": "", "WITHDRAWALS": "", "BALANCE": balance})
                            i += 1
                            continue
                        buffer = []
                        i += 1
                        while i < len(clean_lines):
                            next_line = clean_lines[i]
                            nums = re.findall(r"\d[\d,]*\.\d{2}", next_line)
                            if len(nums) == 2:
                                amount, balance = nums
                                cleaned = next_line.replace(amount, "").replace(balance, "").strip()
                                buffer.append(cleaned)
                                particulars = " ".join(buffer).strip()
                                transactions.append({"DATE": date, "MODE": "", "PARTICULARS": particulars, "DEPOSITS": "", "WITHDRAWALS": "", "BALANCE": balance})
                                break
                            else:
                                buffer.append(next_line)
                            i += 1
                    i += 1
                df = pd.DataFrame(transactions)
                for j in range(1, len(df)):
                    try:
                        prev_balance = float(df.loc[j-1, "BALANCE"].replace(",", ""))
                        curr_balance = float(df.loc[j, "BALANCE"].replace(",", ""))
                        amount = abs(curr_balance - prev_balance)
                        amount_str = f"{amount:,.2f}"
                        if curr_balance > prev_balance:
                            df.loc[j, "DEPOSITS"] = amount_str
                        else:
                            df.loc[j, "WITHDRAWALS"] = amount_str
                    except:
                        pass
                df = df[["DATE", "MODE", "PARTICULARS", "DEPOSITS", "WITHDRAWALS", "BALANCE"]]

            # -----------------------------
            # GENERIC TABLE BANKS (Canara/Union/Axis)
            # -----------------------------
            elif bank_type in ["Canara Bank", "Union Bank", "Axis Bank"]:
                header = rows[0]
                data_rows = rows[1:]
                normalized_rows = []
                num_cols = len(header)
                for row in data_rows:
                    if len(row) > num_cols: row = row[:num_cols]
                    elif len(row) < num_cols: row += [""] * (num_cols - len(row))
                    normalized_rows.append(row)
                df = pd.DataFrame(normalized_rows, columns=header)

            # -----------------------------
            # SBI OCR BANK (FIXED - NO DATE DUPLICATION)
            # -----------------------------
            elif bank_type == "SBI Bank (OCR)":
                df = extract_sbi_transactions(pdf_bytes)

            # -----------------------------
            # EXPORT
            # -----------------------------
            if not df.empty:
                output = io.BytesIO()
                df.to_excel(output, index=False, engine="openpyxl")
                output.seek(0)
                
                # Create clean filename
                base_name = pdf_name.replace(" ", "_")
                excel_filename = f"{base_name}_Statement.xlsx"
                
                st.success(f"✅ {bank_type} statement converted successfully!")
                st.download_button(
                    label="📥 Download Excel",
                    data=output,
                    file_name=excel_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Show preview
                with st.expander("Preview First 10 Rows"):
                    st.dataframe(df.head(10))
            else:
                st.warning("⚠ No transactions detected.")

        except Exception as e:
            st.error(f"❌ Error processing PDF: {e}")

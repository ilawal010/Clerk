
# ---------------------------------------
# NITT Memo & Mail Tracking System
# Fully Consolidated Streamlit App
# Author: Engr. Lawal Ibrahim Umar
# ---------------------------------------

import streamlit as st
import pandas as pd
import os
from datetime import datetime
from PIL import Image
import uuid
import fitz  # PyMuPDF for PDF preview

# --------------------------------------------
# Setup folders for scanned memos and records
# --------------------------------------------
os.makedirs("memos/scanned", exist_ok=True)
os.makedirs("memos/records", exist_ok=True)
os.makedirs("memos/attachments", exist_ok=True)

memo_file = "memos/records/memo_records.xlsx"
scanned_folder = "memos/scanned/"

# --------------------------------------------
# Global Departments
# --------------------------------------------
departments = [
    "DG/CE", "SA","TA","Transport School", "Training Department", 
    "Transport Research & Intelligence Department","Transport Technology Centre (TTC)",
    "Consultancy Services Department", "Library & Information Services Department",
    "Registry", "Bursary", "Internal Audit",
    "PPP, Partnership & Collaboration Unit", "Legal Services Unit", 
    "Press & Public Relations Unit","Procurement Unit", "Physical Planning Unit", 
    "SERVICOM", "ACTU – Anti-Corruption & Transparency Unit","Medical Center"
]

# --------------------------------------------
# Helper Functions
# --------------------------------------------
def generate_memo_number():
    """Generate a unique memo number: NITT/DG/Year/0001"""
    year = datetime.now().year
    unique_id = str(uuid.uuid4().int)[:4]  # 4-digit unique number
    return f"NITT/DG/{year}/{unique_id}"

def save_memo_record(data):
    """Save memo record to Excel; create file if it doesn't exist"""
    if os.path.exists(memo_file):
        df = pd.read_excel(memo_file)
        df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
    else:
        df = pd.DataFrame([data])
    # Ensure History column exists
    if 'History' not in df.columns:
        df['History'] = [[] for _ in range(len(df))]
    else:
        df['History'] = df['History'].apply(lambda x: [] if pd.isna(x) else x)
    df.to_excel(memo_file, index=False)

def load_memos():
    """Load existing memo records safely"""
    if os.path.exists(memo_file):
        df = pd.read_excel(memo_file)
        if 'History' not in df.columns:
            df['History'] = [[] for _ in range(len(df))]
        else:
            df['History'] = df['History'].apply(lambda x: [] if pd.isna(x) else x)
        return df
    else:
        return pd.DataFrame()

def stamp_pdf_with_memo_number(pdf_path, memo_number):
    """Adds memo number as a watermark on first page and saves new file"""
    doc = fitz.open(pdf_path)
    page = doc[0]
    text_position = fitz.Point(50, 50)
    page.insert_text(
        text_position,
        f"Reference No: {memo_number}",
        fontsize=14,
        fontname="helv",
        rotate=0,
        color=(1, 0, 0)
    )
    stamped_path = pdf_path.replace(".pdf", "_stamped.pdf")
    doc.save(stamped_path)
    doc.close()
    return stamped_path

# --------------------------------------------
# Load memo records
# --------------------------------------------
df_memos = load_memos()

# --------------------------------------------
# Sidebar Navigation
# --------------------------------------------
st.sidebar.title("NITT Memo Management System")
page = st.sidebar.radio("Navigate", ["Log Memo", "Dashboard & Search", "Preview & Approve / Forward"])

# --------------------------------------------
# Page 1: Log Memo
# --------------------------------------------
if page == "Log Memo":
    st.title("📝 Log Internal/External Memo")
    
    memo_type = st.selectbox("Select Memo Type:", ["Internal", "External"])
    title = st.text_input("Memo Title")
    date_received = st.date_input("Date Received", datetime.today())
    signatory = st.text_input("Signatory/Author of Memo")
    
    if memo_type == "Internal":
        from_department = st.selectbox("From Department/Unit:", departments)
        to_department = st.selectbox("To Department/Unit:", departments)
    else:
        sender_name = st.text_input("Sender Name / Organization")
        sender_address = st.text_area("Sender Address")
        sender_email = st.text_input("Sender Email")
        sender_phone = st.text_input("Sender Phone Number")
    
    uploaded_file = st.file_uploader("Upload Scanned Memo (PDF/Image)", type=["pdf", "png", "jpg", "jpeg"])
    memo_number = generate_memo_number()
    st.info(f"Generated Memo Number: {memo_number}")
    
    if st.button("Save Memo Record"):
        if title.strip() == "":
            st.error("Please enter memo title.")
        elif uploaded_file is None:
            st.error("Please upload scanned memo.")
        else:
            safe_memo_number = memo_number.replace("/", "-")
            file_extension = uploaded_file.name.split(".")[-1]
            scanned_file_path = f"{scanned_folder}/{safe_memo_number}.{file_extension}"
            
            with open(scanned_file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            if file_extension.lower() == "pdf":
                scanned_file_path = stamp_pdf_with_memo_number(scanned_file_path, memo_number)
            
            memo_data = {
                "Memo Number": memo_number,
                "Title": title,
                "Date Received": date_received,
                "Type": memo_type,
                "Signatory": signatory,
                "Status": "Pending",
                "Current Location": to_department if memo_type=="Internal" else sender_name
            }
            if memo_type == "Internal":
                memo_data.update({"From": from_department, "To": to_department})
            else:
                memo_data.update({
                    "Sender Name": sender_name,
                    "Sender Address": sender_address,
                    "Sender Email": sender_email,
                    "Sender Phone": sender_phone
                })
            
            save_memo_record(memo_data)
            st.success("Memo logged successfully!")

# --------------------------------------------
# Page 2: Dashboard & Search (Updated)
# --------------------------------------------
elif page == "Dashboard & Search":
    st.title("📊 Memo Dashboard & Search")

    # Reload memos to ensure latest data
    df_memos = load_memos()

    if df_memos.empty:
        st.info("No memo records to display.")
    else:
        # Ensure 'Current Location' exists
        if 'Current Location' not in df_memos.columns:
            df_memos['Current Location'] = df_memos.get('To', '')

        # Memo Type Filter
        memo_type_filter = st.selectbox(
            "Filter by Memo Type",
            ["All"] + df_memos['Type'].dropna().unique().tolist()
        )

        # Department / Sender Filter
        dept_filter = "All"
        if memo_type_filter == "Internal" and 'From' in df_memos.columns:
            depts = df_memos['From'].dropna().unique().tolist()
            dept_filter = st.selectbox("Filter by Department", ["All"] + depts)
        elif memo_type_filter == "External" and 'Sender Name' in df_memos.columns:
            senders = df_memos['Sender Name'].dropna().unique().tolist()
            dept_filter = st.selectbox("Filter by Sender Name", ["All"] + senders)

        # Date Range Filter
        min_date = df_memos['Date Received'].min() if 'Date Received' in df_memos.columns else datetime.today()
        max_date = df_memos['Date Received'].max() if 'Date Received' in df_memos.columns else datetime.today()
        date_range = st.date_input("Filter by Date Range", [min_date, max_date])

        # Apply Filters
        df_filtered = df_memos.copy()
        if memo_type_filter != "All":
            df_filtered = df_filtered[df_filtered['Type'] == memo_type_filter]
        if dept_filter != "All":
            if memo_type_filter == "Internal":
                df_filtered = df_filtered[df_filtered['From'] == dept_filter]
            else:
                df_filtered = df_filtered[df_filtered['Sender Name'] == dept_filter]
        if len(date_range) == 2:
            start_date, end_date = date_range
            df_filtered = df_filtered[
                (pd.to_datetime(df_filtered['Date Received']) >= pd.to_datetime(start_date)) &
                (pd.to_datetime(df_filtered['Date Received']) <= pd.to_datetime(end_date))
            ]

        # Display Filtered Records
        display_cols = ["Memo Number","Title","Type","From","To","Date Received","Status","Current Location"]
        df_filtered['Date Received'] = pd.to_datetime(df_filtered['Date Received']).dt.strftime('%Y-%m-%d')
        st.subheader("Filtered Memo Records")
        st.dataframe(df_filtered[display_cols])

        # Select Memo & Check Status/Current Location
        st.subheader("Check Memo Status & Current Location")
        memo_list = df_filtered['Memo Number'].tolist()
        selected_memo = st.selectbox("Select Memo Number", options=[""] + memo_list)
        search_memo = st.text_input("Or enter Memo Number")
        memo_to_check = search_memo.strip() if search_memo.strip() else selected_memo

        if memo_to_check:
            memo_info = df_filtered[df_filtered['Memo Number'] == memo_to_check]
            if not memo_info.empty:
                st.write(f"**Memo Number:** {memo_info.iloc[0]['Memo Number']}")
                st.write(f"**Status:** {memo_info.iloc[0]['Status']}")
                st.write(f"**Current Location:** {memo_info.iloc[0]['Current Location']}")
            else:
                st.warning("Memo not found in the filtered records.")

        # Download Filtered Records
        if st.button("Download Filtered Records"):
            df_filtered.to_csv("filtered_memos.csv", index=False)
            st.success("Filtered records saved as 'filtered_memos.csv'.")

# --------------------------------------------
# Page 3: Preview, Forward & Approve
# --------------------------------------------
elif page == "Preview & Approve / Forward":
    st.title("🖼️ Preview Memo, Forward & Approve")
    
    if df_memos.empty:
        st.info("No memos to preview.")
    else:
        memo_numbers = df_memos['Memo Number'].tolist()
        selected_memo = st.selectbox("Select Memo Number", options=memo_numbers)
        memo_index = df_memos[df_memos['Memo Number']==selected_memo].index[0]
        memo_data = df_memos.loc[memo_index]
        
        st.subheader("📄 Memo Summary")
        st.markdown(f"**Title:** {memo_data.get('Title','')}")
        st.markdown(f"**From:** {memo_data.get('From', memo_data.get('Sender Name',''))}")
        st.markdown(f"**To:** {memo_data.get('To','')}")
        st.markdown(f"**Date Received:** {memo_data.get('Date Received','')}")
        st.markdown(f"**Status:** {memo_data.get('Status','Pending')}")
        st.markdown(f"**Current Location:** {memo_data.get('Current Location','')}")
        
        # --- Display scanned memo ---
        scanned_file_path = os.path.join(scanned_folder, f"{selected_memo}.pdf")
        if not os.path.exists(scanned_file_path):
            for ext in ["png","jpg","jpeg"]:
                temp_path = os.path.join(scanned_folder, f"{selected_memo}.{ext}")
                if os.path.exists(temp_path):
                    scanned_file_path = temp_path
                    break
        if os.path.exists(scanned_file_path):
            if scanned_file_path.lower().endswith(".pdf"):
                st.subheader("PDF Preview")
                pdf_doc = fitz.open(scanned_file_path)
                for i in range(len(pdf_doc)):
                    page = pdf_doc.load_page(i)
                    pix = page.get_pixmap()
                    img = Image.frombytes("RGB",[pix.width,pix.height],pix.samples)
                    st.image(img, caption=f"Page {i+1}", use_column_width=True)
            else:
                st.subheader("Image Preview")
                img = Image.open(scanned_file_path)
                st.image(img, use_column_width=True)
        else:
            st.warning("Scanned memo file not found.")
        
        # Forward / Reply
        st.subheader("↪ Forward / Reply Memo")
        forward_to = st.selectbox("Forward/Reply To:", ["Select Department/User"] + departments)
        comment = st.text_area("Add Comment (optional)")
        uploaded_attachment = st.file_uploader("Attach additional document (optional)", type=["pdf","png","jpg","jpeg"])
        
        if st.button("Forward/Reply"):
            history = df_memos.at[memo_index,"History"]
            if not isinstance(history,list):
                history=[]
            entry = {
                "Date": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "Action":"Forwarded/Reply",
                "To": forward_to,
                "Comment": comment,
                "Attachment": uploaded_attachment.name if uploaded_attachment else None
            }
            history.append(entry)
            df_memos.at[memo_index,"History"]=history
            df_memos.at[memo_index,"Current Location"]=forward_to
            df_memos.at[memo_index,"Status"]=f"Forwarded to {forward_to}"
            
            # Save attachment
            if uploaded_attachment:
                attach_path = os.path.join("memos/attachments", f"{selected_memo}_{uploaded_attachment.name}")
                with open(attach_path,"wb") as f:
                    f.write(uploaded_attachment.getbuffer())
            
            df_memos.to_excel(memo_file,index=False)
            st.success(f"Memo forwarded/replied to {forward_to} successfully!")
        
        st.subheader("🕒 Memo History")
        memo_history = df_memos.at[memo_index,"History"]
        if memo_history and isinstance(memo_history,list):
            for h in memo_history:
                st.markdown(f"- **Date:** {h['Date']}, **Action:** {h['Action']}, **To:** {h['To']}, **Comment:** {h['Comment']}, **Attachment:** {h.get('Attachment','None')}")
        else:
            st.info("No history available for this memo yet.")



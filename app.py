import streamlit as st
import pandas as pd
import pdfplumber
import io

st.set_page_config(page_title="PDF to Excel - Style Breakdown", layout="centered")

st.title("📄 PDF to Excel Converter (Garments WO)")
st.write("আপনার ওয়ার্ক অর্ডার PDF ফাইলটি আপলোড করুন এবং সরাসরি **STYLE, COLOUR, SIZE, Quantity** ফরম্যাটে Excel ডাউনলোড করুন।")

uploaded_file = st.file_uploader("PDF ফাইল সিলেক্ট করুন", type=["pdf"])

if uploaded_file is not None:
    st.success("ফাইল সফলভাবে আপলোড হয়েছে!")
    
    with st.spinner("ডেটা গভীরভাবে প্রসেস করা হচ্ছে... অনুগ্রহ করে অপেক্ষা করুন।"):
        extracted_rows = []
        
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if not row or len(row) < 4:
                            continue
                        
                        # ১. STYLE এক্সট্রাক্ট ও ক্লিন করা (১ম কলাম)
                        col0 = str(row[0]).strip() if row[0] else ""
                        if any(x in col0.upper() for x in ["STYLE", "TOTAL", "GRAND", "INVOICE", "PRODUCT"]):
                            continue
                        
                        # প্রথম লাইন বা স্পেস রিমুভ করে ক্লিন স্টাইল নেওয়া
                        style = col0.split("\n")[0].strip().replace(" ", "")
                        if not style or style.startswith("OFF") or style.startswith("THR"):
                            continue
                        
                        # ২. QUANTITY এক্সট্রাক্ট ও ক্লিন করা (শেষ কলাম)
                        qty_raw = str(row[-1]).strip() if row[-1] else ""
                        if not qty_raw or "PCS" in qty_raw.upper() and len(qty_raw.split("\n")) == 1:
                            continue
                        
                        # যদি কোয়ান্টিটি সেলের ভেতর টোটাল বা এক্সট্রা নম্বর ঢুকে থাকে, শুধু ১ম লাইনের মেইন কোয়ান্টিটি নেওয়া
                        qty_clean = qty_raw.split("\n")[0].replace(",", "").strip()
                        try:
                            qty = float(qty_clean)
                        except ValueError:
                            continue
                        
                        # ৩. COLOUR এবং SIZE নির্ধারণ
                        colour = "N/A"
                        size = "N/A"
                        known_sizes = ["XS", "S", "M", "L", "XL", "XXL", "3XL"]
                        
                        # পুরো রোর সব টেক্সট কম্বাইন করে সার্চ করা (যাতে কলাম এদিক ওদিক হলেও ডেটা মিস না হয়)
                        full_row_text = " ".join([str(cell) for cell in row if cell])
                        lines = [line.strip() for line in full_row_text.split("\n") if line.strip()]
                        
                        # সাইজ খোঁজা (যেমন: XS, S, M অথবা Thermal জবের SACV কোড)
                        for line in lines:
                            clean_line = line.replace(",", "").strip()
                            if clean_line in known_sizes:
                                size = clean_line
                            elif "SACV" in clean_line.upper():
                                # Thermal জবের ক্ষেত্রে SACV কোডটিই সাইজ হিসেবে গণ্য হবে
                                size = clean_line.split()[-1] 
                        
                        # কালার খোঁজা (Euro & Blang স্টিকারের জন্য)
                        if "NERO" in full_row_text.upper():
                            colour = "NERO"
                        elif "ROSA CHIARO" in full_row_text.upper() or "ROSA" in full_row_text.upper():
                            colour = "VAR ROSA CHIARO"
                        elif "BIANCO" in full_row_text.upper():
                            colour = "VAR BIANCO OTTICO"
                        elif "NUDU" in full_row_text.upper():
                            colour = "VAR NUDU"
                        else:
                            colour = "N/A" # Thermal জবের জন্য N/A থাকবে
                        
                        # লিস্টে যোগ করা
                        if style:
                            extracted_rows.append({
                                "STYLE": style,
                                "COLOUR": colour,
                                "SIZE": size,
                                "Quantity": qty
                            })
                            
        if extracted_rows:
            df_result = pd.DataFrame(extracted_rows)
            
            # একই STYLE, COLOUR, SIZE এর ডুপ্লিকেট রো থাকলে কোয়ান্টিটি যোগ (Sum) করে দেওয়া হবে
            df_result = df_result.groupby(["STYLE", "COLOUR", "SIZE"], as_index=False)["Quantity"].sum()
            
            st.subheader("📊 আপনার এক্সেল ডেটার প্রিভিউ:")
            st.dataframe(df_result)
            
            # Excel ফাইল তৈরি
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_result.to_excel(writer, sheet_name="Style_Breakdown", index=False)
            output.seek(0)
            
            st.download_button(
                label="📥 Excel ফাইল ডাউনলোড করুন",
                data=output,
                file_name="Style_Breakdown_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("❌ দুঃখিত! এই পিডিএফ থেকে কোনো ডেটা এক্সট্রাক্ট করা যায়নি। কোডের ম্যাচিং লজিকটি আরেকবার রিভিউ করা দরকার।")

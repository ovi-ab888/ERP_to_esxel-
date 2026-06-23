import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

st.set_page_config(page_title="PDF to Excel - Style Breakdown", layout="centered")

st.title("📄 PDF to Excel Converter (Garments WO)")
st.write("আপনার ওয়ার্ক অর্ডার PDF ফাইলটি আপলোড করুন এবং সরাসরি **STYLE, COLOUR, SIZE, Quantity** ফরম্যাটে Excel ডাউনলোড করুন।")

uploaded_file = st.file_uploader("PDF ফাইল সিলেক্ট করুন", type=["pdf"])

if uploaded_file is not None:
    st.success("ফাইল সফলভাবে আপলোড হয়েছে!")
    
    with st.spinner("ডেটা গভীরভাবে প্রসেস করা হচ্ছে... অনুগ্রহ করে অপেক্ষা করুন।"):
        extracted_rows = []
        
        # সাইজ চেনার স্ট্যান্ডার্ড লিস্ট
        known_sizes = ["XS", "S", "M", "L", "XL", "XXL", "3XL"]
        
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                # টেবিল এক্সট্রাক্ট না করে সরাসরি পেজের পুরো টেক্সট নেওয়া হচ্ছে
                text = page.extract_text()
                if not text:
                    continue
                
                # লাইন বাই লাইন ভাগ করা
                lines = text.split("\n")
                
                for line in lines:
                    line_upper = line.upper()
                    
                    # হেডার, টোটাল বা অপ্রয়োজনীয় লাইন স্কিপ করা
                    if any(x in line_upper for x in ["TOTAL", "GRAND", "PRICE", "INVOICE", "PRODUCT", "DELIVERY", "PACKAGE", "KeyEntry"]):
                        continue
                    
                    # যদি লাইনে স্টাইল নম্বর (SLMD) থাকে, তবেই আমরা ডেটা প্রসেস করব
                    if "SLMD" in line_upper:
                        
                        # ১. STYLE এক্সট্রাক্ট করা
                        style_match = re.search(r'(SLMD\d+P\d+)', line_upper)
                        style = style_match.group(1) if style_match else "SLMD50197P27"
                        
                        # ২. QUANTITY এক্সট্রাক্ট করা (লাইনের একদম শেষ অংশ যা একটি সংখ্যা)
                        # কমা এবং দশমিকসহ সংখ্যা খোঁজার রেজেক্স
                        qty_match = re.findall(r'(\d{1,3}(?:,\d{3})*(?:\.\d+))', line)
                        if not qty_match:
                            continue
                        
                        # সাধারণত লাইনের একদম শেষের সংখ্যাটিই কোয়ান্টিটি হয়
                        qty_clean = qty_match[-1].replace(",", "")
                        try:
                            qty = float(qty_clean)
                        except ValueError:
                            continue
                        
                        # ৩. SIZE এক্সট্রাক্ট করা
                        size = "N/A"
                        # Thermal জবের SACV কোড চেক করা
                        sacv_match = re.search(r'(SACV\d+)', line_upper)
                        if sacv_match:
                            size = sacv_match.group(1)
                        else:
                            # নরমাল সাইজ (XS, S, M...) চেক করা
                            for s in known_sizes:
                                # লাইনে যেন সাইজটি আলাদা শব্দ হিসেবে থাকে
                                if re.search(r'\b' + re.escape(s) + r'\b', line_upper):
                                    size = s
                                    break
                        
                        # ৪. COLOUR এক্সট্রাক্ট করা
                        colour = "N/A"
                        if "NERO" in line_upper:
                            colour = "NERO"
                        elif "ROSA" in line_upper:
                            colour = "VAR ROSA CHIARO"
                        elif "BIANCO" in line_upper:
                            colour = "VAR BIANCO OTTICO"
                        elif "NUDU" in line_upper:
                            colour = "VAR NUDU"
                        
                        # ডেটা অ্যাপেন্ড করা
                        extracted_rows.append({
                            "STYLE": style,
                            "COLOUR": colour,
                            "SIZE": size,
                            "Quantity": qty
                        })
                            
        if extracted_rows:
            df_result = pd.DataFrame(extracted_rows)
            
            # একই STYLE, COLOUR, SIZE থাকলে কোয়ান্টিটি যোগ (Sum) হবে
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
            st.error("❌ দুঃখিত! এই টেক্সট স্ক্যানিং মেথডেও কোনো ডেটা মেলানো যায়নি।")

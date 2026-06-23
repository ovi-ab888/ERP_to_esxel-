import streamlit as st
import pandas as pd
import pdfplumber
import io
import re

st.set_page_config(page_title="PDF to Excel - Style Breakdown", layout="centered")

st.title("📄 PDF to Excel Converter (Garments WO)")
st.write("আপনার ওয়ার্ক অর্ডার PDF ফাইলটি আপলোড করুন এবং সরাসরি **STYLE, COLOUR, SIZE, Quantity** ফরম্যাটে Excel ডাউনলোড করুন।")

uploaded_file = st.file_uploader("PDF ফাইল সিলেক্ট করুন", type=["pdf"])

if uploaded_file is not None:
    st.success("ফাইল সফলভাবে আপলোড হয়েছে!")
    
    with st.spinner("অফসেট এবং থার্মাল উভয় ডেটা গভীরভাবে স্ক্যান করা হচ্ছে... অনুগ্রহ করে অপেক্ষা করুন।"):
        extracted_rows = []
        
        known_sizes = ["XS", "S", "M", "L", "XL", "XXL", "3XL"]
        known_colors = {
            "NERO": "NERO",
            "ROSA": "VAR ROSA CHIARO",
            "BIANCO": "VAR BIANCO OTTICO",
            "NUDU": "VAR NUDU"
        }
        
        # ডিফল্ট স্টাইল (যদি শুরুতে কোনো স্টাইল না পায়)
        current_style = "SLMD50197P27"
        
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                
                lines = text.split("\n")
                for line in lines:
                    line_upper = line.upper().strip()
                    
                    # ১. অপ্রয়োজনীয় হেডার/টোটাল লাইন ফিল্টার
                    if any(x in line_upper for x in ["TOTAL", "GRAND", "PRICE", "INVOICE", "DELIVERY", "PRODUCT", "WIDTH", "LENGTH", "PACKAGE", "KEYENTRY"]):
                        continue
                    
                    # ২. স্টাইল ট্র্যাক করা (লাইনটিতে SLMD থাকলে কারেন্ট স্টাইল আপডেট হবে)
                    style_match = re.search(r'(SLMD\d+P\d+)', line_upper)
                    if style_match:
                        current_style = style_match.group(1)
                    
                    # ৩. থার্মাল সাবু আইটেম চেক (যদি লাইনে SACV এবং কোয়ান্টিটি দুটোই থাকে)
                    if "SACV" in line_upper:
                        sacv_match = re.search(r'(SACV\d+)', line_upper)
                        qty_match = re.search(r'(\d{1,4}\.\d{2})', line_upper)
                        
                        if sacv_match and qty_match:
                            size = sacv_match.group(1)
                            qty = float(qty_match.group(1).replace(",", ""))
                            if qty > 0 and qty != 8030.0:  # কোনো এক্সট্রা কোড ফিল্টার
                                extracted_rows.append({
                                    "STYLE": current_style,
                                    "COLOUR": "N/A",
                                    "SIZE": size,
                                    "Quantity": qty
                                })
                        continue  # থার্মাল প্রসেস শেষ হলে পরের লাইনে চলে যাবে
                    
                    # ৪. অফসেট আইটেম চেক (পৃষ্ঠা ১-৪ এর জন্য)
                    # লাইনের একদম শেষ শব্দটি যদি একটি দশমিক সংখ্যা হয় (যেমন: 1029.00)
                    qty_match = re.search(r'(\d{1,4}\.\d{2})$', line_upper)
                    if qty_match:
                        qty = float(qty_match.group(1).replace(",", ""))
                        
                        detected_size = "N/A"
                        detected_color = "N/A"
                        
                        # সাইজ নির্ধারণ
                        for sz in known_sizes:
                            if re.search(r'\b' + re.escape(sz) + r'\b', line_upper):
                                detected_size = sz
                                break
                        
                        # কালার নির্ধারণ
                        for c_key, c_val in known_colors.items():
                            if c_key in line_upper:
                                detected_color = c_val
                                break
                        
                        # সাইজ অথবা কালার যেকোনো একটি পাওয়া গেলেই সেটি ভ্যালিড অফসেট ডেটা
                        if detected_size != "N/A" or detected_color != "N/A":
                            extracted_rows.append({
                                "STYLE": current_style,
                                "COLOUR": detected_color,
                                "SIZE": detected_size,
                                "Quantity": qty
                            })
                            
        if extracted_rows:
            df_result = pd.DataFrame(extracted_rows)
            
            # একই STYLE, COLOUR, SIZE এর ডেটা একসাথে যোগ (Sum) করা
            df_result = df_result.groupby(["STYLE", "COLOUR", "SIZE"], as_index=False)["Quantity"].sum()
            
            st.subheader("📊 আপনার পূর্ণাঙ্গ এক্সেল ডেটার প্রিভিউ:")
            st.dataframe(df_result)
            
            # Excel ফাইল তৈরি করা
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_result.to_excel(writer, sheet_name="Style_Breakdown", index=False)
            output.seek(0)
            
            st.download_button(
                label="📥 সম্পূর্ণ Excel ফাইল ডাউনলোড করুন",
                data=output,
                file_name="Final_Style_Breakdown_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("❌ দুঃখিত! এই মডিউলেও কোনো ডেটা ম্যাচ করানো যায়নি।")

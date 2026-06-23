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
    
    with st.spinner("অফসেট এবং থার্মাল উভয় ডেটা সম্পূর্ণভাবে প্রসেস করা হচ্ছে... অনুগ্রহ করে অপেক্ষা করুন।"):
        extracted_rows = []
        
        # পিডিএফ-এর পুরো টেক্সট একসাথে জমা করার জন্য
        full_pdf_text = ""
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    full_pdf_text += page_text + "\n"
        
        # ১. থার্মাল সাবু আইটেমগুলোর জন্য রেজেক্স (SACV কোড যুক্ত লাইন)
        # উদাহরণ: "SLMD50197P27 PK/121-0137 PAK-137- SABU STICKER(4CMX4 CM) N/A SACV095903001 138.00"
        thermal_pattern = re.compile(
            r'(SLMD\d+P\d+).*?(SACV\d+)\s+([\d,]+\.\d+)'
        )
        
        for match in thermal_pattern.finditer(full_pdf_text):
            style = match.group(1)
            size = match.group(2)
            qty_str = match.group(3).replace(",", "")
            
            try:
                qty = float(qty_str)
                # টোটাল বা অপ্রাসঙ্গিক বড় সংখ্যা বাদ দেওয়া
                if qty > 0 and qty != 8030.0:
                    extracted_rows.append({
                        "STYLE": style,
                        "COLOUR": "N/A",
                        "SIZE": size,
                        "Quantity": qty
                    })
            except ValueError:
                continue

        # ২. অফসেট আইটেমগুলোর জন্য রেজেক্স (XS, S, M, L, XL যুক্ত লাইন)
        # উদাহরণ: "SLMD50197P27 PK/121-0137 PAK-137- EURO STICKER( 4.4 CM X 3.3 CM) XS NERO 1029.00"
        # এখানে কালার ও সাইজ আগে-পরে বা নিউলাইনেও থাকতে পারে, তাই ফ্লেক্সিবল ম্যাচিং ব্যবহার করা হয়েছে।
        known_sizes = ["XS", "S", "M", "L", "XL", "XXL", "3XL"]
        known_colors = {
            "NERO": "NERO",
            "ROSA": "VAR ROSA CHIARO",
            "BIANCO": "VAR BIANCO OTTICO",
            "NUDU": "VAR NUDU"
        }
        
        # লাইন বাই লাইন অফসেট ব্লক চেক করা (১ থেকে ৪ নম্বর পৃষ্ঠা)
        lines = full_pdf_text.split("\n")
        current_style = "SLMD50197P27" # ডিফল্ট ব্যাকআপ স্টাইল
        
        for line in lines:
            line_upper = line.upper()
            
            # হেডার বা সামারি লাইন স্কিপ করা
            if any(x in line_upper for x in ["TOTAL", "GRAND", "PRICE", "INVOICE", "DELIVERY", "PRODUCT", "WIDTH", "LENGTH", "PACKAGE", "PCS"]):
                continue
            
            # লাইনে যদি স্টাইল থাকে তবে আপডেট করা
            if "SLMD" in line_upper:
                style_match = re.search(r'(SLMD\d+P\d+)', line_upper)
                if style_match:
                    current_style = style_match.group(1)
            
            # থার্মাল লাইনগুলো ইতিমধ্যে কভারড, তাই এখানে স্কিপ
            if "SACV" in line_upper:
                continue
                
            # লাইনের শেষে যদি কোয়ান্টিটি থাকে (যেমন: 1029.00)
            qty_match = re.search(r'(\d{1,4}\.\d{2})$', line.strip())
            if qty_match:
                qty = float(qty_match.group(1))
                
                # সাইজ এবং কালার খোঁজা এই নির্দিষ্ট লাইনে
                detected_size = "N/A"
                detected_color = "N/A"
                
                # সাইজ ডিটেকশন
                for sz in known_sizes:
                    if re.search(r'\b' + re.escape(sz) + r'\b', line_upper):
                        detected_size = sz
                        break
                
                # কালার ডিটেকশন
                for c_key, c_val in known_colors.items():
                    if c_key in line_upper:
                        detected_color = c_val
                        break
                
                # যদি অন্তত সাইজ বা কালার কোনো একটি পাওয়া যায়, তবেই এটি ভ্যালিড অফসেট রো
                if detected_size != "N/A" or detected_color != "N/A":
                    extracted_rows.append({
                        "STYLE": current_style,
                        "COLOUR": detected_color,
                        "SIZE": detected_size,
                        "Quantity": qty
                    })

        if extracted_rows:
            df_result = pd.DataFrame(extracted_rows)
            
            # ডেটা ক্লিনআপ: একই গ্রুপের ডেটা যোগ (Sum) করা
            df_result = df_result.groupby(["STYLE", "COLOUR", "SIZE"], as_index=False)["Quantity"].sum()
            
            st.subheader("📊 আপনার পূর্ণাঙ্গ এক্সেল ডেটার প্রিভিউ:")
            st.dataframe(df_result)
            
            # Excel ফাইল তৈরি করা ইন-মেমোরি বাফারে
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_result.to_excel(writer, sheet_name="Style_Breakdown", index=False)
            output.seek(0)
            
            st.download_button(
                label="📥 চূড়ান্ত Excel ফাইল ডাউনলোড করুন",
                data=output,
                file_name="Complete_Style_Breakdown_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("❌ দুঃখিত! এই পিডিএফ থেকে কোনো ডেটা ম্যাচ করানো যায়নি।")

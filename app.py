import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

st.set_page_config(page_title="PDF to Excel - Style Breakdown", layout="centered")

st.title("📄 PDF to Excel Converter")
st.write("আপনার ওয়ার্ক অর্ডার PDF ফাইলটি আপলোড করুন এবং সরাসরি **STYLE, COLOUR, SIZE, Quantity** ফরম্যাটে Excel ডাউনলোড করুন।")

uploaded_file = st.file_uploader("PDF ফাইল সিলেক্ট করুন", type=["pdf"])

if uploaded_file is not None:
    st.success("ফাইল সফলভাবে আপলোড হয়েছে!")
    
    with st.spinner("ডেটা প্রসেস করা হচ্ছে... অনুগ্রহ করে অপেক্ষা করুন।"):
        extracted_rows = []
        
        # pdfplumber দিয়ে PDF ওপেন করা
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        # খালি রো বা হেডার রো বাদ দেওয়ার ফিল্টার
                        if not row or any(h in str(row[0]).upper() for h in ["STYLE", "OFF0", "THR0", "TOTAL", "GRAND"]):
                            continue
                        
                        # কলাম সংখ্যা চেক করা (মূল টেবিলগুলোতে সাধারণত ৫ থেকে ১০টি কলাম থাকে)
                        if len(row) >= 5:
                            style_raw = str(row[0]).strip()
                            item_col_size_raw = str(row[2]).strip() if row[2] else ""
                            qty_raw = str(row[-1]).strip() if row[-1] else ""
                            
                            # যদি রো-তে ক্লিনিংয়ের জন্য ফাঁকা ডেটা থাকে তা স্কিপ করা
                            if not style_raw or "PCS" in qty_raw or "/" in qty_raw:
                                continue
                            
                            # ১. STYLE ক্লিন করা (নিউলাইন বাদ দেওয়া)
                            style = style_raw.replace("\n", "").replace(" ", "")
                            
                            # ২. Quantity ক্লিন করা
                            qty = qty_raw.split("\n")[0].replace(",", "")
                            try:
                                qty = float(qty)
                            except ValueError:
                                continue
                            
                            # ৩. COLOUR এবং SIZE এক্সট্রাক্ট করা (নিউলাইন টেক্সট থেকে)
                            # টেক্সটের ভেতরের অতিরিক্ত স্পেস ও নিউলাইন হ্যান্ডেল করা
                            lines = [line.strip() for line in item_col_size_raw.split("\n") if line.strip()]
                            
                            colour = "N/A"
                            size = "N/A"
                            
                            # সাইজ চেনার জন্য একটি স্ট্যান্ডার্ড লিস্ট (XS, S, M, L, XL, বা SACV কোড)
                            known_sizes = ["XS", "S", "M", "L", "XL", "XXL"]
                            
                            # টেক্সট লাইনগুলো থেকে সাইজ ও কালার খোঁজা
                            for line in lines:
                                if line in known_sizes or line.startswith("SACV"):
                                    size = line
                                elif "STICKER" not in line and line != "VAR" and len(line) > 1:
                                    # সাধারণত STICKER বা সাইজ বাদে বাকি অংশই কালার (যেমন: NERO, VAR ROSA CHIARO)
                                    if colour == "N/A":
                                        colour = line
                                    else:
                                        colour += " " + line
                            
                            # যদি ৩ নম্বর কলামে কালার আলাদা থাকে (যেমন পৃষ্ঠা ২-এর টেবিল ২)
                            if len(row) > 3 and str(row[3]).strip() in ["NERO", "VAR ROSA", "VAR BIANCO", "VAR NUDU"]:
                                colour = str(row[3]).strip().replace("\n", " ")
                            if len(row) > 5 and str(row[5]).strip() in known_sizes:
                                size = str(row[5]).strip()
                            
                            # ফাইনাল ডেটা লিস্টে অ্যাপেন্ড করা
                            extracted_rows.append({
                                "STYLE": style,
                                "COLOUR": colour,
                                "SIZE": size,
                                "Quantity": qty
                            })
        
        # DataFrame তৈরি
        df_result = pd.DataFrame(extracted_rows)
        
        if not df_result.empty:
            # একই Style, Colour, Size এর ডুপ্লিকেট থাকলে কোয়ান্টিটি যোগ করে দেবে
            df_result = df_result.groupby(["STYLE", "COLOUR", "SIZE"], as_index=False)["Quantity"].sum()
            
            st.subheader("📊 এক্সেল ডেটা প্রিভিউ:")
            st.dataframe(df_result)
            
            # Excel ফাইল তৈরি
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_result.to_excel(writer, sheet_name="Order_Breakdown", index=False)
            output.seek(0)
            
            # ডাউনলোড বাটন
            st.download_button(
                label="📥 Excel ফাইল ডাউনলোড করুন",
                data=output,
                file_name="Style_Breakdown_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("ফাইল থেকে নির্দিষ্ট ফরম্যাটে কোনো ডেটা খুঁজে পাওয়া যায়নি। পিডিএফ-এর ফরম্যাটটি আরেকবার চেক করুন।")

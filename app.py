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
        
        # সাইজ খোঁজার স্ট্যান্ডার্ড লিস্ট
        known_sizes = ["XS", "S", "M", "L", "XL", "XXL", "3XL"]
        
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if not row or len(row) < 4:
                            continue
                        
                        # ১. STYLE চেক ও ক্লিন করা (১ম কলাম)
                        col0 = str(row[0]).strip() if row[0] else ""
                        if any(x in col0.upper() for x in ["STYLE", "TOTAL", "GRAND", "INVOICE", "PRODUCT", "OFF0", "THR0"]):
                            continue
                        
                        # ১ম লাইনের ক্লিন স্টাইল আইডি নেওয়া
                        style = col0.split("\n")[0].strip().replace(" ", "")
                        if not style or len(style) < 5: 
                            continue
                        
                        # ২. QUANTITY চেক ও ক্লিন করা (শেষ কলাম)
                        qty_raw = str(row[-1]).strip() if row[-1] else ""
                        if not qty_raw or "QUANTITY" in qty_raw.upper():
                            continue
                        
                        # সেলের প্রথম লাইন থেকে পিওর কোয়ান্টিটি নম্বর বের করা
                        qty_line = qty_raw.split("\n")[0].replace(",", "").strip()
                        try:
                            qty = float(qty_line)
                        except ValueError:
                            continue
                        
                        # ৩. রো-এর সব টেক্সট একসাথে করে কালার ও সাইজ এক্সট্রাক্ট করা
                        # পুরো রোর খালি সেল বাদ দিয়ে স্পেস দিয়ে জোড়া লাগানো
                        full_row_text = " ".join([str(cell) for cell in row if cell])
                        
                        colour = "N/A"
                        size = "N/A"
                        
                        # --- কালার ডিটেকশন (regex বা explicit চেক) ---
                        if "NERO" in full_row_text.upper():
                            colour = "NERO"
                        elif "ROSA" in full_row_text.upper():
                            colour = "VAR ROSA CHIARO"
                        elif "BIANCO" in full_row_text.upper():
                            colour = "VAR BIANCO OTTICO"
                        elif "NUDU" in full_row_text.upper():
                            colour = "VAR NUDU"
                        
                        # --- সাইজ ডিটেকশন ---
                        # টেক্সটকে ভেঙে সিঙ্গেল ওয়ার্ডে নিয়ে চেক করা
                        words = [w.strip() for w in re.split(r'[\s\n]+', full_row_text) if w.strip()]
                        
                        for word in words:
                            # অফসেট জবের স্ট্যান্ডার্ড সাইজ চেক
                            if word in known_sizes:
                                size = word
                                break
                            # থার্মাল জবের SACV কোড চেক (যেমন: SACV095903001)
                            if word.upper().startswith("SACV"):
                                size = word
                                break
                        
                        # ফাইনাল লিস্টে যোগ করা
                        extracted_rows.append({
                            "STYLE": style,
                            "COLOUR": colour,
                            "SIZE": size,
                            "Quantity": qty
                        })
                            
        if extracted_rows:
            df_result = pd.DataFrame(extracted_rows)
            
            # ডুপ্লিকেট রো থাকলে কোয়ান্টিটি যোগ (Sum) করে ইউনিক ডেটা রাখা
            df_result = df_result.groupby(["STYLE", "COLOUR", "SIZE"], as_index=False)["Quantity"].sum()
            
            st.subheader("📊 আপনার এক্সেল ডেটার প্রিভিউ:")
            st.dataframe(df_result)
            
            # Excel ফাইল তৈরি (Memory Buffer-এ)
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

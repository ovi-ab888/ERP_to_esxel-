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
                        
                        # ১. প্রথম কলাম (STYLE) চেক ও ক্লিন
                        col0 = str(row[0]).strip() if row[0] else ""
                        if any(x in col0.upper() for x in ["STYLE", "TOTAL", "GRAND", "INVOICE", "PRODUCT", "OFF0", "THR0"]):
                            continue
                        
                        # স্টাইল আইডেন্টিফাই করার জন্য বেসিক ফিল্টার (যেহেতু স্টাইল সাধারণত SLMD দিয়ে শুরু হচ্ছে)
                        if "SLMD" not in col0.upper():
                            continue
                            
                        # স্টাইল থেকে নিউলাইন ও স্পেস ক্লিন করা
                        style = col0.replace("\n", " ").replace("  ", " ").strip()
                        
                        # ২. শেষ কলাম (QUANTITY) ক্লিন
                        qty_raw = str(row[-1]).strip() if row[-1] else ""
                        if not qty_raw or "PCS" in qty_raw.upper() and len(qty_raw.split("\n")) == 1:
                            continue
                        
                        # কোয়ান্টিটি সেলের ১ম লাইন থেকে শুধু সংখ্যাটি নেওয়া
                        qty_clean = qty_raw.split("\n")[0].replace(",", "").strip()
                        try:
                            qty = float(qty_clean)
                        except ValueError:
                            continue
                        
                        # ৩. COLOUR এবং SIZE ডিটেকশন
                        colour = "N/A"
                        size = "N/A"
                        
                        # রোর সমস্ত সেল একসাথে করে সার্চ করা
                        full_row_text = " ".join([str(cell) for cell in row if cell]).upper()
                        
                        # সাইজ চেনার স্ট্যান্ডার্ড লিস্ট
                        known_sizes = ["XS", "S", "M", "L", "XL", "XXL", "3XL"]
                        
                        # সাইজ নির্ধারণ (Euro & Blang এর জন্য)
                        for s in known_sizes:
                            # টেক্সটের মধ্যে স্পেস বা নিউলাইন দিয়ে আলাদা করা সাইড চেক
                            if f" {s} " in f" {full_row_text.replace('\n', ' ')} ":
                                size = s
                                break
                        
                        # Thermal সাবু স্টিকারের জন্য SACV কোডই সাইজ
                        if "SACV" in full_row_text:
                            for cell in row:
                                if cell and "SACV" in str(cell).upper():
                                    size = str(cell).strip().replace("\n", "")
                                    break
                        
                        # কালার নির্ধারণ
                        if "NERO" in full_row_text:
                            colour = "NERO"
                        elif "ROSA" in full_row_text:
                            colour = "VAR ROSA CHIARO"
                        elif "BIANCO" in full_row_text:
                            colour = "VAR BIANCO OTTICO"
                        elif "NUDU" in full_row_text:
                            colour = "VAR NUDU"
                        
                        # ফাইনাল লিস্টে ডেটা যোগ করা
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
            st.error("❌ দুঃখিত! এই পিডিএফ থেকে কোনো ডেটা এক্সট্রাক্ট করা যায়নি। কোডের ম্যাচিং লজিকটি আরেকবার রিভিউ করা দরকার।")

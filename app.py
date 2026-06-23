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
                        if not row or len(row) < 3:
                            continue
                        
                        # সম্পূর্ণ রোর টেক্সট চেক করে হেডার/টোটাল বাদ দেওয়া
                        full_row_text = " ".join([str(cell) for cell in row if cell]).upper()
                        if any(x in full_row_text for x in ["TOTAL", "GRAND", "PRICE", "INVOICE", "DELIVERY", "PRODUCT", "WIDTH", "LENGTH"]):
                            continue
                        
                        # ১. STYLE নির্ধারণ (যে সেলে SLMD আছে সেটি খুঁজে বের করা)
                        style = None
                        for cell in row:
                            if cell and "SLMD" in str(cell).upper():
                                for line in str(cell).split('\n'):
                                    if "SLMD" in line.upper():
                                        style = line.strip().replace(" ", "")
                                        break
                                if style:
                                    break
                        
                        if not style:
                            continue
                        
                        # ২. QUANTITY নির্ধারণ (ডান দিক থেকে প্রথম যে সেলে ভ্যালিড নাম্বার বা নাম্বারের লিস্ট আছে)
                        qty_list = []
                        for cell in reversed(row):
                            if cell:
                                lines = [l.strip() for l in str(cell).split('\n') if l.strip()]
                                temp_qtys = []
                                for l in lines:
                                    clean_l = l.replace(',', '').strip()
                                    if "PCS" in clean_l.upper() or "/" in clean_l:
                                        continue
                                    try:
                                        temp_qtys.append(float(clean_l))
                                    except ValueError:
                                        break
                                if temp_qtys:
                                    qty_list = temp_qtys
                                    break
                        
                        if not qty_list:
                            continue
                        
                        # ৩. SIZE এবং COLOUR নির্ধারণ (লাইন বাই লাইন এক্সট্রাকশন)
                        sizes = []
                        colors = []
                        known_sizes = ["XS", "S", "M", "L", "XL", "XXL", "3XL"]
                        
                        for cell in row:
                            if not cell:
                                continue
                            lines = [l.strip() for l in str(cell).split('\n') if l.strip()]
                            for l in lines:
                                # সাইজ ম্যাচিং
                                if l.upper() in known_sizes:
                                    sizes.append(l.upper())
                                elif "SACV" in l.upper():
                                    sizes.append(l.strip())
                                
                                # কালার ম্যাচিং
                                if "NERO" in l.upper():
                                    colors.append("NERO")
                                elif "ROSA" in l.upper():
                                    colors.append("VAR ROSA CHIARO")
                                elif "BIANCO" in l.upper():
                                    colors.append("VAR BIANCO OTTICO")
                                elif "NUDU" in l.upper():
                                    colors.append("VAR NUDU")
                        
                        # ৪. মাল্টিপল কোয়ান্টিটি বা মার্জড রো স্প্লিট করে ডেটা সাজানো
                        for i, q in enumerate(qty_list):
                            s = sizes[i] if i < len(sizes) else (sizes[0] if sizes else "N/A")
                            c = colors[i] if i < len(colors) else (colors[0] if colors else "N/A")
                            
                            extracted_rows.append({
                                "STYLE": style,
                                "COLOUR": c,
                                "SIZE": s,
                                "Quantity": q
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
            st.error("❌ দুঃখিত! এই পিডিএফ থেকে নির্দিষ্ট ফরম্যাটে কোনো ডেটা খুঁজে পাওয়া যায়নি।")

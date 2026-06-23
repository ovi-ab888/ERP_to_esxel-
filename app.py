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
        known_sizes = ["XS", "S", "M", "L", "XL", "XXL", "3XL"]
        
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if not row or len(row) < 3:
                            continue
                        
                        # অপ্রয়োজনীয় হেডার ফিল্টার করা
                        full_row_text = " ".join([str(cell) for cell in row if cell]).upper()
                        if any(x in full_row_text for x in ["TOTAL", "GRAND", "PRICE", "INVOICE", "DELIVERY", "PRODUCT", "WIDTH", "LENGTH", "PACKAGE"]):
                            continue
                        
                        # ১. STYLE ডিটেক্ট করা
                        style = None
                        for cell in row:
                            if cell and "SLMD" in str(cell).upper():
                                lines = [l.strip() for l in str(cell).split("\n") if l.strip()]
                                for line in lines:
                                    if "SLMD" in line.upper():
                                        style = line.replace(" ", "")
                                        break
                                if style:
                                    break
                        
                        if not style:
                            continue
                        
                        # ২. QUANTITY লিস্ট বের করা (মাল্টিপল লাইন হ্যান্ডেল করতে)
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
                                        val = float(clean_l)
                                        if val > 0 and val != 231.0 and val != 35.511 and val != 80.3:
                                            temp_qtys.append(val)
                                    except ValueError:
                                        break
                                if temp_qtys:
                                    qty_list = temp_qtys
                                    break
                        
                        if not qty_list:
                            continue
                        
                        # ৩. রোর ভেতরের সব লেখা ভেঙে সাইজ ও কালার খোঁজা
                        sizes = []
                        colors = []
                        
                        for cell in row:
                            if not cell:
                                continue
                            lines = [l.strip() for l in str(cell).split('\n') if l.strip()]
                            for l in lines:
                                l_upper = l.upper()
                                
                                # স্ট্যান্ডার্ড সাইজ চেক (\b দিয়ে যেন শব্দের অংশ না হয়, যেমন 'M' আলাদা চেনা)
                                if l_upper in known_sizes:
                                    sizes.append(l_upper)
                                # থার্মাল কোড সাইজ চেক
                                elif "SACV" in l_upper:
                                    sacv_match = re.search(r'(SACV\d+)', l_upper)
                                    if sacv_match:
                                        sizes.append(sacv_match.group(1))
                                
                                # কালার চেক
                                if "NERO" in l_upper:
                                    colors.append("NERO")
                                elif "ROSA" in l_upper:
                                    colors.append("VAR ROSA CHIARO")
                                elif "BIANCO" in l_upper:
                                    colors.append("VAR BIANCO OTTICO")
                                elif "NUDU" in l_upper:
                                    colors.append("VAR NUDU")
                        
                        # ৪. এক্সট্রাক্ট করা ডেটা সুবিন্যস্তভাবে সাজানো
                        for i, q in enumerate(qty_list):
                            s = sizes[i] if i < len(sizes) else (sizes[0] if sizes else "N/A")
                            c = colors[i] if i < len(colors) else (colors[0] if colors else "N/A")
                            
                            # থার্মাল জবের সাইজ হলে কালার সব সময় N/A থাকবে
                            if "SACV" in str(s):
                                c = "N/A"
                                
                            extracted_rows.append({
                                "STYLE": style,
                                "COLOUR": c,
                                "SIZE": s,
                                "Quantity": q
                            })
                            
        if extracted_rows:
            df_result = pd.DataFrame(extracted_rows)
            
            # ডুপ্লিকেট এন্ট্রি সামারি (Sum) করা
            df_result = df_result.groupby(["STYLE", "COLOUR", "SIZE"], as_index=False)["Quantity"].sum()
            
            st.subheader("📊 আপনার পূর্ণাঙ্গ এক্সেল ডেটার প্রিভিউ:")
            st.dataframe(df_result)
            
            # Excel ফাইল তৈরি
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
            st.error("❌ দুঃখিত! নির্দিষ্ট ফরম্যাটে কোনো ডেটা খুঁজে পাওয়া যায়নি।")

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
            for page_num, page in enumerate(pdf.pages, start=1):
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if not row or len(row) < 3:
                            continue
                        
                        # ১. পুরো রোর টেক্সট কম্বাইন করে অপ্রয়োজনীয় হেডার ফিল্টার করা
                        full_row_text = " ".join([str(cell) for cell in row if cell]).upper()
                        if any(x in full_row_text for x in ["TOTAL", "GRAND", "PRICE", "INVOICE", "DELIVERY", "PRODUCT", "WIDTH", "LENGTH", "PACKAGE"]):
                            continue
                        
                        # ২. STYLE খুঁজে বের করা (রোর যেকোনো জায়গায় SLMD থাকলে)
                        style = None
                        for cell in row:
                            if cell and "SLMD" in str(cell).upper():
                                # প্রথম লাইনটি স্টাইল হিসেবে নেওয়া (যেমন: SLMD50197P27)
                                lines = [l.strip() for l in str(cell).split("\n") if l.strip()]
                                if lines:
                                    style = lines[0].replace(" ", "")
                                    break
                        
                        if not style:
                            continue
                        
                        # ৩. রোর ভেতরের সব এলিমেন্টকে ক্লিন করে লিস্ট আকারে বের করা
                        all_lines = []
                        for cell in row:
                            if cell:
                                cell_lines = [l.strip() for l in str(cell).split("\n") if l.strip()]
                                all_lines.extend(cell_lines)
                        
                        # ৪. QUANTITY বের করা (লিস্টের সব ভ্যালিড ফ্লোট/নাম্বার যা অন্য কোনো কোড নয়)
                        qtys = []
                        for token in all_lines:
                            clean_token = token.replace(",", "").strip()
                            if "PCS" in clean_token.upper() or "/" in clean_token or "SLMD" in clean_token.upper() or "SACV" in clean_token.upper():
                                continue
                            try:
                                val = float(clean_token)
                                # সাধারণত দশমিকের পর ০ বসা কোয়ান্টিটিগুলো নেওয়া (যেমন: ১৩৮.০০, ২৩০.০০)
                                if val > 0 and val != 231.0 and val != 35.511 and val != 80.3: 
                                    qtys.append(val)
                            except ValueError:
                                continue
                        
                        if not qtys:
                            continue
                        
                        # ৫. SIZE এবং COLOUR খোঁজা
                        known_sizes = ["XS", "S", "M", "L", "XL", "XXL", "3XL"]
                        sizes = []
                        colors = []
                        
                        # টোকেন ওয়াইজ সাইজ ও কালার ডিটেক্ট করা
                        for token in all_lines:
                            token_upper = token.upper()
                            
                            # সাইজ ডিটেকশন (Standard Size অথবা Thermal Job এর SACV কোড)
                            if token_upper in known_sizes:
                                sizes.append(token_upper)
                            elif "SACV" in token_upper:
                                sizes.append(token.strip())
                                
                            # কালার ডিটেকশন
                            if "NERO" in token_upper:
                                colors.append("NERO")
                            elif "ROSA" in token_upper:
                                colors.append("VAR ROSA CHIARO")
                            elif "BIANCO" in token_upper:
                                colors.append("VAR BIANCO OTTICO")
                            elif "NUDU" in token_upper:
                                colors.append("VAR NUDU")
                        
                        # ৬. জোড়া লাগানো বা মাল্টিপল রো থাকলে সমানুপাতিক হারে সাজানো (পৃষ্ঠা ৫-৭ এর জন্য ক্রুশিয়াল)
                        for idx, q in enumerate(qtys):
                            s = sizes[idx] if idx < len(sizes) else (sizes[0] if sizes else "N/A")
                            c = colors[idx] if idx < len(colors) else (colors[0] if colors else "N/A")
                            
                            # Thermal Sabu জবের জন্য কালার সব সময় N/A থাকবে পিডিএফ অনুযায়ী
                            if "SACV" in s:
                                c = "N/A"
                                
                            extracted_rows.append({
                                "STYLE": style,
                                "COLOUR": c,
                                "SIZE": s,
                                "Quantity": q
                            })
                            
        if extracted_rows:
            df_result = pd.DataFrame(extracted_rows)
            
            # ডেটা গ্রুপ করে ডুপ্লিকেট আইটেমের কোয়ান্টিটি যোগ (Sum) করা
            df_result = df_result.groupby(["STYLE", "COLOUR", "SIZE"], as_index=False)["Quantity"].sum()
            
            st.subheader("📊 আপনার এক্সেল ডেটার প্রিভিউ:")
            st.dataframe(df_result)
            
            # Excel ফাইল তৈরি (In-memory Buffer)
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

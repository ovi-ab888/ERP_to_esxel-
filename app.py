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
        
        # পিডিএফ রিড করা শুরু
        with pdfplumber.open(uploaded_file) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        # যদি রো খালি হয় বা হেডার/টোটাল টাইপ রো হয় তবে স্কিপ
                        if not row or len(row) < 4:
                            continue
                        
                        # প্রথম কলামের টেক্সট চেক
                        col0 = str(row[0]).strip() if row[0] else ""
                        if any(x in col0.upper() for x in ["STYLE", "OFF0", "THR0", "TOTAL", "GRAND", "INVOICE", "PRODUCT"]):
                            continue
                        
                        # STYLE এবং QUANTITY বেসিক ক্লিন
                        style = col0.replace("\n", "").replace(" ", "")
                        
                        # শেষ কলাম সাধারণত Quantity হয়
                        qty_raw = str(row[-1]).strip() if row[-1] else ""
                        if not qty_raw or "PCS" in qty_raw.upper() or "/" in qty_raw:
                            continue
                        
                        # কোয়ান্টিটি নম্বর ফরম্যাটে নেওয়া (প্রথম অংশ, যদি নিউলাইন থাকে)
                        qty_clean = qty_raw.split("\n")[0].replace(",", "")
                        try:
                            qty = float(qty_clean)
                        except ValueError:
                            continue
                        
                        # প্রাথমিক মান নির্ধারণ
                        colour = "N/A"
                        size = "N/A"
                        
                        # সাইজ চেনার স্ট্যান্ডার্ড লিস্ট
                        known_sizes = ["XS", "S", "M", "L", "XL", "XXL", "3XL"]
                        
                        # --- প্যাটার্ন ১: ৩ নম্বর কলামে (row[2]) সব ডেটা একসাথে থাকা (যেমন পৃষ্ঠা ১ এর Euro Sticker) ---
                        col2_raw = str(row[2]).strip() if row[2] else ""
                        lines = [line.strip() for line in col2_raw.split("\n") if line.strip()]
                        
                        for line in lines:
                            if line in known_sizes:
                                size = line
                            elif "STICKER" not in line.upper() and "CM X" not in line.upper() and line != "VAR" and len(line) > 1:
                                # কালার ফিল্টার (যেমন: NERO, VAR ROSA CHIARO)
                                if colour == "N/A":
                                    colour = line
                                elif line not in colour:
                                    colour += " " + line
                        
                        # --- প্যাটার্ন ২: আলাদা কলামে কালার ও সাইজ থাকা (যেমন পৃষ্ঠা ২ এর Blang Sticker) ---
                        # ৪ নম্বর কলামে যদি কালার থাকে
                        if len(row) > 3 and row[3]:
                            c_val = str(row[3]).strip().replace("\n", " ")
                            if any(k in c_val.upper() for k in ["NERO", "ROSA", "BIANCO", "NUDU", "VAR"]):
                                colour = c_val
                                
                        # ৬ নম্বর কলামে যদি সাইজ থাকে
                        if len(row) > 5 and row[5]:
                            s_val = str(row[5]).strip()
                            if s_val in known_sizes:
                                size = s_val
                                
                        # --- প্যাটার্ন ৩: Thermal Sabu Sticker (পৃষ্ঠা ৪-৭) ---
                        # এখানে ৩ নম্বর কলামে 'N/A' থাকে এবং ৫ নম্বর কলামে 'SACV...' কোডটিই সাইজ হিসেবে কাজ করে
                        if colour == "N/A" and len(row) > 3 and str(row[2]).strip() == "N/A":
                            colour = "N/A"
                            if len(row) > 4 and row[3]: # ৪ নম্বর কলামে SACV কোড থাকে
                                size = str(row[3]).strip().replace("\n", "")
                            elif len(row) > 5 and row[4]: # অথবা ৫ নম্বর কলামে
                                size = str(row[4]).strip().replace("\n", "")
                        
                        # ফাইনাল ডেটা স্টোর
                        if style and (colour != "N/A" or size != "N/A"):
                            extracted_rows.append({
                                "STYLE": style,
                                "COLOUR": colour,
                                "SIZE": size,
                                "Quantity": qty
                            })
                            
        # DataFrame তৈরি এবং ডুপ্লিকেট মার্জ করা
        if extracted_rows:
            df_result = pd.DataFrame(extracted_rows)
            
            # একই STYLE, COLOUR, SIZE হলে কোয়ান্টিটি যোগ (Sum) হবে
            df_result = df_result.groupby(["STYLE", "COLOUR", "SIZE"], as_index=False)["Quantity"].sum()
            
            st.subheader("📊 আপনার এক্সেল ডেটার প্রিভিউ:")
            st.dataframe(df_result)
            
            # Excel ফাইল তৈরি (Memory Buffer-এ)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_result.to_excel(writer, sheet_name="Style_Breakdown", index=False)
            output.seek(0)
            
            # ডাউনলোড বাটন
            st.download_button(
                label="📥 Excel ফাইল ডাউনলোড করুন",
                data=output,
                file_name="Style_Breakdown_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("❌ দুঃখিত! এই পিডিএফ-এর টেবিল ফরম্যাটের সাথে লজিক মিলছে না। অনুগ্রহ করে টেক্সট ফরম্যাটটি আরেকবার নিশ্চিত করুন।")

import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

# পেজ সেটআপ
st.set_page_config(page_title="PO Report Generator", layout="centered")
st.title("📄 Purchase Order Report Generator")
st.write("ফাইল আপলোড করুন এবং ম্যাজিক দেখুন।")

# ফাইল আপলোডার
uploaded_files = st.file_uploader("পিডিএফ ফাইলগুলো এখানে দিন", type="pdf", accept_multiple_files=True)

def parse_cotton_club_pdf(file):
    extracted_rows = []
    try:
        with pdfplumber.open(file) as pdf:
            # ১. অর্ডার নম্বর বের করা (পেজ ১ থেকে)
            first_page_text = pdf.pages[0].extract_text() or ""
            order_match = re.search(r'Order no[:\s]+(\d+)', first_page_text, re.IGNORECASE)
            
            short_order_no = "Unknown"
            if order_match:
                full_order = order_match.group(1)
                # শেষের ২ ডিজিট বাদ দেওয়া (যেমন: 17379900 -> 173799)
                short_order_no = full_order[:-2] if len(full_order) > 2 else full_order

            # ২. টেবিল খোঁজা
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    # টেবিল ক্লিন করা
                    clean_table = [[str(cell).replace("\n", " ").strip() if cell else "" for cell in row] for row in table]
                    
                    # --- লজিক: সাইজ হেডার খোঁজা ---
                    size_map = {} # {Column Index: Size Name}
                    header_found = False
                    start_row_index = -1
                    
                    # আমরা খুঁজব এমন রো যেখানে সাইজ আছে (S, M, 3A, 4A ইত্যাদি)
                    for r_idx, row in enumerate(clean_table):
                        # রো-এর ভ্যালুগুলো চেক করি
                        for c_idx, cell in enumerate(row):
                            # কমন সাইজগুলো খুঁজব (লিস্ট আরও বড় করা যেতে পারে)
                            # স্পেস রিমুভ করে চেক করা ভালো
                            clean_cell = cell.replace(" ", "")
                            if clean_cell in ["3A", "4A", "5A", "6A", "8A", "10A", "12A", "S", "M", "L", "XL", "XXL", "3XL", "3M", "6M", "9M", "12M", "18M", "2A"]:
                                size_map[c_idx] = clean_cell
                        
                        # যদি অন্তত ২টা সাইজ পাওয়া যায়, তাহলে এটাই হেডার রো
                        if len(size_map) >= 2:
                            header_found = True
                            start_row_index = r_idx
                            break
                    
                    # যদি হেডার পাওয়া যায়, তবে ডাটা খুঁজব
                    if header_found:
                        for i in range(start_row_index + 1, len(clean_table)):
                            row = clean_table[i]
                            if not row: continue
                            
                            first_cell = row[0]
                            
                            # কালার চেনার উপায়:
                            # ১. টেক্সট হতে হবে
                            # ২. "Total" বা "Spec" শব্দ থাকবে না
                            # ৩. সাধারণত ২ অক্ষরের বেশি হয়
                            is_color_row = False
                            
                            # অনাকাঙ্ক্ষিত রো বাদ দেওয়া
                            bad_keywords = ["Total", "Spec", "Page", "Quantity", "Amount", "Price", "Currency"]
                            if len(first_cell) > 2 and not any(x in first_cell for x in bad_keywords):
                                # কালার রো সাধারণত সংখ্যা দিয়ে শুরু হয় না
                                if not any(char.isdigit() for char in first_cell):
                                    is_color_row = True
                            
                            if is_color_row:
                                row_data = {
                                    "Color": first_cell,
                                    "Order No": short_order_no
                                }
                                
                                total_qty = 0
                                # ম্যাপ করা কলাম থেকে কোয়ান্টিটি নেওয়া
                                for col_idx, size_name in size_map.items():
                                    if col_idx < len(row):
                                        try:
                                            # কমা বা স্পেস থাকলে সরিয়ে ফেলা
                                            val = str(row[col_idx]).replace(",", "").replace(" ", "").replace(".", "")
                                            # যদি ভ্যালু থাকে এবং সংখ্যা হয়
                                            if val.isdigit():
                                                qty = int(val)
                                                # সেফটি চেক: ১ লক্ষের বেশি হলে বাদ (গারবেজ)
                                                if qty > 100000: qty = 0
                                            else:
                                                qty = 0
                                        except:
                                            qty = 0
                                    else:
                                        qty = 0
                                    
                                    row_data[size_name] = qty
                                    total_qty += qty
                                
                                # ম্যানুয়ালি টোটাল বসাচ্ছি
                                row_data["Total"] = total_qty
                                
                                # শুধু যদি কোয়ান্টিটি থাকে তবেই এড করব
                                if total_qty > 0:
                                    extracted_rows.append(row_data)

    except Exception as e:
        st.error(f"Error in {file.name}: {e}")
        
    return extracted_rows

if uploaded_files:
    if st.button("Generate Report Now"):
        all_data = []
        progress_bar = st.progress(0)
        
        for idx, f in enumerate(uploaded_files):
            all_data.extend(parse_cotton_club_pdf(f))
            progress_bar.progress((idx + 1) / len(uploaded_files))
            
        progress_bar.empty()
            
        if all_data:
            df = pd.DataFrame(all_data)
            df = df.fillna(0)
            
            # --- কলাম সাজানো ---
            # ফিক্সড কলাম
            cols = list(df.columns)
            fixed_cols = ["Color", "Order No"]
            
            # সাইজ কলামগুলো আলাদা করা
            size_cols = [c for c in cols if c not in fixed_cols and c != "Total"]
            
            # সাইজ সর্টিং (লজিক: বাচ্চারা আগে, তারপর বড়রা)
            def sort_key(val):
                order = ["3M", "6M", "9M", "12M", "18M", "2A", "3A", "4A", "5A", "6A", "8A", "10A", "12A", "XS", "S", "M", "L", "XL", "XXL", "3XL"]
                return order.index(val) if val in order else 99
            
            size_cols.sort(key=sort_key)
            
            # ফাইনাল অর্ডার: Color -> Order No -> Sizes -> Total
            final_cols = ["Color", "Order No"] + size_cols + ["Total"]
            
            # সেইফটি চেক: ডাটাফ্রেমে সব কলাম আছে কিনা
            available_cols = [c for c in final_cols if c in df.columns]
            df = df[available_cols]
            
            # সর্টিং: কালার আগে
            if "Color" in df.columns and "Order No" in df.columns:
                df = df.sort_values(by=["Color", "Order No"])
            
            # এক্সেল ডাউনলোড
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            excel_data = output.getvalue()
            
            st.success("✅ রিপোর্ট রেডি! নিচে ক্লিক করে ডাউনলোড করুন।")
            st.download_button("📥 ডাউনলোড এক্সেল", data=excel_data, file_name="Final_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            # প্রিভিউ টেবিল
            st.dataframe(df)
        else:
            st.warning("কোনো ডাটা পাওয়া যায়নি। সম্ভবত পিডিএফ ফরম্যাট মিলছে না।")

import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

# পেজ সেটআপ
st.set_page_config(page_title="PDF Report Generator", layout="centered")
st.header("📋 PDF to Excel: Color & Size Report")

# ফাইল আপলোডার
uploaded_files = st.file_uploader("পিডিএফ ফাইলগুলো এখানে আপলোড করুন", type="pdf", accept_multiple_files=True)

def extract_clean_data(file):
    data_list = []
    try:
        with pdfplumber.open(file) as pdf:
            # ১. অর্ডার নম্বর বের করা
            first_page_text = pdf.pages[0].extract_text() or ""
            order_match = re.search(r'Order no:\s*(\d+)', first_page_text)
            
            short_order_no = "Unknown"
            if order_match:
                full_order = order_match.group(1)
                # শেষের ২ ডিজিট বাদ দেওয়া (যেমন: 17379900 -> 173799)
                short_order_no = full_order[:-2] if len(full_order) > 2 else full_order

            # ২. টেবিল থেকে ডাটা নেওয়া
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    # টেবিল ক্লিন করা
                    clean_table = [[str(cell).replace("\n", " ").strip() if cell else "" for cell in row] for row in table]
                    
                    header_found = False
                    size_indices = {}
                    total_idx = -1
                    
                    # হেডার খোঁজা (Total কলাম দিয়ে)
                    for r_idx, row in enumerate(clean_table):
                        # Total কলাম খোঁজা
                        for c_idx, cell in enumerate(row):
                            if "Total" in cell and "Amount" not in cell:
                                total_idx = c_idx
                                break
                        
                        if total_idx != -1:
                            # সাইজ কলাম ম্যাপ করা (Total এর আগের কলামগুলো)
                            for c in range(1, total_idx):
                                col_name = row[c]
                                # অপ্রয়োজনীয় কলাম বাদে শুধু সাইজ নেওয়া
                                if col_name and col_name not in ["Spec. price", "Price", "Color", "Size"]:
                                    size_indices[col_name] = c
                            
                            if size_indices:
                                header_found = True
                                header_row_idx = r_idx
                                break
                    
                    # ডাটা রিড করা
                    if header_found:
                        for i in range(header_row_idx + 1, len(clean_table)):
                            row = clean_table[i]
                            if not row: continue
                            
                            first_cell = row[0]
                            
                            # কালার ফিল্টার: কালারের নাম সাধারণত টেক্সট হয় এবং ২ অক্ষরের বেশি হয়
                            is_color = False
                            if len(first_cell) > 2 and "Total" not in first_cell and "Spec" not in first_cell:
                                is_color = True

                            if is_color:
                                row_data = {
                                    "Color": first_cell,
                                    "Order No": short_order_no
                                }
                                
                                row_total = 0
                                # সাইজ অনুযায়ী কোয়ান্টিটি বসানো
                                for size_name, col_idx in size_indices.items():
                                    try:
                                        val = row[col_idx].replace(",", "").replace(" ", "")
                                        qty = int(float(val)) if val else 0
                                    except:
                                        qty = 0
                                    
                                    row_data[size_name] = qty
                                    row_total += qty
                                
                                # টেবিলের টোটাল নেওয়া
                                try:
                                    t_val = row[total_idx].replace(",", "").replace(" ", "")
                                    final_total = int(float(t_val)) if t_val else row_total
                                except:
                                    final_total = row_total
                                
                                row_data["Total"] = final_total
                                data_list.append(row_data)

    except Exception as e:
        st.error(f"Error in {file.name}: {e}")
        
    return data_list

if uploaded_files:
    if st.button("Generate Report"):
        all_data = []
        for f in uploaded_files:
            all_data.extend(extract_clean_data(f))
            
        if all_data:
            df = pd.DataFrame(all_data)
            df = df.fillna(0)
            
            # --- কলাম সাজানো (ছবি অনুযায়ী) ---
            cols = list(df.columns)
            # ফিক্সড কলাম
            fixed = ["Color", "Order No", "Total"]
            # সাইজ কলাম (বাকি সব)
            sizes = [c for c in cols if c not in fixed]
            
            # সাইজগুলোকে সুন্দর অর্ডারে সাজানো (3A আগে, S পরে)
            def sort_sizes(val):
                order = ["3A", "4A", "5A", "6A", "8A", "10A", "12A", "XS", "S", "M", "L", "XL", "XXL"]
                return order.index(val) if val in order else 99
            
            sizes.sort(key=sort_sizes)
            
            # ফাইনাল কলাম অর্ডার: Color -> Order No -> Sizes -> Total
            final_cols = ["Color", "Order No"] + sizes + ["Total"]
            df = df[final_cols]
            
            # সর্টিং: কালার আগে, তারপর অর্ডার নম্বর
            df = df.sort_values(by=["Color", "Order No"])
            
            # এক্সেল ডাউনলোড
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            excel_data = output.getvalue()
            
            st.success("✅ রিপোর্ট রেডি!")
            st.download_button("📥 ডাউনলোড এক্সেল", data=excel_data, file_name="Color_Wise_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            # প্রিভিউ
            st.dataframe(df)
        else:
            st.warning("কোনো ডাটা পাওয়া যায়নি।")

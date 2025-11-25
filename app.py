import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

# পেজের কনফিগারেশন (টাইটেল এবং লেআউট)
st.set_page_config(page_title="PDF to Excel Converter", layout="wide")

st.title("📄 PDF Purchase Order to Excel Converter")
st.write("আপনার সব পিডিএফ ফাইল নিচে আপলোড করুন। সিস্টেম অটোমেটিক অর্ডার এবং কালার অনুযায়ী রিপোর্ট তৈরি করে দেবে।")

# ফাইল আপলোডার উইজেট
uploaded_files = st.file_uploader("পিডিএফ ফাইলগুলো এখানে ড্রপ করুন (Multiple Allowed)", type="pdf", accept_multiple_files=True)

def extract_data_from_pdf(file):
    data = []
    try:
        with pdfplumber.open(file) as pdf:
            # --- ১. অর্ডার নম্বর বের করা ---
            first_page_text = pdf.pages[0].extract_text() or ""
            order_match = re.search(r'Order no:\s*(\d+)', first_page_text)
            
            short_order_no = "Unknown"
            if order_match:
                full_order = order_match.group(1)
                # শেষের ২ ডিজিট বাদ দেওয়া
                short_order_no = full_order[:-2] if len(full_order) > 2 else full_order

            # --- ২. টেবিল প্রসেসিং ---
            for page in pdf.pages:
                tables = page.extract_tables()
                
                for table in tables:
                    # টেবিল ক্লিন করা (None ভ্যালু সরানো)
                    clean_table = [[str(cell).replace("\n", " ").strip() if cell else "" for cell in row] for row in table]
                    
                    header_row_idx = -1
                    size_indices = {}
                    total_idx = -1
                    
                    # --- হেডার ডিটেকশন লজিক ---
                    # আমরা খুঁজব এমন একটি রো যেখানে 'Total' আছে এবং সাইজের নাম (S, M, 3A, 4A) আছে
                    for r_idx, row in enumerate(clean_table):
                        # Total কলাম খোঁজা
                        for c_idx, cell in enumerate(row):
                            if "Total" in cell and "Amount" not in cell: # Total Quantity বা শুধু Total
                                total_idx = c_idx
                                break
                        
                        if total_idx != -1:
                            # এবার সাইজ কলামগুলো ম্যাপ করি (Total এর আগের কলামগুলো)
                            # সাধারণত ২য় কলাম থেকে Total এর আগ পর্যন্ত সাইজ থাকে
                            for c in range(1, total_idx):
                                col_name = row[c]
                                # সাইজ ফিল্টার (অপ্রয়োজনীয় কলাম বাদ দেওয়া)
                                if col_name and col_name not in ["Spec. price", "Price", "Color", "Size"]:
                                    size_indices[col_name] = c
                            
                            if size_indices:
                                header_row_idx = r_idx
                                break
                    
                    # --- ডাটা এক্সট্রাকশন ---
                    if header_row_idx != -1:
                        # হেডারের পরের রো গুলো চেক করা
                        for i in range(header_row_idx + 1, len(clean_table)):
                            row = clean_table[i]
                            if not row: continue
                            
                            first_cell = row[0]
                            
                            # কালার রো চেনার লজিক
                            # সাধারণত কালার নাম ১মে থাকে এবং লম্বায় ২ অক্ষরের বেশি হয়
                            # এবং এটি হেডার বা টোটাল রো হবে না
                            is_color = False
                            if len(first_cell) > 2 and "Total" not in first_cell and "Spec" not in first_cell and "Page" not in first_cell:
                                is_color = True
                            
                            if is_color:
                                row_data = {
                                    "Color": first_cell,
                                    "Order No": short_order_no,
                                    "File Name": file.name
                                }
                                
                                row_qty_total = 0
                                
                                # ডাইনামিক সাইজ ভ্যালু বসানো
                                for size_name, col_idx in size_indices.items():
                                    try:
                                        val = row[col_idx].replace(",", "").replace(" ", "")
                                        qty = int(float(val)) if val else 0
                                    except:
                                        qty = 0
                                    
                                    row_data[size_name] = qty
                                    row_qty_total += qty
                                
                                # টেবিলের টোটাল ভ্যালু নেওয়া
                                try:
                                    t_val = row[total_idx].replace(",", "").replace(" ", "")
                                    final_total = int(float(t_val)) if t_val else row_qty_total
                                except:
                                    final_total = row_qty_total
                                
                                row_data["Total"] = final_total
                                data.append(row_data)

    except Exception as e:
        st.error(f"Error extracting {file.name}: {str(e)}")
        
    return data

# --- মেইন এক্সিকিউশন ---
if uploaded_files:
    if st.button("Generate Excel Report"):
        with st.spinner('Processing files...'):
            all_data = []
            for pdf_file in uploaded_files:
                file_data = extract_data_from_pdf(pdf_file)
                all_data.extend(file_data)
            
            if all_data:
                df = pd.DataFrame(all_data)
                df = df.fillna(0)
                
                # কলাম সাজানো (Custom Sorting)
                cols = list(df.columns)
                base_cols = ["Color", "Order No"]
                end_cols = ["Total", "File Name"]
                size_cols = [c for c in cols if c not in base_cols and c not in end_cols]
                
                # সাইজগুলোকে লজিক্যালি সাজানো (3A আগে, S পরে)
                def size_sort_key(val):
                    order = ["3A", "4A", "5A", "6A", "8A", "10A", "12A", "XS", "S", "M", "L", "XL", "XXL", "3XL"]
                    if val in order:
                        return order.index(val)
                    return 99
                
                size_cols.sort(key=size_sort_key)
                
                final_cols = base_cols + size_cols + end_cols
                df = df[final_cols]
                
                # মেইন রিকোয়ারমেন্ট: কালার আগে, তারপর অর্ডার নম্বর অনুযায়ী সর্ট
                df = df.sort_values(by=["Color", "Order No"])
                
                # এক্সেল বাফার তৈরি
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                processed_data = output.getvalue()
                
                st.success("রিপোর্ট তৈরি সম্পন্ন!")
                
                # ডাউনলোড বাটন
                st.download_button(
                    label="📥 Download Excel File",
                    data=processed_data,
                    file_name="Final_Order_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # প্রিভিউ দেখানো
                st.subheader("Preview:")
                st.dataframe(df)
                
            else:
                st.warning("কোনো ডাটা পাওয়া যায়নি। পিডিএফ ফাইলগুলো চেক করুন।")

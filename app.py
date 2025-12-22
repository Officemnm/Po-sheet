import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# --- Page Configuration ---
st.set_page_config(
    page_title="PDF Order Extractor",
    page_icon="📄",
    layout="wide"
)

# --- Helper Functions ---

def clean_number(value):
    """
    Clean numeric values with STRICT limits.
    Returns 0 if empty, invalid, or unreasonably large.
    """
    if value is None or value == "":
        return 0
    try:
        # শুধুমাত্র সংখ্যা রাখা
        clean_str = re.sub(r'[^\d]', '', str(value))
        
        # STRICT SAFETY CHECK:
        # 1. যদি ফাঁকা হয় -> 0
        # 2. যদি ৬ ডিজিটের বেশি হয় (মানে ৯,৯৯,৯৯৯ এর বেশি) -> 0 (বারকোড/ভুল ডাটা এড়াতে)
        if not clean_str or len(clean_str) > 6:
            return 0
            
        val = int(clean_str)
        # অতিরিক্ত সেফটি: ১০,০০০ এর বেশি কোয়ান্টিটি এক সাইজে অস্বাভাবিক, 
        # তবুও যদি বড় অর্ডার হয় তাই লিমিট ১,০০,০০০ রাখা হলো। এর বেশি হলে বাদ।
        if val > 100000: 
            return 0
            
        return val
    except ValueError:
        return 0

def clean_color_name(text):
    """Clean color name string."""
    if not text:
        return ""
    text = str(text).replace('\n', ' ')
    # অপ্রয়োজনীয় শব্দ বাদ দেওয়া
    text = text.replace("Spec. price", "").replace("Total Quantity", "").replace("Main purchase price", "")
    # স্পেশাল ক্যারেক্টার ক্লিন করা (শুধুমাত্র অক্ষর, সংখ্যা, হাইফেন ও স্পেস রাখা)
    text = re.sub(r'[^\w\s-]', '', text) 
    return re.sub(' +', ' ', text).strip()

def process_pdf_file(uploaded_file):
    """Process a single uploaded PDF file object."""
    extracted_data = []
    order_no = "Unknown"
    
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            # 1. Extract Order No (Page 1)
            if len(pdf.pages) > 0:
                p1_text = pdf.pages[0].extract_text() or ""
                # Order no প্যাটার্ন খোঁজা
                order_match = re.search(r'Order no[:\s]+(\d+)', p1_text, re.IGNORECASE)
                if order_match:
                    order_no = order_match.group(1)

            # 2. Extract Tables (All Pages)
            for page in pdf.pages:
                tables = page.extract_tables()
                
                for table in tables:
                    if not table: continue
                    
                    # Find Header Row
                    header_row_index = -1
                    size_columns = []
                    
                    for i, row in enumerate(table):
                        # None ভ্যালু হ্যান্ডেল করা
                        row_text = [str(x) if x else "" for x in row]
                        
                        # কলাম হেডার ডিটেকশন (Colo বা Size শব্দ খুঁজবে)
                        if any("Colo" in col or "Size" in col for col in row_text):
                            header_row_index = i
                            for col_idx, col_name in enumerate(row_text):
                                c_name = col_name.replace('\n', ' ').strip()
                                # হেডার থেকে অপ্রয়োজনীয় কলাম বাদ দেওয়া
                                if c_name and "Colo" not in c_name and "Total" not in c_name and "Spec" not in c_name:
                                    size_columns.append({'index': col_idx, 'name': c_name})
                            break
                    
                    if header_row_index == -1 or not size_columns:
                        continue

                    # Process Rows
                    current_color = None
                    
                    for i in range(header_row_index + 1, len(table)):
                        row = table[i]
                        first_col = row[0] if row[0] else ""
                        
                        # --- STRICT FILTERING (গারবেজ ডাটা আটকানোর জন্য) ---
                        
                        # ১. পুরো রো-এর সব টেক্সট চেক করা
                        row_full_text = " ".join([str(x).lower() for x in row if x])
                        
                        # ২. যদি রো-তে Total, Amount, Price, Assortment থাকে -> SKIP
                        if any(bad_word in row_full_text for bad_word in ['total', 'amount', 'assortment', 'main purchase', 'currency']):
                            # এখানে current_color রিসেট করা ভালো, যাতে পরের লাইনে ভুল করে আগের কালার না ধরে
                            current_color = None 
                            continue
                        
                        temp_color = clean_color_name(first_col)
                        
                        # ৩. কালার ডিটেকশন লজিক
                        # যদি টেক্সট থাকে এবং সেটি সংখ্যা না হয়
                        if temp_color and not any(char.isdigit() for char in temp_color):
                            current_color = temp_color
                        elif not temp_color and current_color:
                            # যদি কালার সেল ফাঁকা থাকে, আমরা আগের কালার ধরব কি না?
                            # সাধারণত কোয়ান্টিটি টেবিলের মাঝখানে ফাঁকা রো থাকে না।
                            # তাই যদি কোনো সংখ্যা না পাওয়া যায়, তবে স্কিপ করা ভালো।
                            pass
                        elif not temp_color and not current_color:
                            # কালারও নেই, আগের কালারও নেই -> স্কিপ
                            continue

                        # ডাটা এক্সট্রাকশন
                        row_has_data = False
                        row_data = {'Order No': order_no, 'Color': current_color}
                        
                        qty_found_count = 0
                        for col_info in size_columns:
                            idx = col_info['index']
                            if idx < len(row):
                                val = clean_number(row[idx])
                                if val > 0:
                                    row_has_data = True
                                    qty_found_count += 1
                                row_data[col_info['name']] = val
                            else:
                                row_data[col_info['name']] = 0
                        
                        # ডাটা অ্যাড করা (যদি ভ্যালিড কালার থাকে এবং অন্তত একটি সাইজের ভ্যালিড কোয়ান্টিটি থাকে)
                        if row_has_data and current_color:
                             extracted_data.append(row_data)

    except Exception as e:
        st.error(f"Error processing file: {e}")
        
    return extracted_data

# --- Main App Layout ---

st.title("📊 Professional PDF Order Extractor")
st.markdown("""
<style>
div.stButton > button:first-child {
    background-color: #0099ff;
    color: white;
    font-size: 20px;
    border-radius: 10px;
    padding: 10px 24px;
}
</style>
""", unsafe_allow_html=True)

st.info("Please upload your Purchase Order PDFs below. The app will extract quantities by Size and Color.")

# File Uploader
uploaded_files = st.file_uploader("Upload PDF Files", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Process Files & Generate Report"):
        all_records = []
        progress_bar = st.progress(0)
        
        for idx, file in enumerate(uploaded_files):
            # Process each file
            data = process_pdf_file(file)
            all_records.extend(data)
            progress_bar.progress((idx + 1) / len(uploaded_files))
            
        progress_bar.empty()
        
        if all_records:
            df = pd.DataFrame(all_records)
            df = df.fillna(0)
            
            # --- Views ---
            st.success("✅ Processing Complete!")
            
            tab1, tab2 = st.tabs(["📌 Summary View (Color Wise)", "📋 Detailed Raw Data"])
            
            with tab1:
                st.subheader("Color & Order Summary")
                # Dynamic Pivot
                size_cols = [c for c in df.columns if c not in ['Order No', 'Color']]
                
                # Ensure all size columns are numeric INT32 (to prevent OverflowError in UI)
                for col in size_cols:
                    # errors='coerce' will turn bad strings to NaN, then fillna(0)
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype('int32')

                pivot_df = df.pivot_table(index=['Color', 'Order No'], values=size_cols, aggfunc='sum', fill_value=0)
                st.dataframe(pivot_df, use_container_width=True)
                
            with tab2:
                st.subheader("Extracted Raw Data")
                st.dataframe(df, use_container_width=True)

            # --- Excel Download ---
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Raw Data', index=False)
                pivot_df.to_excel(writer, sheet_name='Color Wise View')
                
            buffer.seek(0)
            
            st.download_button(
                label="📥 Download Excel Report",
                data=buffer,
                file_name="Order_Report_Professional.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        else:
            st.warning("⚠️ No valid data found in the uploaded PDFs.")

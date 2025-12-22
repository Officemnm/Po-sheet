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
    Clean numeric values.
    Returns 0 if empty, invalid, or unreasonably large (likely garbage text/barcodes).
    """
    if value is None or value == "":
        return 0
    try:
        # শুধুমাত্র সংখ্যা রাখা
        clean_str = re.sub(r'[^\d]', '', str(value))
        
        # Safety Check: যদি কোনো সংখ্যা ৯ ডিজিটের বেশি হয় (১০০,০০০,০০০ এর উপরে),
        # তাহলে এটি সম্ভবত কোনো বারকোড বা ফোন নম্বর, কোয়ান্টিটি নয়।
        # এটি OverflowError আটকাবে।
        if not clean_str or len(clean_str) > 9:
            return 0
            
        return int(clean_str)
    except ValueError:
        return 0

def clean_color_name(text):
    """Clean color name string."""
    if not text:
        return ""
    text = text.replace('\n', ' ')
    text = text.replace("Spec. price", "").replace("Total Quantity", "")
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
                order_match = re.search(r'Order no[:\s]+(\d+)', p1_text)
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
                        row_text = [str(x) if x else "" for x in row]
                        
                        if any("Colo/Size" in col or "Size" in col for col in row_text):
                            header_row_index = i
                            for col_idx, col_name in enumerate(row_text):
                                c_name = col_name.replace('\n', ' ').strip()
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
                        
                        temp_color = clean_color_name(first_col)
                        
                        # Logic to identify color row vs data row
                        if temp_color and not any(char.isdigit() for char in temp_color) and "Total" not in temp_color:
                            current_color = temp_color

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
                        
                        if row_has_data and current_color:
                            if "Spec. price" in str(first_col) or qty_found_count > 0:
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
                
                # Ensure all size columns are numeric to prevent Arrow errors
                for col in size_cols:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

                pivot_df = df.pivot_table(index=['Color', 'Order No'], values=size_cols, aggfunc='sum', fill_value=0)
                
                # Display using container width (Fixed Warning syntax too)
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

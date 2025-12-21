from flask import Flask, request, render_template_string
import pdfplumber
import pandas as pd
import os
import re
import shutil
import numpy as np
import traceback

app = Flask(__name__)

# কনফিগারেশন
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# ডিবাগ মোড
DEBUG = True  # Render-এ True করুন

# HTML টেম্পলেটগুলো একই থাকবে (উপরের মতো)
# ...

# ==========================================
#  IMPROVED EXTRACTION LOGIC
# ==========================================

def is_potential_size(header):
    h = str(header).strip().upper()
    if not h:
        return False
    if h in ["COLO", "SIZE", "TOTAL", "QUANTITY", "PRICE", "AMOUNT", "CURRENCY", "ORDER NO", "P.O NO", "COLOR", "COLO/", "COLO /"]:
        return False
    if re.match(r'^\d+$', h): return True
    if re.match(r'^\d+[AMYT]$', h): return True
    if re.match(r'^(XXS|XS|S|M|L|XL|XXL|XXXL|TU|ONE\s*SIZE)$', h): return True
    if re.match(r'^[A-Z]\d{2,}$', h): return False
    return False

def sort_sizes(size_list):
    STANDARD_ORDER = [
        '0M', '1M', '3M', '6M', '9M', '12M', '18M', '24M', '36M',
        '2A', '3A', '4A', '5A', '6A', '8A', '10A', '12A', '14A', '16A', '18A',
        'XXS', 'XS', 'S', 'M', 'L', 'XL', 'XXL', '3XL', '4XL', '5XL',
        'TU', 'One Size'
    ]
    def sort_key(s):
        s = str(s).strip()
        if s in STANDARD_ORDER: return (0, STANDARD_ORDER.index(s))
        if s.isdigit(): return (1, int(s))
        match = re.match(r'^(\d+)([A-Z]+)$', s)
        if match: return (2, int(match.group(1)), match.group(2))
        return (3, s)
    return sorted(size_list, key=sort_key)

def extract_metadata(text):
    meta = {
        'buyer': 'N/A', 'booking': 'N/A', 'style': 'N/A', 
        'season': 'N/A', 'dept': 'N/A', 'item': 'N/A'
    }
    
    text_upper = text.upper()
    
    # Buyer detection
    if "KIABI" in text_upper:
        meta['buyer'] = "KIABI"
    elif "HM" in text_upper:
        meta['buyer'] = "H&M"
    elif "ZARA" in text_upper:
        meta['buyer'] = "ZARA"
    else:
        buyer_patterns = [
            r"Buyer\s*[:\-]?\s*([A-Za-z0-9\s&]+)",
            r"Customer\s*[:\-]?\s*([A-Za-z0-9\s&]+)",
            r"Client\s*[:\-]?\s*([A-Za-z0-9\s&]+)"
        ]
        for pattern in buyer_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                meta['buyer'] = match.group(1).strip()
                break
    
    # Booking number
    booking_patterns = [
        r"Booking\s*No\.?\s*[:\-]?\s*([A-Za-z0-9\-]+)",
        r"Booking\s*Number\s*[:\-]?\s*([A-Za-z0-9\-]+)",
        r"Internal\s*Booking\s*[:\-]?\s*([A-Za-z0-9\-]+)"
    ]
    for pattern in booking_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            meta['booking'] = match.group(1).strip()
            break
    
    # Style
    style_patterns = [
        r"Style\s*Ref\.?\s*[:\-]?\s*([A-Za-z0-9\-]+)",
        r"Style\s*[:\-]?\s*([A-Za-z0-9\-]+)",
        r"Style\s*No\.?\s*[:\-]?\s*([A-Za-z0-9\-]+)"
    ]
    for pattern in style_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            meta['style'] = match.group(1).strip()
            break
    
    # Season
    season_patterns = [
        r"Season\s*[:\-]?\s*([A-Za-z0-9\-]+)",
        r"Season\s*Code\s*[:\-]?\s*([A-Za-z0-9\-]+)"
    ]
    for pattern in season_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            meta['season'] = match.group(1).strip()
            break
    
    # Department
    dept_patterns = [
        r"Dept\.?\s*[:\-]?\s*([A-Za-z]+)",
        r"Department\s*[:\-]?\s*([A-Za-z]+)"
    ]
    for pattern in dept_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            meta['dept'] = match.group(1).strip()
            break
    
    # Item
    item_patterns = [
        r"Garment\s*Item\s*[:\-]?\s*([A-Za-z0-9\s\-]+)",
        r"Item\s*[:\-]?\s*([A-Za-z0-9\s\-]+)",
        r"Product\s*[:\-]?\s*([A-Za-z0-9\s\-]+)"
    ]
    for pattern in item_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            item_text = match.group(1).strip()
            if "Style" in item_text:
                item_text = item_text.split("Style")[0].strip()
            meta['item'] = item_text
            break
    
    return meta

def extract_data_from_pdf(file_path):
    """পিডিএফ থেকে ডেটা এক্সট্র্যাক্ট করে"""
    extracted_data = []
    metadata = {
        'buyer': 'N/A', 'booking': 'N/A', 'style': 'N/A', 
        'season': 'N/A', 'dept': 'N/A', 'item': 'N/A'
    }
    
    try:
        with pdfplumber.open(file_path) as pdf:
            # প্রথম পৃষ্ঠার টেক্সট নিন
            first_page = pdf.pages[0]
            first_page_text = first_page.extract_text()
            
            if DEBUG:
                print(f"=== First page text (first 1000 chars) ===")
                print(first_page_text[:1000])
                print("="*50)
            
            # মেটাডাটা এক্সট্র্যাক্ট করুন
            metadata = extract_metadata(first_page_text)
            
            # বুকিং ফাইল হলে শুধু মেটাডাটা রিটার্ন করুন
            if "MAIN FABRIC BOOKING" in first_page_text.upper() or "FABRIC BOOKING SHEET" in first_page_text.upper():
                if DEBUG:
                    print("This is a Booking file, returning only metadata")
                return [], metadata
            
            # অর্ডার নম্বর খুঁজুন
            order_no = "Unknown"
            order_patterns = [
                r"Order\s*No\.?\s*[:\-]?\s*([A-Za-z0-9\-]+)",
                r"P\.?O\.?\s*No\.?\s*[:\-]?\s*([A-Za-z0-9\-]+)",
                r"Purchase\s*Order\s*[:\-]?\s*([A-Za-z0-9\-]+)"
            ]
            
            for pattern in order_patterns:
                match = re.search(pattern, first_page_text, re.IGNORECASE)
                if match:
                    order_no = match.group(1).strip()
                    if order_no.endswith("00"):
                        order_no = order_no[:-2]
                    break
            
            if DEBUG:
                print(f"Order No: {order_no}")
            
            # সকল পৃষ্ঠায় টেবিল খুঁজুন
            all_tables = []
            for page_num, page in enumerate(pdf.pages):
                try:
                    # টেবিল এক্সট্র্যাক্ট করার চেষ্টা করুন
                    tables = page.extract_tables({
                        "vertical_strategy": "text", 
                        "horizontal_strategy": "text",
                        "explicit_vertical_lines": [],
                        "explicit_horizontal_lines": [],
                        "snap_tolerance": 4,
                        "join_tolerance": 4,
                        "edge_min_length": 3,
                        "min_words_vertical": 1,
                        "min_words_horizontal": 1,
                    })
                    
                    if tables:
                        for table in tables:
                            if table and len(table) > 1:
                                all_tables.append((page_num, table))
                                if DEBUG:
                                    print(f"Found table on page {page_num+1} with {len(table)} rows")
                except Exception as e:
                    if DEBUG:
                        print(f"Error extracting tables from page {page_num+1}: {e}")
            
            # যদি টেবিল পাওয়া যায়
            if all_tables:
                if DEBUG:
                    print(f"Total tables found: {len(all_tables)}")
                
                for page_num, table in all_tables:
                    # টেবিল প্রসেস করুন
                    process_table_data(table, order_no, extracted_data)
            
            # যদি টেবিল না পাওয়া যায়, টেক্সট থেকে ডেটা এক্সট্র্যাক্ট করুন
            if not extracted_data:
                if DEBUG:
                    print("No tables found, trying text extraction")
                
                for page_num, page in enumerate(pdf.pages):
                    text = page.extract_text()
                    if text:
                        extract_data_from_text(text, order_no, extracted_data)
    
    except Exception as e:
        if DEBUG:
            print(f"Error processing PDF {file_path}: {e}")
            traceback.print_exc()
    
    if DEBUG:
        print(f"Total extracted records: {len(extracted_data)}")
        if extracted_data:
            print("Sample records:")
            for i, record in enumerate(extracted_data[:5]):
                print(f"  {i+1}. {record}")
    
    return extracted_data, metadata

def process_table_data(table, order_no, extracted_data):
    """টেবিল ডেটা প্রসেস করুন"""
    try:
        if not table or len(table) < 2:
            return
        
        # হেডার সারি খুঁজুন
        header_row_idx = -1
        size_columns = []
        color_column = -1
        
        for row_idx, row in enumerate(table):
            if row_idx > 5:  # প্রথম ৫ সারির মধ্যে হেডার খুঁজুন
                break
                
            for col_idx, cell in enumerate(row):
                if cell:
                    cell_str = str(cell).strip().upper()
                    if 'COLO' in cell_str or 'COLOR' in cell_str:
                        color_column = col_idx
                        header_row_idx = row_idx
                    elif is_potential_size(cell_str):
                        size_columns.append((col_idx, cell_str))
                        if header_row_idx == -1:
                            header_row_idx = row_idx
        
        if header_row_idx == -1 or not size_columns or color_column == -1:
            if DEBUG:
                print("Could not identify header row")
            return
        
        # সাইজ কলাম সাজান
        size_columns = sorted(size_columns, key=lambda x: sort_sizes([x[1]])[0])
        
        if DEBUG:
            print(f"Header row: {header_row_idx}, Color column: {color_column}")
            print(f"Size columns: {size_columns}")
        
        # ডেটা সারি প্রসেস করুন
        for row_idx in range(header_row_idx + 1, len(table)):
            row = table[row_idx]
            if not row:
                continue
            
            # কালার নাম নিন
            color_name = ""
            if color_column < len(row) and row[color_column]:
                color_name = str(row[color_column]).strip()
            
            if not color_name or color_name.isdigit():
                continue
            
            if DEBUG:
                print(f"Processing row {row_idx}: Color={color_name}")
            
            # প্রতিটি সাইজের জন্য কোয়ান্টিটি নিন
            for col_idx, size_name in size_columns:
                if col_idx < len(row):
                    qty_str = str(row[col_idx]).strip() if row[col_idx] else ""
                    qty = 0
                    
                    try:
                        if qty_str.isdigit():
                            qty = int(qty_str)
                        elif qty_str and re.match(r'^\d+$', qty_str.replace(',', '')):
                            qty = int(qty_str.replace(',', ''))
                    except:
                        qty = 0
                    
                    extracted_data.append({
                        'P.O NO': order_no,
                        'Color': color_name,
                        'Size': size_name,
                        'Quantity': qty
                    })
                    
                    if DEBUG and qty > 0:
                        print(f"  Size: {size_name}, Qty: {qty}")
    
    except Exception as e:
        if DEBUG:
            print(f"Error processing table: {e}")

def extract_data_from_text(text, order_no, extracted_data):
    """টেক্সট থেকে ডেটা এক্সট্র্যাক্ট করুন"""
    lines = text.split('\n')
    sizes = []
    capturing = False
    current_color = ""
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # সাইজ হেডার ডিটেক্ট করুন
        if ("COLO" in line.upper() or "SIZE" in line.upper()) and "TOTAL" in line.upper():
            parts = line.split()
            temp_sizes = []
            for part in parts:
                if is_potential_size(part):
                    temp_sizes.append(part)
            
            if len(temp_sizes) >= 2:  # কমপক্ষে ২টি সাইজ থাকতে হবে
                sizes = temp_sizes
                capturing = True
                if DEBUG:
                    print(f"Found sizes in text: {sizes}")
            continue
        
        if capturing:
            # নতুন কালার লাইন
            if not re.search(r'\d', line) and re.search(r'[A-Za-z]', line):
                current_color = line.strip()
                if DEBUG:
                    print(f"New color: {current_color}")
                continue
            
            # সংখ্যা আছে এমন লাইন
            numbers = re.findall(r'\b\d+\b', line)
            if numbers and current_color:
                # প্রতিটি সাইজের জন্য সংখ্যা অ্যাসাইন করুন
                for i, size in enumerate(sizes):
                    qty = 0
                    if i < len(numbers):
                        qty = int(numbers[i])
                    
                    extracted_data.append({
                        'P.O NO': order_no,
                        'Color': current_color,
                        'Size': size,
                        'Quantity': qty
                    })

# ==========================================
#  FLASK ROUTES
# ==========================================

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if os.path.exists(UPLOAD_FOLDER): 
            shutil.rmtree(UPLOAD_FOLDER)
        os.makedirs(UPLOAD_FOLDER)

        uploaded_files = request.files.getlist('pdf_files')
        all_data = []
        final_meta = {
            'buyer': 'N/A', 'booking': 'N/A', 'style': 'N/A',
            'season': 'N/A', 'dept': 'N/A', 'item': 'N/A'
        }
        
        file_count = 0
        data_count = 0
        
        for file in uploaded_files:
            if file.filename == '': 
                continue
            
            file_count += 1
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(file_path)
            
            if DEBUG:
                print(f"\n{'='*50}")
                print(f"Processing file: {file.filename}")
                print(f"{'='*50}")
            
            data, meta = extract_data_from_pdf(file_path)
            
            # মেটাডাটা আপডেট করুন
            for key in meta:
                if meta[key] != 'N/A' and meta[key]:
                    final_meta[key] = meta[key]
            
            if data:
                data_count += len(data)
                all_data.extend(data)
                
                if DEBUG:
                    print(f"Extracted {len(data)} records from {file.filename}")
        
        if DEBUG:
            print(f"\n{'='*50}")
            print(f"SUMMARY: Processed {file_count} files, extracted {data_count} records")
            print(f"Metadata: {final_meta}")
            print(f"{'='*50}")
        
        if not all_data:
            return render_template_string(RESULT_HTML, 
                                        tables=None, 
                                        message="No PO table data found in the uploaded PDFs. Please check if the PDF contains purchase order tables with color and size information.")
        
        # ডেটা প্রসেস করুন
        try:
            df = pd.DataFrame(all_data)
            df['Color'] = df['Color'].str.strip()
            df = df[df['Color'] != ""]
            
            if df.empty:
                return render_template_string(RESULT_HTML, 
                                            tables=None, 
                                            message="Data extracted but no valid color information found.")
            
            unique_colors = df['Color'].unique()
            
            final_tables = []
            grand_total_qty = 0
            
            for color in unique_colors:
                color_df = df[df['Color'] == color]
                pivot = color_df.pivot_table(index='P.O NO', columns='Size', values='Quantity', aggfunc='sum', fill_value=0)
                
                existing_sizes = pivot.columns.tolist()
                sorted_sizes = sort_sizes(existing_sizes)
                pivot = pivot[sorted_sizes]
                
                pivot['Total'] = pivot.sum(axis=1)
                grand_total_qty += pivot['Total'].sum()
                
                actual_qty = pivot.sum()
                actual_qty.name = 'Actual Qty'
                
                qty_plus_3 = (actual_qty * 1.03).round().astype(int)
                qty_plus_3.name = '3% Order Qty'
                
                pivot = pd.concat([pivot, actual_qty.to_frame().T, qty_plus_3.to_frame().T])
                
                pivot = pivot.reset_index()
                pivot = pivot.rename(columns={'index': 'P.O NO'})
                pivot.columns.name = None
                
                pd.set_option('colheader_justify', 'center')
                table_html = pivot.to_html(classes='table table-bordered table-striped', index=False, border=0)
                
                # Injections
                table_html = re.sub(r'<tr>\s*<td>', '<tr><td class="order-col">', table_html)
                table_html = table_html.replace('<th>Total</th>', '<th class="total-col-header">Total</th>')
                table_html = table_html.replace('<td>Total</td>', '<td class="total-col">Total</td>')
                
                # Color Fix
                table_html = table_html.replace('<td>Actual Qty</td>', '<td class="summary-label">Actual Qty</td>')
                table_html = table_html.replace('<td>3% Order Qty</td>', '<td class="summary-label">3% Order Qty</td>')
                table_html = re.sub(r'<tr>\s*<td class="summary-label">', '<tr class="summary-row"><td class="summary-label">', table_html)
                
                final_tables.append({'color': color, 'table': table_html})
            
            return render_template_string(RESULT_HTML, 
                                        tables=final_tables, 
                                        meta=final_meta, 
                                        grand_total=f"{grand_total_qty:,}")
        
        except Exception as e:
            if DEBUG:
                print(f"Error processing data: {e}")
                traceback.print_exc()
            
            return render_template_string(RESULT_HTML, 
                                        tables=None, 
                                        message=f"Error processing extracted data: {str(e)}")
    
    return render_template_string(INDEX_HTML)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=DEBUG)


from flask import Flask, request, render_template_string
import pdfplumber
import pandas as pd
import os
import re
import shutil
import numpy as np

app = Flask(__name__)

# কনফিগারেশন
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# HTML টেম্পলেটগুলো একই থাকবে (উপরের মতো)
# ...

# ==========================================
#  IMPROVED LOGIC WITH PDFPLUMBER
# ==========================================

def is_potential_size(header):
    h = header.strip().upper()
    if h in ["COLO", "SIZE", "TOTAL", "QUANTITY", "PRICE", "AMOUNT", "CURRENCY", "ORDER NO", "P.O NO"]:
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
        s = s.strip()
        if s in STANDARD_ORDER: return (0, STANDARD_ORDER.index(s))
        if s.isdigit(): return (1, int(s))
        match = re.match(r'^(\d+)([A-Z]+)$', s)
        if match: return (2, int(match.group(1)), match.group(2))
        return (3, s)
    return sorted(size_list, key=sort_key)

def extract_metadata(first_page_text):
    meta = {
        'buyer': 'N/A', 'booking': 'N/A', 'style': 'N/A', 
        'season': 'N/A', 'dept': 'N/A', 'item': 'N/A'
    }
    
    if "KIABI" in first_page_text.upper():
        meta['buyer'] = "KIABI"
    else:
        buyer_match = re.search(r"Buyer.*?Name[\s\S]*?([\w\s&]+)(?:\n|$)", first_page_text)
        if buyer_match: meta['buyer'] = buyer_match.group(1).strip()

    booking_block_match = re.search(r"(?:Internal )?Booking NO\.?[:\s]*([\s\S]*?)(?:System NO|Control No|Buyer)", first_page_text, re.IGNORECASE)
    if booking_block_match: 
        raw_booking = booking_block_match.group(1).strip()
        clean_booking = raw_booking.replace('\n', '').replace('\r', '').replace(' ', '')
        if "System" in clean_booking: clean_booking = clean_booking.split("System")[0]
        meta['booking'] = clean_booking

    style_match = re.search(r"Style Ref\.?[:\s]*([\w-]+)", first_page_text, re.IGNORECASE)
    if style_match: meta['style'] = style_match.group(1).strip()
    else:
        style_match = re.search(r"Style Des\.?[\s\S]*?([\w-]+)", first_page_text, re.IGNORECASE)
        if style_match: meta['style'] = style_match.group(1).strip()

    season_match = re.search(r"Season\s*[:\n\"]*([\w\d-]+)", first_page_text, re.IGNORECASE)
    if season_match: meta['season'] = season_match.group(1).strip()

    dept_match = re.search(r"Dept\.?[\s\n:]*([A-Za-z]+)", first_page_text, re.IGNORECASE)
    if dept_match: meta['dept'] = dept_match.group(1).strip()

    item_match = re.search(r"Garments? Item[\s\n:]*([^\n\r]+)", first_page_text, re.IGNORECASE)
    if item_match: 
        item_text = item_match.group(1).strip()
        if "Style" in item_text: item_text = item_text.split("Style")[0].strip()
        meta['item'] = item_text

    return meta

def extract_table_with_pdfplumber(pdf_path):
    """pdfplumber ব্যবহার করে টেবিল এক্সট্র্যাক্ট করে"""
    tables_data = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # টেবিল ডিটেক্ট করো
            tables = page.extract_tables({
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "explicit_vertical_lines": [],
                "explicit_horizontal_lines": [],
                "snap_tolerance": 3,
                "join_tolerance": 3,
                "edge_min_length": 3,
                "min_words_vertical": 3,
                "min_words_horizontal": 1,
                "keep_blank_chars": False,
                "text_tolerance": 3,
                "text_x_tolerance": 3,
                "text_y_tolerance": 3,
                "intersection_tolerance": 3,
                "intersection_x_tolerance": 3,
                "intersection_y_tolerance": 3,
            })
            
            for table in tables:
                if len(table) > 1:  # শুধু হেডার থাকলে স্টোর করবে না
                    tables_data.extend(table)
    
    return tables_data

def clean_table_data(table_data):
    """টেবিল ডেটা ক্লিনআপ করে"""
    cleaned = []
    for row in table_data:
        cleaned_row = []
        for cell in row:
            if cell is None:
                cleaned_row.append('')
            else:
                # এক্সট্রা স্পেস এবং নিউলাইন রিমুভ করো
                cleaned_cell = ' '.join(str(cell).split())
                cleaned_row.append(cleaned_cell)
        cleaned.append(cleaned_row)
    return cleaned

def extract_data_dynamic(file_path):
    extracted_data = []
    metadata = {
        'buyer': 'N/A', 'booking': 'N/A', 'style': 'N/A', 
        'season': 'N/A', 'dept': 'N/A', 'item': 'N/A'
    }
    order_no = "Unknown"
    
    try:
        # প্রথমে টেক্সট এক্সট্র্যাক্ট করে মেটাডাটা নাও
        with pdfplumber.open(file_path) as pdf:
            first_page = pdf.pages[0]
            first_page_text = first_page.extract_text()
            
            if "Main Fabric Booking" in first_page_text or "Fabric Booking Sheet" in first_page_text:
                metadata = extract_metadata(first_page_text)
                return [], metadata 

            order_match = re.search(r"Order no\D*(\d+)", first_page_text, re.IGNORECASE)
            if order_match: order_no = order_match.group(1)
            else:
                alt_match = re.search(r"Order\s*[:\.]?\s*(\d+)", first_page_text, re.IGNORECASE)
                if alt_match: order_no = alt_match.group(1)
            
            order_no = str(order_no).strip()
            if order_no.endswith("00"): order_no = order_no[:-2]
            
            # টেবিল এক্সট্র্যাক্ট করো
            tables = extract_table_with_pdfplumber(file_path)
            cleaned_tables = clean_table_data(tables)
            
            # টেবিল প্রসেস করো
            for table in cleaned_tables:
                if len(table) < 2:
                    continue
                    
                # হেডার সার্চ করো
                header_row = None
                for i, row in enumerate(table):
                    if any('size' in str(cell).lower() for cell in row if cell) and \
                       any('colo' in str(cell).lower() for cell in row if cell):
                        header_row = i
                        break
                
                if header_row is None:
                    continue
                    
                # সাইজ কলাম খুঁজে বের করো
                header = table[header_row]
                size_columns = []
                color_column = None
                
                for idx, cell in enumerate(header):
                    cell_str = str(cell).strip().upper() if cell else ""
                    if 'COLO' in cell_str or 'COLOR' in cell_str:
                        color_column = idx
                    elif is_potential_size(cell_str):
                        size_columns.append((idx, cell_str))
                
                if not size_columns or color_column is None:
                    continue
                
                # সাইজ কলামগুলো সর্ট করো
                size_columns = sorted(size_columns, key=lambda x: sort_sizes([x[1]])[0])
                
                # ডেটা রো প্রসেস করো
                for row_idx in range(header_row + 1, len(table)):
                    row = table[row_idx]
                    
                    # এম্পটি রো চেক করো
                    if all(cell == '' for cell in row):
                        continue
                    
                    # কালার নেম নাও
                    color_name = str(row[color_column]).strip() if color_column < len(row) else ""
                    if not color_name or color_name.isdigit():
                        continue
                    
                    # প্রতিটি সাইজের জন্য কন্টিটি নাও
                    for col_idx, size_name in size_columns:
                        if col_idx < len(row):
                            qty_str = str(row[col_idx]).strip()
                            if qty_str.isdigit():
                                qty = int(qty_str)
                                extracted_data.append({
                                    'P.O NO': order_no,
                                    'Color': color_name,
                                    'Size': size_name,
                                    'Quantity': qty
                                })
                            elif qty_str == '':
                                # ফাঁকা ঘরের জন্য 0 সেট করো
                                extracted_data.append({
                                    'P.O NO': order_no,
                                    'Color': color_name,
                                    'Size': size_name,
                                    'Quantity': 0
                                })
    
    except Exception as e: 
        print(f"Error processing file {file_path}: {e}")
        # ফলব্যাক হিসেবে পুরানো মেথড
        return fallback_extraction(file_path, metadata)
    
    return extracted_data, metadata

def fallback_extraction(file_path, metadata):
    """পুরানো মেথড হিসেবে ফলব্যাক"""
    extracted_data = []
    order_no = "Unknown"
    
    try:
        with pdfplumber.open(file_path) as pdf:
            first_page = pdf.pages[0]
            first_page_text = first_page.extract_text()
            
            order_match = re.search(r"Order no\D*(\d+)", first_page_text, re.IGNORECASE)
            if order_match: order_no = order_match.group(1)
            else:
                alt_match = re.search(r"Order\s*[:\.]?\s*(\d+)", first_page_text, re.IGNORECASE)
                if alt_match: order_no = alt_match.group(1)
            
            order_no = str(order_no).strip()
            if order_no.endswith("00"): order_no = order_no[:-2]
            
            for page in pdf.pages:
                text = page.extract_text()
                lines = text.split('\n')
                sizes = []
                capturing_data = False
                current_color = None
                
                for i, line in enumerate(lines):
                    line = line.strip()
                    if not line: continue

                    # সাইজ হেডার ডিটেক্ট করো
                    if ("Colo" in line or "Size" in line) and "Total" in line:
                        parts = line.split()
                        try:
                            total_idx = [idx for idx, x in enumerate(parts) if 'Total' in x][0]
                            raw_sizes = parts[:total_idx]
                            temp_sizes = [s for s in raw_sizes if s not in ["Colo", "/", "Size", "Colo/Size", "Colo/", "Size's"]]
                            
                            valid_size_count = sum(1 for s in temp_sizes if is_potential_size(s))
                            if temp_sizes and valid_size_count >= len(temp_sizes) / 2:
                                sizes = temp_sizes
                                capturing_data = True
                            else:
                                sizes = []
                                capturing_data = False
                        except: pass
                        continue
                    
                    if capturing_data:
                        if line.startswith("Total Quantity") or line.startswith("Total Amount"):
                            capturing_data = False
                            continue
                        
                        lower_line = line.lower()
                        if "quantity" in lower_line or "currency" in lower_line or "price" in lower_line or "amount" in lower_line:
                            continue
                            
                        clean_line = line.replace("Spec. price", "").replace("Spec", "").strip()
                        if not re.search(r'[a-zA-Z]', clean_line): continue
                        if re.match(r'^[A-Z]\d+$', clean_line) or "Assortment" in clean_line: continue

                        # কালার লাইন খুঁজো
                        if not re.search(r'\d', line):
                            current_color = clean_line
                            continue
                        
                        # সংখ্যা এক্সট্র্যাক্ট করো
                        numbers = re.findall(r'\b\d+\b', line)
                        if numbers and current_color:
                            # প্রতিটি সাইজের জন্য সংখ্যা অ্যাসাইন করো
                            for size_idx, size in enumerate(sizes):
                                qty = 0
                                if size_idx < len(numbers):
                                    qty = int(numbers[size_idx])
                                
                                extracted_data.append({
                                    'P.O NO': order_no,
                                    'Color': current_color,
                                    'Size': size,
                                    'Quantity': qty
                                })
    except Exception as e:
        print(f"Fallback also failed: {e}")
    
    return extracted_data, metadata

# ==========================================
#  FLASK ROUTES (একই থাকবে)
# ==========================================

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if os.path.exists(UPLOAD_FOLDER): shutil.rmtree(UPLOAD_FOLDER)
        os.makedirs(UPLOAD_FOLDER)

        uploaded_files = request.files.getlist('pdf_files')
        all_data = []
        final_meta = {
            'buyer': 'N/A', 'booking': 'N/A', 'style': 'N/A',
            'season': 'N/A', 'dept': 'N/A', 'item': 'N/A'
        }
        
        for file in uploaded_files:
            if file.filename == '': continue
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(file_path)
            
            data, meta = extract_data_dynamic(file_path)
            
            if meta['buyer'] != 'N/A':
                final_meta = meta
            
            if data:
                all_data.extend(data)
        
        if not all_data:
            return render_template_string(RESULT_HTML, tables=None, message="No PO table data found.")

        df = pd.DataFrame(all_data)
        df['Color'] = df['Color'].str.strip()
        df = df[df['Color'] != ""]
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

    return render_template_string(INDEX_HTML)

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)

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

# ==========================================
# HTML & CSS TEMPLATES (BIG FONT & BOLD)
# ==========================================

INDEX_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Purchase Order Parser</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background-color: #f0f2f5; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .main-card { border: none; border-radius: 15px; box-shadow: 0 10px 25px rgba(0,0,0,0.1); }
        .card-header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border-radius: 15px 15px 0 0 !important; padding: 25px; }
        .btn-upload { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border: none; padding: 12px 30px; font-weight: 600; transition: all 0.3s; }
        .btn-upload:hover { transform: translateY(-2px); box-shadow: 0 5px 15px rgba(118, 75, 162, 0.4); }
        .upload-icon { font-size: 3rem; color: #764ba2; margin-bottom: 20px; }
        .file-input-wrapper { border: 2px dashed #cbd5e0; border-radius: 10px; padding: 40px; background: #f8fafc; transition: all 0.3s; }
        .file-input-wrapper:hover { border-color: #764ba2; background: #fff; }
        .footer-credit { margin-top: 30px; font-size: 0.8rem; color: #6c757d; }
    </style>
</head>
<body>
    <div class="container mt-5">
        <div class="row justify-content-center">
            <div class="col-md-8">
                <div class="card main-card">
                    <div class="card-header text-center">
                        <h2 class="mb-0">PDF Report Generator</h2>
                        <p class="mb-0 opacity-75">Cotton Clothing BD Limited</p>
                    </div>
                    <div class="card-body p-5 text-center">
                        <form action="/" method="post" enctype="multipart/form-data">
                            <div class="file-input-wrapper mb-4">
                                <div class="upload-icon">📂</div>
                                <h5>Select PDF Files</h5>
                                <p class="text-muted small">Select both Booking File & PO Files together</p>
                                <input class="form-control form-control-lg mt-3" type="file" name="pdf_files" multiple accept=".pdf" required>
                            </div>
                            <button type="submit" class="btn btn-primary btn-upload btn-lg w-100">Generate Report</button>
                        </form>
                        <div class="footer-credit">
                            Report Created By <strong>Mehedi Hasan</strong>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
"""

RESULT_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PO Report - Cotton Clothing BD</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background-color: #f8f9fa; padding: 30px 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .container { max-width: 1200px; }
        
        /* Header Styles (Unchanged Size) */
        .company-header { text-align: center; margin-bottom: 20px; border-bottom: 2px solid #000; padding-bottom: 10px; }
        .company-name { font-size: 2.2rem; font-weight: 800; color: #2c3e50; text-transform: uppercase; letter-spacing: 1px; line-height: 1; }
        .report-title { font-size: 1.1rem; color: #555; font-weight: 600; text-transform: uppercase; margin-top: 5px; }
        .date-section { font-size: 1.2rem; font-weight: 800; color: #000; margin-top: 5px; }
        
        /* Info Boxes (Text Size Increased +3pt approx) */
        .info-container { display: flex; justify-content: space-between; margin-bottom: 15px; gap: 15px; }
        
        .info-box { 
            background: white; 
            border: 1px solid #ddd; 
            border-left: 5px solid #2c3e50; 
            padding: 10px 15px; 
            border-radius: 5px; 
            flex: 2; 
            box-shadow: 0 2px 5px rgba(0,0,0,0.05); 
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
        }

        .total-box { 
            background: #2c3e50; 
            color: white; 
            padding: 10px 15px; 
            border-radius: 5px; 
            width: 240px; /* Width increased for bigger text */
            text-align: right; 
            display: flex; 
            flex-direction: column; 
            justify-content: center; 
            box-shadow: 0 4px 10px rgba(44, 62, 80, 0.3); 
        }
        
        /* Increased Font Size for Info Items */
        .info-item { 
            margin-bottom: 6px; 
            font-size: 1.3rem; /* Increased (~13pt/14pt) */
            font-weight: 700; 
            white-space: nowrap; 
            overflow: hidden; 
            text-overflow: ellipsis; 
        }
        
        .info-label { font-weight: 800; color: #444; width: 90px; display: inline-block; }
        .info-value { font-weight: 800; color: #000; }
        
        .total-label { font-size: 1.1rem; opacity: 0.9; text-transform: uppercase; letter-spacing: 1px; font-weight: 700; }
        .total-value { font-size: 2.5rem; font-weight: 800; line-height: 1.1; }

        /* Table Styles */
        .table-card { background: white; border-radius: 0; margin-bottom: 20px; overflow: hidden; border: 1px solid #dee2e6; }
        
        .color-header { 
            background-color: #e9ecef; 
            color: #2c3e50; 
            padding: 10px 12px; 
            font-size: 1.5rem; /* Increased Header Size */
            font-weight: 900; 
            border-bottom: 1px solid #dee2e6; 
            text-transform: uppercase;
        }

        .table { margin-bottom: 0; width: 100%; border-collapse: collapse; }
        
        /* Table Header Cells (Bigger & Bolder) */
        .table th { 
            background-color: #2c3e50; 
            color: white; 
            font-weight: 900; 
            font-size: 1.2rem; /* Increased (~13-14pt) */
            text-align: center; 
            border: 1px solid #34495e; 
            padding: 8px 4px; 
            vertical-align: middle; 
        }
        
        /* Table Data Cells (Bigger & Bolder) */
        .table td { 
            text-align: center; 
            vertical-align: middle; 
            border: 1px solid #dee2e6; 
            padding: 6px 3px; 
            color: #000; 
            font-weight: 800; /* Extra Bold */
            font-size: 1.15rem; /* Increased (~13pt) */
        }
        
        .table-striped tbody tr:nth-of-type(odd) { background-color: #f8f9fa; }
        
        .order-col { font-weight: 900 !important; text-align: center !important; background-color: #fdfdfd; white-space: nowrap; width: 1%; }
        .total-col { font-weight: 900; background-color: #e8f6f3 !important; color: #16a085; border-left: 2px solid #1abc9c !important; }
        .total-col-header { background-color: #e8f6f3 !important; color: #000 !important; font-weight: 900 !important; border: 1px solid #34495e !important; }

        /* SUMMARY ROW STYLES (Light Blue & Bold) */
        .table-striped tbody tr.summary-row,
        .table-striped tbody tr.summary-row td { 
            background-color: #d1ecff !important; 
            --bs-table-accent-bg: #d1ecff !important; 
            color: #000 !important;
            font-weight: 900 !important; /* Extra Bold */
            border-top: 2px solid #aaa !important;
            font-size: 1.2rem !important; /* Even Bigger for Summary */
        }
        
        .summary-label { text-align: right !important; padding-right: 15px !important; color: #000 !important; }

        .action-bar { margin-bottom: 20px; display: flex; justify-content: flex-end; gap: 10px; }
        .btn-print { background-color: #2c3e50; color: white; border-radius: 50px; padding: 8px 30px; font-weight: 600; }
        
        /* Footer Smaller (-1pt) */
        .footer-credit { 
            text-align: center; 
            margin-top: 30px; 
            margin-bottom: 20px; 
            font-size: 0.8rem; /* Reduced Size (~9pt) */
            color: #2c3e50; 
            padding-top: 10px; 
            border-top: 1px solid #ddd; 
        }

        /* =========================================
           PRINT SPECIFIC STYLES
           ========================================= */
        @media print {
            @page { margin: 5mm; size: portrait; }
            
            body { 
                background-color: white; 
                padding: 0; 
                -webkit-print-color-adjust: exact !important; 
                print-color-adjust: exact !important;
                color-adjust: exact !important;
            }
            
            .container { max-width: 100% !important; width: 100% !important; padding: 0; margin: 0; }
            .no-print { display: none !important; }
            
            .company-header { border-bottom: 2px solid #000; margin-bottom: 5px; padding-bottom: 5px; }
            /* Header Size Unchanged in Print */
            .company-name { font-size: 1.8rem; } 
            
            .info-container { margin-bottom: 10px; }
            .info-box { 
                border: 1px solid #000 !important; 
                border-left: 5px solid #000 !important; 
                padding: 5px 10px; 
                display: grid; 
                grid-template-columns: 1fr 1fr;
                gap: 10px;
            }
            .total-box { border: 2px solid #000 !important; background: white !important; color: black !important; padding: 5px 10px; }
            
            /* Info Items Print Size (+3pt) */
            .info-item { font-size: 13pt !important; font-weight: 800 !important; }
            
            /* Table Print Size (+3pt) */
            .table th, .table td { 
                border: 1px solid #000 !important; 
                padding: 2px !important; 
                font-size: 13pt !important; /* ~3pt increased from standard 10pt */
                font-weight: 800 !important;
            }
            
            /* Summary Row Color Force */
            .table-striped tbody tr.summary-row td { 
                background-color: #d1ecff !important; 
                box-shadow: inset 0 0 0 9999px #d1ecff !important; 
                color: #000 !important;
                font-weight: 900 !important;
            }
            
            .color-header { 
                background-color: #f1f1f1 !important; 
                border: 1px solid #000 !important; 
                font-size: 1.4rem !important; /* Bigger */
                font-weight: 900 !important;
                padding: 5px;
                margin-top: 10px;
                box-shadow: inset 0 0 0 9999px #f1f1f1 !important;
            }
            
            .total-col-header {
                background-color: #e8f6f3 !important;
                box-shadow: inset 0 0 0 9999px #e8f6f3 !important;
                color: #000 !important;
            }
            
            .table-card { border: none; margin-bottom: 10px; break-inside: avoid; }
            
            /* Footer Smaller */
            .footer-credit { 
                display: block !important; 
                color: black; 
                border-top: 1px solid #000; 
                margin-top: 10px; 
                font-size: 8pt !important; /* Smaller */
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="action-bar no-print">
            <a href="/" class="btn btn-outline-secondary rounded-pill px-4">Upload New</a>
            <button onclick="window.print()" class="btn btn-print">🖨️ Print Report</button>
        </div>

        <div class="company-header">
            <div class="company-name">Cotton Clothing BD Limited</div>
            <div class="report-title">Purchase Order Summary</div>
            <div class="date-section">Date: <span id="date"></span></div>
        </div>

        {% if message %}
            <div class="alert alert-warning text-center no-print">{{ message }}</div>
        {% endif %}

        {% if tables %}
            <div class="info-container">
                <div class="info-box">
                    <div>
                        <div class="info-item"><span class="info-label">Buyer:</span> <span class="info-value">{{ meta.buyer }}</span></div>
                        <div class="info-item"><span class="info-label">Booking:</span> <span class="info-value">{{ meta.booking }}</span></div>
                        <div class="info-item"><span class="info-label">Style:</span> <span class="info-value">{{ meta.style }}</span></div>
                    </div>
                    <div>
                        <div class="info-item"><span class="info-label">Season:</span> <span class="info-value">{{ meta.season }}</span></div>
                        <div class="info-item"><span class="info-label">Dept:</span> <span class="info-value">{{ meta.dept }}</span></div>
                        <div class="info-item"><span class="info-label">Item:</span> <span class="info-value">{{ meta.item }}</span></div>
                    </div>
                </div>
                
                <div class="total-box">
                    <div class="total-label">Grand Total</div>
                    <div class="total-value">{{ grand_total }}</div>
                    <small>Pieces</small>
                </div>
            </div>

            {% for item in tables %}
                <div class="table-card">
                    <div class="color-header">
                        COLOR: {{ item.color }}
                    </div>
                    <div class="table-responsive">
                        {{ item.table | safe }}
                    </div>
                </div>
            {% endfor %}
            
            <div class="footer-credit">
                Report Created By <strong>Mehedi Hasan</strong>
            </div>
        {% endif %}
    </div>

    <script>
        const dateObj = new Date();
        const day = String(dateObj.getDate()).padStart(2, '0');
        const month = String(dateObj.getMonth() + 1).padStart(2, '0');
        const year = dateObj.getFullYear();
        document.getElementById('date').innerText = `${day}-${month}-${year}`;
    </script>
</body>
</html>
"""

# ==========================================
# LOGIC PART
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
# FLASK ROUTES
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
        
        for file in uploaded_files:
            if file.filename == '': 
                continue
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

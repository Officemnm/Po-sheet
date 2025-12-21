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
#  HTML & CSS TEMPLATES (BIG FONT & BOLD)
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
        .card-header { background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 100%); color: white; border-radius: 15px 15px 0 0 !important; padding: 25px; }
        .btn-upload { background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 100%); border: none; padding: 12px 30px; font-weight: 600; transition: all 0.3s; }
        .btn-upload:hover { transform: translateY(-2px); box-shadow: 0 5px 15px rgba(30, 58, 138, 0.4); }
        .upload-icon { font-size: 3rem; color: #1e3a8a; margin-bottom: 20px; }
        .file-input-wrapper { border: 2px dashed #cbd5e0; border-radius: 10px; padding: 40px; background: #f8fafc; transition: all 0.3s; }
        .file-input-wrapper:hover { border-color: #1e3a8a; background: #fff; }
        .footer-credit { margin-top: 30px; font-size: 0.8rem; color: #6c757d; }
    </style>
</head>
<body>
    <div class="container mt-5">
        <div class="row justify-content-center">
            <div class="col-md-8">
                <div class="card main-card">
                    <div class="card-header text-center">
                        <h2 class="mb-0">PDF Report Generator v2.0</h2>
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
                            <button type="submit" class="btn btn-primary btn-upload btn-lg w-100">Generate Accurate Report</button>
                        </form>
                        <div class="footer-credit">
                            Report Engine Optimized for Empty Cells | Created By <strong>Mehedi Hasan</strong>
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
        .company-header { text-align: center; margin-bottom: 20px; border-bottom: 2px solid #000; padding-bottom: 10px; }
        .company-name { font-size: 2.2rem; font-weight: 800; color: #1e3a8a; text-transform: uppercase; letter-spacing: 1px; }
        .report-title { font-size: 1.1rem; color: #555; font-weight: 600; text-transform: uppercase; }
        .date-section { font-size: 1.2rem; font-weight: 800; color: #000; }
        .info-container { display: flex; justify-content: space-between; margin-bottom: 15px; gap: 15px; }
        .info-box { background: white; border: 1px solid #ddd; border-left: 5px solid #1e3a8a; padding: 15px; border-radius: 5px; flex: 2; display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
        .total-box { background: #1e3a8a; color: white; padding: 15px; border-radius: 5px; width: 260px; text-align: right; display: flex; flex-direction: column; justify-content: center; }
        .info-item { font-size: 1.25rem; font-weight: 700; color: #000; margin-bottom: 5px; }
        .info-label { color: #555; width: 95px; display: inline-block; font-weight: 800; }
        .total-value { font-size: 2.8rem; font-weight: 900; line-height: 1; }
        .color-header { background-color: #1e3a8a; color: white; padding: 12px; font-size: 1.6rem; font-weight: 900; text-transform: uppercase; }
        .table th { background-color: #334155; color: white; font-weight: 900; font-size: 1.1rem; text-align: center; vertical-align: middle; border: 1px solid #000; }
        .table td { text-align: center; vertical-align: middle; border: 1px solid #ccc; padding: 8px; font-weight: 800; font-size: 1.1rem; color: #000; }
        .summary-row td { background-color: #e2e8f0 !important; font-weight: 900 !important; border-top: 2px solid #000 !important; }
        .order-col { background-color: #f8fafc; font-weight: 900 !important; }
        .total-col { background-color: #f1f5f9 !important; font-weight: 900; color: #1e3a8a !important; }
        @media print { .no-print { display: none !important; } .container { max-width: 100% !important; } body { background: white; } }
    </style>
</head>
<body>
    <div class="container">
        <div class="action-bar no-print mb-4 d-flex justify-content-end gap-2">
            <a href="/" class="btn btn-outline-primary rounded-pill px-4">Upload New</a>
            <button onclick="window.print()" class="btn btn-primary rounded-pill px-4">🖨️ Print Report</button>
        </div>

        <div class="company-header">
            <div class="company-name">Cotton Clothing BD Limited</div>
            <div class="report-title">Precise Purchase Order Summary</div>
            <div class="date-section">Date: <span id="date"></span></div>
        </div>

        {% if tables %}
            <div class="info-container">
                <div class="info-box">
                    <div>
                        <div class="info-item"><span class="info-label">Buyer:</span> {{ meta.buyer }}</div>
                        <div class="info-item"><span class="info-label">Booking:</span> {{ meta.booking }}</div>
                        <div class="info-item"><span class="info-label">Style:</span> {{ meta.style }}</div>
                    </div>
                    <div>
                        <div class="info-item"><span class="info-label">Season:</span> {{ meta.season }}</div>
                        <div class="info-item"><span class="info-label">Dept:</span> {{ meta.dept }}</div>
                        <div class="info-item"><span class="info-label">Item:</span> {{ meta.item }}</div>
                    </div>
                </div>
                <div class="total-box">
                    <div style="font-size: 0.9rem; text-transform: uppercase; font-weight: 700;">Grand Total Pieces</div>
                    <div class="total-value">{{ grand_total }}</div>
                </div>
            </div>

            {% for item in tables %}
                <div class="card mb-4 border-0 shadow-sm overflow-hidden">
                    <div class="color-header">COLOR: {{ item.color }}</div>
                    <div class="table-responsive">
                        {{ item.table | safe }}
                    </div>
                </div>
            {% endfor %}
        {% else %}
            <div class="alert alert-danger text-center">{{ message }}</div>
        {% endif %}
    </div>
    <script>
        const d = new Date();
        document.getElementById('date').innerText = `${d.getDate()}-${d.getMonth()+1}-${d.getFullYear()}`;
    </script>
</body>
</html>
"""

# ==========================================
#  ADVANCED LOGIC PART (USING PDFPLUMBER)
# ==========================================

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

def extract_metadata_robust(text):
    meta = {'buyer': 'N/A', 'booking': 'N/A', 'style': 'N/A', 'season': 'N/A', 'dept': 'N/A', 'item': 'N/A'}
    
    # Buyer logic
    if "KIABI" in text.upper(): meta['buyer'] = "KIABI"
    
    # Booking logic
    booking_match = re.search(r"Booking NO\.?[:\s]*([\w-]+)", text, re.IGNORECASE)
    if booking_match: meta['booking'] = booking_match.group(1).strip()
    
    # Style logic
    style_match = re.search(r"Style (?:Ref|Des)\.?[:\s]*([\w-]+)", text, re.IGNORECASE)
    if style_match: meta['style'] = style_match.group(1).strip()
    
    # Season logic
    season_match = re.search(r"Season\s*[:\s]*([\w\d-]+)", text, re.IGNORECASE)
    if season_match: meta['season'] = season_match.group(1).strip()
    
    # Item logic
    item_match = re.search(r"(?:Garments?|Item)[\s\n:]*([A-Za-z\s]+)", text, re.IGNORECASE)
    if item_match: meta['item'] = item_match.group(1).strip().split('\n')[0]

    return meta

def process_pdf_with_plumber(file_path):
    extracted_rows = []
    meta = {}
    
    with pdfplumber.open(file_path) as pdf:
        # ১ম পেজ থেকে মেটাডাটা নিই
        first_page_text = pdf.pages[0].extract_text() or ""
        meta = extract_metadata_robust(first_page_text)
        
        # যদি এটি শুধু বুকিং ফাইল হয় (টেবিল না থাকে), তবে শুধু মেটাডাটা ব্যাক করবে
        if "Main Fabric Booking" in first_page_text:
            return [], meta

        order_no = "N/A"
        order_match = re.search(r"Order (?:no|PO)[\s\.]*(\d+)", first_page_text, re.IGNORECASE)
        if order_match: order_no = order_match.group(1)

        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if not table or len(table) < 2: continue
                
                header_row = None
                size_cols = {} # index: size_name
                
                # টেবিলের হেডার খোঁজা (যেখানে Colo এবং Total আছে)
                for r_idx, row in enumerate(table):
                    row_str = " ".join([str(x) for x in row if x])
                    if ("Colo" in row_str or "Size" in row_str) and "Total" in row_str:
                        header_row = row
                        # সাইজ কলামগুলো শনাক্ত করা
                        for c_idx, cell in enumerate(row):
                            if not cell: continue
                            c_clean = str(cell).strip()
                            if c_clean in ["Colo", "Size", "Colo/Size", "Total", "Quantity", "Price", "Amount", "Currency"]:
                                continue
                            size_cols[c_idx] = c_clean
                        
                        # ডাটা রো প্রসেস করা (হেডারের নিচ থেকে)
                        for data_row in table[r_idx + 1:]:
                            if not data_row or not any(data_row): continue
                            
                            # প্রথম কলাম সাধারণত কালার হয়
                            color_val = str(data_row[0]).strip() if data_row[0] else ""
                            
                            # যদি কালার না থাকে বা শুধু সংখ্যা হয়, তবে এটি কালার রো না হওয়ার সম্ভাবনা বেশি
                            if not color_val or color_val.isdigit() or "Total" in color_val:
                                continue
                                
                            for c_idx, size_name in size_cols.items():
                                if c_idx < len(data_row):
                                    qty_val = data_row[c_idx]
                                    # খালি ঘর হলে ০ ধরবে
                                    try:
                                        qty = int(str(qty_val).strip().replace(',', '')) if qty_val and str(qty_val).strip().isdigit() else 0
                                    except:
                                        qty = 0
                                        
                                    if qty >= 0: # এমনকি ০ হলেও আমরা ডাটা রাখব সঠিক অ্যালাইনমেন্টের জন্য
                                        extracted_rows.append({
                                            'P.O NO': order_no,
                                            'Color': color_val,
                                            'Size': size_name,
                                            'Quantity': qty
                                        })
                        break # এই টেবিল শেষ, পরের টেবিলে যাও
    return extracted_rows, meta

# ==========================================
#  FLASK ROUTES
# ==========================================

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if os.path.exists(UPLOAD_FOLDER): shutil.rmtree(UPLOAD_FOLDER)
        os.makedirs(UPLOAD_FOLDER)

        uploaded_files = request.files.getlist('pdf_files')
        combined_data = []
        final_meta = {'buyer': 'N/A', 'booking': 'N/A', 'style': 'N/A', 'season': 'N/A', 'dept': 'N/A', 'item': 'N/A'}
        
        for file in uploaded_files:
            if not file.filename: continue
            path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(path)
            
            data, meta = process_pdf_with_plumber(path)
            if data: combined_data.extend(data)
            for k, v in meta.items():
                if v != 'N/A': final_meta[k] = v
        
        if not combined_data:
            return render_template_string(RESULT_HTML, tables=None, message="No PO table data found or files are unreadable.", meta=final_meta)

        df = pd.DataFrame(combined_data)
        # ডাটা ক্লিনিং
        df['Color'] = df['Color'].str.strip()
        df = df[df['Color'] != ""]
        
        unique_colors = df['Color'].unique()
        final_tables = []
        grand_total_qty = 0

        for color in unique_colors:
            color_df = df[df['Color'] == color]
            # পিভট টেবিল (সঠিক অ্যালাইনমেন্ট নিশ্চিত করে)
            pivot = color_df.pivot_table(index='P.O NO', columns='Size', values='Quantity', aggfunc='sum', fill_value=0)
            
            # সাইজ সর্টিং
            sorted_cols = sort_sizes(pivot.columns.tolist())
            pivot = pivot[sorted_cols]
            
            # রো এবং কলাম টোটাল
            pivot['Total'] = pivot.sum(axis=1)
            grand_total_qty += pivot['Total'].sum()

            actual_sum = pivot.sum()
            actual_sum.name = 'Actual Qty'
            
            plus_3_sum = (actual_sum * 1.03).round().astype(int)
            plus_3_sum.name = '3% Order Qty'
            
            pivot = pd.concat([pivot, actual_sum.to_frame().T, plus_3_sum.to_frame().T])
            pivot = pivot.reset_index().rename(columns={'index': 'P.O NO'})
            
            # HTML জেনারেট এবং ক্লাসিং
            html = pivot.to_html(classes='table table-bordered table-hover mb-0', index=False, border=0)
            html = html.replace('<td>Actual Qty</td>', '<td class="summary-label">Actual Qty</td>')
            html = html.replace('<td>3% Order Qty</td>', '<td class="summary-label">3% Order Qty</td>')
            html = html.replace('<tr>', '<tr class="data-row">')
            html = html.replace('<tr class="data-row"><td>Actual Qty', '<tr class="summary-row"><td>Actual Qty')
            html = html.replace('<tr class="data-row"><td>3% Order Qty', '<tr class="summary-row"><td>3% Order Qty')
            
            final_tables.append({'color': color, 'table': html})

        return render_template_string(RESULT_HTML, tables=final_tables, meta=final_meta, grand_total=f"{grand_total_qty:,}")

    return render_template_string(INDEX_HTML)

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000, debug=True)

from flask import Flask, request, render_template_string
import pypdf
import pandas as pd
import os
import re
import shutil

app = Flask(__name__)

# কনফিগারেশন
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# ==========================================
#  HTML & CSS TEMPLATES
# ==========================================

# ১. আপলোড পেজের ডিজাইন (অপরিবর্তিত)
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
    </style>
</head>
<body>
    <div class="container mt-5">
        <div class="row justify-content-center">
            <div class="col-md-8">
                <div class="card main-card">
                    <div class="card-header text-center">
                        <h2 class="mb-0">PDF Report Generator</h2>
                        <p class="mb-0 opacity-75">Upload Purchase Orders to extract data</p>
                    </div>
                    <div class="card-body p-5 text-center">
                        <form action="/" method="post" enctype="multipart/form-data">
                            <div class="file-input-wrapper mb-4">
                                <div class="upload-icon">📂</div>
                                <h5>Select PDF Files</h5>
                                <p class="text-muted small">You can select multiple files at once</p>
                                <input class="form-control form-control-lg mt-3" type="file" name="pdf_files" multiple accept=".pdf" required>
                            </div>
                            <button type="submit" class="btn btn-primary btn-upload btn-lg w-100">Generate Report</button>
                        </form>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>
"""

# ২. রেজাল্ট পেজের ডিজাইন (প্রিন্ট স্টাইল সহ)
RESULT_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Purchase Order Report</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background-color: #f8f9fa; padding: 30px 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .container { max-width: 1200px; }
        .header-section { display: flex; justify-content: space-between; align-items: center; margin-bottom: 30px; }
        .btn-action { border-radius: 50px; padding: 8px 25px; font-weight: 600; }
        .table-card { background: white; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); margin-bottom: 30px; overflow: hidden; border: none; }
        .color-header { background-color: #2c3e50; color: white; padding: 15px 20px; font-size: 1.1rem; font-weight: 700; display: flex; align-items: center; justify-content: space-between; }
        .color-badge { background: rgba(255,255,255,0.2); padding: 5px 15px; border-radius: 20px; font-size: 0.9rem; }
        .table-responsive { padding: 0; }
        .table { margin-bottom: 0; }
        .table th { background-color: #ecf0f1; color: #2c3e50; font-weight: 700; text-align: center; border-bottom: 2px solid #bdc3c7; vertical-align: middle; }
        .table td { text-align: center; vertical-align: middle; border-color: #ecf0f1; padding: 12px 8px; font-weight: 500; }
        .table-striped tbody tr:nth-of-type(odd) { background-color: #f9fbfd; }
        .total-col { font-weight: 800; background-color: #e8f6f3 !important; color: #16a085; }
        
        /* প্রিন্ট করার জন্য স্পেশাল CSS */
        @media print {
            body { background-color: white; padding: 0; }
            .container { max-width: 100%; width: 100%; padding: 0; }
            .no-print { display: none !important; } /* বাটন হাইড করার জন্য */
            .table-card { box-shadow: none; border: 1px solid #ddd; margin-bottom: 20px; break-inside: avoid; }
            .color-header { background-color: #f1f1f1 !important; color: black !important; border-bottom: 2px solid #000; -webkit-print-color-adjust: exact; }
            .color-badge { border: 1px solid #000; color: black; }
            .total-col { background-color: #f0f0f0 !important; color: black !important; -webkit-print-color-adjust: exact; }
            a { text-decoration: none; color: black; }
            .table th, .table td { padding: 8px; font-size: 12px; border: 1px solid #ddd; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header-section no-print">
            <div>
                <h2 class="text-dark fw-bold mb-0">Generated Report</h2>
                <small class="text-muted">Review the extracted data below</small>
            </div>
            <div class="d-flex gap-2">
                <button onclick="window.print()" class="btn btn-success btn-action">🖨️ Print Report</button>
                <a href="/" class="btn btn-outline-primary btn-action">Upload New</a>
            </div>
        </div>

        <div class="d-none d-print-block text-center mb-4">
            <h2>Purchase Order Summary</h2>
            <p>Generated on: <span id="date"></span></p>
        </div>

        {% if message %}
            <div class="alert alert-warning p-4 text-center no-print">
                <h4>⚠️ Notice</h4>
                <p class="mb-0">{{ message }}</p>
            </div>
        {% endif %}

        {% if tables %}
            {% for item in tables %}
                <div class="table-card">
                    <div class="color-header">
                        <span>COLOR: {{ item.color }}</span>
                    </div>
                    <div class="table-responsive">
                        {{ item.table | safe }}
                    </div>
                </div>
            {% endfor %}
        {% else %}
            {% if not message %}
                <div class="alert alert-info text-center no-print">Processing complete but no tables generated.</div>
            {% endif %}
        {% endif %}
        
        <div class="text-center mt-5 no-print">
             <p class="text-muted small">Generated by Purchase Order Parser AI</p>
        </div>
    </div>

    <script>
        // বর্তমান তারিখ দেখানোর জন্য
        document.getElementById('date').innerText = new Date().toLocaleDateString();
    </script>
</body>
</html>
"""

# ==========================================
#  লজিক অংশ (LOGIC PART)
# ==========================================

def is_potential_size(header):
    """চেক করে যে টেক্সটটি সাইজ হওয়ার মতো দেখতে কিনা"""
    h = header.strip().upper()
    if h in ["COLO", "SIZE", "TOTAL", "QUANTITY", "PRICE", "AMOUNT", "CURRENCY"]:
        return False
    # প্যাটার্ন ১: শুধু সংখ্যা
    if re.match(r'^\d+$', h): return True
    # প্যাটার্ন ২: সংখ্যা + A/M/Y/T
    if re.match(r'^\d+[AMYT]$', h): return True
    # প্যাটার্ন ৩: অক্ষর সাইজ
    if re.match(r'^(XXS|XS|S|M|L|XL|XXL|XXXL|TU|ONE\s*SIZE)$', h): return True
    # Assortment কোড বাদ
    if re.match(r'^[A-Z]\d{2,}$', h): return False
    return False

def sort_sizes(size_list):
    """সাইজগুলোকে ছোট থেকে বড় সাজানোর লজিক"""
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

def extract_data_dynamic(file_path):
    extracted_data = []
    order_no = "Unknown"
    
    try:
        reader = pypdf.PdfReader(file_path)
        first_page_text = reader.pages[0].extract_text()
        
        # অর্ডার নম্বর বের করা
        order_match = re.search(r"Order no\D*(\d+)", first_page_text, re.IGNORECASE)
        if order_match: 
            order_no = order_match.group(1)
        else:
            alt_match = re.search(r"Order\s*[:\.]?\s*(\d+)", first_page_text, re.IGNORECASE)
            if alt_match: 
                order_no = alt_match.group(1)
        
        # ==========================================
        # লজিক: Order No থেকে শেষের '00' বাদ দেওয়া
        # ==========================================
        order_no = str(order_no).strip()
        if order_no.endswith("00"):
            order_no = order_no[:-2]

        for page in reader.pages:
            text = page.extract_text()
            lines = text.split('\n')
            sizes = []
            capturing_data = False
            
            for i, line in enumerate(lines):
                line = line.strip()
                if not line: continue

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

                    numbers_in_line = re.findall(r'\b\d+\b', line)
                    quantities = [int(n) for n in numbers_in_line]
                    color_name = clean_line
                    final_qtys = []

                    if len(quantities) >= len(sizes):
                        if len(quantities) == len(sizes) + 1: final_qtys = quantities[:-1] 
                        else: final_qtys = quantities[:len(sizes)]
                        color_name = re.sub(r'\s\d+$', '', color_name).strip()
                    elif len(quantities) < len(sizes): 
                        vertical_qtys = []
                        for next_line in lines[i+1:]:
                            next_line = next_line.strip()
                            if "Total" in next_line or re.search(r'[a-zA-Z]', next_line.replace("Spec", "").replace("price", "")): break
                            if re.match(r'^\d+$', next_line): vertical_qtys.append(int(next_line))
                        if len(vertical_qtys) >= len(sizes): final_qtys = vertical_qtys[:len(sizes)]
                    
                    if final_qtys and color_name:
                         for idx, size in enumerate(sizes):
                            extracted_data.append({
                                'Order No': order_no,
                                'Color': color_name,
                                'Size': size,
                                'Quantity': final_qtys[idx]
                            })
    except Exception as e: print(f"Error processing file: {e}")
    return extracted_data

# ==========================================
#  FLASK ROUTES
# ==========================================

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if os.path.exists(UPLOAD_FOLDER): shutil.rmtree(UPLOAD_FOLDER)
        os.makedirs(UPLOAD_FOLDER)

        uploaded_files = request.files.getlist('pdf_files')
        all_data = []
        
        for file in uploaded_files:
            if file.filename == '': continue
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(file_path)
            all_data.extend(extract_data_dynamic(file_path))
        
        if not all_data:
            return render_template_string(RESULT_HTML, tables=None, message="No valid data found in PDFs.")

        df = pd.DataFrame(all_data)
        df['Color'] = df['Color'].str.strip()
        df = df[df['Color'] != ""]
        unique_colors = df['Color'].unique()
        
        final_tables = []
        for color in unique_colors:
            color_df = df[df['Color'] == color]
            pivot = color_df.pivot_table(index='Order No', columns='Size', values='Quantity', aggfunc='sum', fill_value=0)
            
            existing_sizes = pivot.columns.tolist()
            sorted_sizes = sort_sizes(existing_sizes)
            pivot = pivot[sorted_sizes]
            pivot['Total Qty'] = pivot.sum(axis=1)
            
            # HTML টেবিল স্টাইল
            table_html = pivot.to_html(classes='table table-bordered table-striped', border=0)
            final_tables.append({'color': color, 'table': table_html})
            
        return render_template_string(RESULT_HTML, tables=final_tables)

    return render_template_string(INDEX_HTML)

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)

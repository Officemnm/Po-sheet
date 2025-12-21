import os
import re
import shutil
import pdfplumber
import pandas as pd
from flask import Flask, request, render_template_string

app = Flask(__name__)

# কনফিগারেশন
UPLOAD_FOLDER = '/tmp/uploads'  # ক্লাউড সার্ভারের জন্য /tmp ফোল্ডার নিরাপদ
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# ==========================================
#  UI TEMPLATES
# ==========================================

INDEX_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>KIABI Parser Pro</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background: #f8f9fa; font-family: 'Segoe UI', sans-serif; height: 100vh; display: flex; align-items: center; justify-content: center; }
        .upload-card { background: white; padding: 40px; border-radius: 20px; box-shadow: 0 10px 40px rgba(0,0,0,0.1); width: 100%; max-width: 500px; text-align: center; }
        .icon { font-size: 60px; margin-bottom: 20px; }
        .btn-primary { background: #0d6efd; border: none; padding: 12px 30px; font-weight: 600; width: 100%; margin-top: 20px; }
    </style>
</head>
<body>
    <div class="upload-card">
        <div class="icon">📂</div>
        <h3 class="mb-3">KIABI Report Generator</h3>
        <p class="text-muted">Upload Booking & PO PDF files</p>
        <form action="/" method="post" enctype="multipart/form-data">
            <input class="form-control form-control-lg" type="file" name="pdf_files" multiple accept=".pdf" required>
            <button type="submit" class="btn btn-primary">Generate Report</button>
        </form>
    </div>
</body>
</html>
"""

RESULT_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>PO Summary</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { padding: 40px; background: #fff; font-family: 'Segoe UI', sans-serif; color: #000; }
        .header { text-align: center; border-bottom: 3px solid #000; padding-bottom: 20px; margin-bottom: 30px; }
        .meta-box { border: 1px solid #000; padding: 20px; margin-bottom: 30px; background: #f8f9fa; }
        .meta-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; font-weight: 600; font-size: 1.1rem; }
        .grand-total { background: #000; color: #fff; padding: 15px; text-align: center; font-size: 2rem; font-weight: 800; margin-bottom: 40px; border-radius: 8px; }
        .color-section { margin-bottom: 40px; break-inside: avoid; }
        .color-header { background: #0d6efd; color: white; padding: 10px 20px; font-weight: 800; font-size: 1.2rem; text-transform: uppercase; margin-bottom: 0; }
        .table { margin-bottom: 0; border: 1px solid #000; }
        .table th { background: #343a40 !important; color: white !important; text-align: center; border: 1px solid #000; }
        .table td { text-align: center; border: 1px solid #000; font-weight: 700; font-size: 1.1rem; vertical-align: middle; }
        .summary-row td { background: #e9ecef !important; font-weight: 900 !important; border-top: 2px solid #000 !important; }
        @media print { .no-print { display: none; } body { padding: 0; } .grand-total { background: #000 !important; -webkit-print-color-adjust: exact; } .color-header { background: #0d6efd !important; -webkit-print-color-adjust: exact; } }
    </style>
</head>
<body>
    <div class="container">
        <div class="no-print text-end mb-4">
            <a href="/" class="btn btn-outline-dark rounded-pill">Upload Again</a>
            <button onclick="window.print()" class="btn btn-primary rounded-pill px-4 ms-2">Print Report</button>
        </div>

        <div class="header">
            <h1 class="fw-bold">Cotton Clothing BD Limited</h1>
            <h4 class="text-uppercase text-muted">Purchase Order Summary</h4>
        </div>

        <div class="meta-box">
            <div class="meta-grid">
                <div>Buyer: {{ meta.buyer }}</div>
                <div>Season: {{ meta.season }}</div>
                <div>Booking: {{ meta.booking }}</div>
                <div>Dept: {{ meta.dept }}</div>
                <div>Style: {{ meta.style }}</div>
                <div>Item: {{ meta.item }}</div>
            </div>
        </div>

        <div class="grand-total">GRAND TOTAL: {{ grand_total }} PCS</div>

        {% for item in tables %}
            <div class="color-section">
                <div class="color-header">COLOR: {{ item.color }}</div>
                {{ item.table | safe }}
            </div>
        {% endfor %}
    </div>
</body>
</html>
"""

# ==========================================
#  ROBUST PARSING LOGIC
# ==========================================

def clean_qty(val):
    if not val: return 0
    # সংখ্যা বাদে সব মুছে ফেলা
    digits = re.sub(r'[^\d]', '', str(val))
    return int(digits) if digits else 0

def process_file(path):
    rows = []
    meta = {'buyer': 'N/A', 'booking': 'N/A', 'style': 'N/A', 'season': 'N/A', 'dept': 'N/A', 'item': 'N/A'}
    order_no = "N/A"

    with pdfplumber.open(path) as pdf:
        # মেটাডাটা এক্সট্রাকশন
        p1_text = pdf.pages[0].extract_text() or ""
        
        if "KIABI" in p1_text.upper(): meta['buyer'] = "KIABI"
        
        m_bk = re.search(r"Booking NO\.?[:\s]*([\w-]+)", p1_text, re.I)
        if m_bk: meta['booking'] = m_bk.group(1)
        
        m_st = re.search(r"Style (?:Ref|Des)\.?[:\s]*([\w-]+)", p1_text, re.I)
        if m_st: meta['style'] = m_st.group(1)
        
        m_se = re.search(r"Season\s*[:\s]*([\w\d-]+)", p1_text, re.I)
        if m_se: meta['season'] = m_se.group(1)
        
        m_dp = re.search(r"Dept\.?[\s\n:]*([\w\d]+)", p1_text, re.I)
        if m_dp: meta['dept'] = m_dp.group(1)
        
        m_it = re.search(r"(?:Garments?|Item)[\s\n:]*([A-Za-z\s]+)", p1_text, re.I)
        if m_it: meta['item'] = m_it.group(1).split('\n')[0].strip()

        m_ord = re.search(r"Order no: (\d+)", p1_text, re.I)
        if m_ord: order_no = m_ord.group(1)
        
        if "Main Fabric Booking" in p1_text:
            return [], meta

        for page in pdf.pages:
            words = page.extract_words()
            if not words: continue

            # লাইন গ্রুপিং (y-axis tolerance সহ)
            lines = []
            current_line = []
            current_y = None
            
            # উপর থেকে নিচে সর্ট করা
            sorted_words = sorted(words, key=lambda w: w['top'])
            
            for w in sorted_words:
                if current_y is None:
                    current_y = w['top']
                    current_line.append(w)
                elif abs(w['top'] - current_y) <= 3: # 3 পিক্সেল টলারেন্স
                    current_line.append(w)
                else:
                    lines.append(sorted(current_line, key=lambda x: x['x0']))
                    current_line = [w]
                    current_y = w['top']
            if current_line:
                lines.append(sorted(current_line, key=lambda x: x['x0']))

            # হেডার এবং কলাম শনাক্ত করা
            size_cols = []
            header_found = False
            
            for line_words in lines:
                line_text = " ".join([w['text'] for w in line_words]).upper()
                
                # হেডার লাইন
                if "COLO/SIZE" in line_text and "TOTAL" in line_text:
                    header_found = True
                    for w in line_words:
                        txt = w['text'].strip().upper()
                        if txt not in ["COLO/SIZE", "TOTAL", "PRICE", "AMOUNT", "CURRENCY", "COLOR/SIZE"]:
                            # কলামের এরিয়া সেভ করা (একটু বাফার সহ)
                            size_cols.append({
                                'name': w['text'],
                                'x0': w['x0'] - 5,
                                'x1': w['x1'] + 5
                            })
                    continue

                if header_found and size_cols:
                    # টেবিল শেষ চেক
                    if "TOTAL QUANTITY" in line_text or "TOTAL AMOUNT" in line_text:
                        break
                        
                    # কালার নাম বের করা (প্রথম সাইজ কলামের বামে যা আছে)
                    first_col_x = size_cols[0]['x0']
                    color_words = [w['text'] for w in line_words if w['x1'] < first_col_x]
                    color_name = " ".join(color_words).replace("Spec. price", "").strip()
                    
                    if color_name and not color_name[0].isdigit():
                        for col in size_cols:
                            # এই কলামের নিচে কোনো সংখ্যা আছে কি?
                            cell_val_words = [w['text'] for w in line_words if w['x0'] >= col['x0'] and w['x1'] <= col['x1']]
                            cell_val = "".join(cell_val_words)
                            
                            qty = clean_qty(cell_val)
                            
                            # যদি সংখ্যা থাকে তবেই যোগ হবে (0 হলেও যোগ হবে অ্যালাইনমেন্টের জন্য)
                            rows.append({
                                'P.O NO': order_no,
                                'Color': color_name,
                                'Size': col['name'],
                                'Quantity': qty
                            })

    return rows, meta

# ==========================================
#  ROUTES
# ==========================================

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if os.path.exists(app.config['UPLOAD_FOLDER']):
            shutil.rmtree(app.config['UPLOAD_FOLDER'])
        os.makedirs(app.config['UPLOAD_FOLDER'])

        files = request.files.getlist('pdf_files')
        all_data = []
        final_meta = {}
        
        for f in files:
            if not f.filename: continue
            p = os.path.join(app.config['UPLOAD_FOLDER'], f.filename)
            f.save(p)
            d, m = process_file(p)
            if d: all_data.extend(d)
            if m.get('buyer') != 'N/A': final_meta.update(m)
        
        if not all_data:
            return render_template_string(INDEX_HTML) # ডাটা না পেলে আবার আপলোড পেজে

        df = pd.DataFrame(all_data)
        # একই কালার-সাইজের ডাটা যোগ করা
        df = df.groupby(['P.O NO', 'Color', 'Size'])['Quantity'].sum().reset_index()
        
        # সাইজ সর্টিং
        def sort_key(s):
            order = ['XXS','XS','S','M','L','XL','XXL','3XL','4XL','5XL','TU','ONE SIZE']
            u = s.upper().strip()
            return order.index(u) if u in order else 99

        final_tables = []
        grand_total = 0

        for color in df['Color'].unique():
            c_df = df[df['Color'] == color]
            pivot = c_df.pivot_table(index='P.O NO', columns='Size', values='Quantity', aggfunc='sum', fill_value=0)
            
            # কলাম সাজানো
            cols = sorted(pivot.columns.tolist(), key=sort_key)
            pivot = pivot[cols]
            
            pivot['Total'] = pivot.sum(axis=1)
            grand_total += pivot['Total'].sum()

            # টোটাল রো
            act = pivot.sum(); act.name = 'Actual Qty'
            p3 = (act * 1.03).round().astype(int); p3.name = '3% Order Qty'
            
            pivot = pd.concat([pivot, act.to_frame().T, p3.to_frame().T])
            pivot = pivot.reset_index().rename(columns={'index': 'P.O NO'})
            
            html = pivot.to_html(classes='table table-bordered table-hover', index=False)
            
            # স্টাইলিং ইনজেকশন
            html = html.replace('<td>Actual Qty</td>', '<td class="summary-row">Actual Qty</td>')
            html = html.replace('<td>3% Order Qty</td>', '<td class="summary-row">3% Order Qty</td>')
            html = html.replace('<tr>', '<tr>') # ক্লিনআপ
            html = html.replace('<tr><td class="summary-row">Actual Qty', '<tr class="summary-row"><td class="summary-row">Actual Qty')
            html = html.replace('<tr><td class="summary-row">3% Order Qty', '<tr class="summary-row"><td class="summary-row">3% Order Qty')
            
            final_tables.append({'color': color, 'table': html})

        return render_template_string(RESULT_HTML, tables=final_tables, meta=final_meta, grand_total=f"{grand_total:,}")

    return render_template_string(INDEX_HTML)

if __name__ == "__main__":
    # লোকাল ডেভেলপমেন্টের জন্য
    app.run(host='0.0.0.0', port=5000, debug=True)

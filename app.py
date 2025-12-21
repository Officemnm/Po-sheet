from flask import Flask, request, render_template_string
import pdfplumber
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
#  UI TEMPLATES
# ==========================================

INDEX_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>KIABI Precise Parser</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background: #f0f2f5; font-family: sans-serif; }
        .card { margin-top: 50px; border: none; border-radius: 15px; box-shadow: 0 10px 30px rgba(0,0,0,0.1); }
        .header { background: #1e3a8a; color: white; padding: 25px; border-radius: 15px 15px 0 0; }
        .upload-box { border: 2px dashed #1e3a8a; padding: 40px; border-radius: 10px; background: #fff; }
    </style>
</head>
<body>
    <div class="container mt-5">
        <div class="row justify-content-center">
            <div class="col-md-7">
                <div class="card shadow-lg">
                    <div class="header text-center">
                        <h2 class="fw-bold">KIABI PO REPORT GENERATOR</h2>
                        <p class="mb-0">Cotton Clothing BD Limited | Precision Engine v3.0</p>
                    </div>
                    <div class="card-body p-5 text-center">
                        <form action="/" method="post" enctype="multipart/form-data">
                            <div class="upload-box mb-4">
                                <div style="font-size: 4rem;">📁</div>
                                <h5 class="mt-2">সিলেক্ট করুন বুকিং এবং পিও ফাইল</h5>
                                <p class="text-muted small">Select multiple PDF files at once</p>
                                <input class="form-control mt-3" type="file" name="pdf_files" multiple accept=".pdf" required>
                            </div>
                            <button type="submit" class="btn btn-primary btn-lg w-100 shadow-sm">রিপোর্ট জেনারেট করুন</button>
                        </form>
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
    <title>PO Summary Report</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { padding: 30px; font-family: 'Segoe UI', sans-serif; color: #000; }
        .report-header { text-align: center; border-bottom: 3px solid #000; padding-bottom: 10px; margin-bottom: 20px; }
        .meta-table { width: 100%; border: 1px solid #000; margin-bottom: 20px; background: #fff; }
        .meta-table td { padding: 10px; font-weight: 800; border: 1px solid #000; font-size: 1.1rem; }
        .color-title { background: #1e3a8a; color: white; padding: 12px; font-weight: 900; margin-top: 35px; border-left: 10px solid #000; }
        .table th { background: #334155 !important; color: white !important; border: 1px solid #000 !important; text-align: center; font-weight: 900; }
        .table td { border: 1px solid #000 !important; text-align: center; font-weight: 800; font-size: 1.15rem; }
        .summary-row td { background: #d1ecff !important; border-top: 2px solid #000 !important; color: #000; }
        .grand-total { background: #1e3a8a; color: white; padding: 15px; text-align: center; font-size: 2.2rem; font-weight: 900; box-shadow: 0 4px 10px rgba(0,0,0,0.2); }
        .print-btn { background: #1e3a8a; color: #fff; font-weight: bold; border-radius: 50px; padding: 10px 30px; }
        @media print { .no-print { display: none; } @page { margin: 10mm; } }
    </style>
</head>
<body>
    <div class="container-fluid">
        <div class="no-print text-end mb-4">
            <a href="/" class="btn btn-outline-dark rounded-pill px-4">New Upload</a>
            <button onclick="window.print()" class="btn print-btn px-4 shadow-sm">🖨️ Print Report</button>
        </div>

        <div class="report-header">
            <h1 class="display-5 fw-bold" style="color: #1e3a8a;">Cotton Clothing BD Limited</h1>
            <h4 class="text-uppercase tracking-wider">Purchase Order Summary Report</h4>
        </div>

        <table class="meta-table shadow-sm">
            <tr>
                <td width="50%">Buyer: {{ meta.buyer }}</td>
                <td>Season: {{ meta.season }}</td>
            </tr>
            <tr>
                <td>Booking: {{ meta.booking }}</td>
                <td>Dept: {{ meta.dept }}</td>
            </tr>
            <tr>
                <td>Style: {{ meta.style }}</td>
                <td>Item: {{ meta.item }}</td>
            </tr>
        </table>

        <div class="grand-total mb-4">GRAND TOTAL: {{ grand_total }} Pieces</div>

        {% for item in tables %}
            <div class="color-title text-uppercase">COLOR: {{ item.color }}</div>
            <div class="table-responsive shadow-sm">
                {{ item.table | safe }}
            </div>
        {% endfor %}
        
        <div class="mt-5 text-center border-top pt-3 no-print">
            <p class="text-muted">Software Engine Optimized for Precise Cell Matching | Report Created by Mehedi Hasan</p>
        </div>
    </div>
</body>
</html>
"""

# ==========================================
#  PRECISE COORDINATE-BASED LOGIC
# ==========================================

def process_kiabi_pdf(path):
    rows = []
    meta = {'buyer': 'N/A', 'booking': 'N/A', 'style': 'N/A', 'season': 'N/A', 'dept': 'N/A', 'item': 'N/A'}
    order_no = "N/A"

    with pdfplumber.open(path) as pdf:
        # ১ম পাতা থেকে মেটাডাটা
        text_first = pdf.pages[0].extract_text() or ""
        
        if "KIABI" in text_first.upper(): meta['buyer'] = "KIABI"
        
        # Booking No.
        m_book = re.search(r"Booking NO\.?[:\s]*([\w-]+)", text_first, re.I)
        if m_book: meta['booking'] = m_book.group(1)
        
        # Style
        m_style = re.search(r"Style (?:Ref|Des)\.?[:\s]*([\w-]+)", text_first, re.I)
        if m_style: meta['style'] = m_style.group(1)
        
        # Season
        m_season = re.search(r"Season\s*[:\s]*([\w\d-]+)", text_first, re.I)
        if m_season: meta['season'] = m_season.group(1)
        
        # Dept
        m_dept = re.search(r"Dept\s*[:\s]*([\w\d-]+)", text_first, re.I)
        if m_dept: meta['dept'] = m_dept.group(1)

        # Order No
        m_order = re.search(r"Order no: (\d+)", text_first, re.I)
        if m_order: order_no = m_order.group(1)
        
        if "Main Fabric Booking" in text_first:
            return [], meta

        for page in pdf.pages:
            words = page.extract_words()
            if not words: continue

            # রো অনুযায়ী সাজানো
            y_lines = {}
            for w in words:
                y = round(w['top'], 0)
                if y not in y_lines: y_lines[y] = []
                y_lines[y].append(w)
            
            sorted_ys = sorted(y_lines.keys())
            
            size_cols = []
            header_y = -1
            
            for y in sorted_ys:
                line_words = sorted(y_lines[y], key=lambda x: x['x0'])
                line_txt = " ".join([w['text'] for w in line_words]).upper()
                
                # 'Quantity/Prices' সামারি টেবিল খোঁজা (KIABI এর মূল ডাটা এখানে থাকে)
                if ("COLO/SIZE" in line_txt or "COLOR/SIZE" in line_txt) and "TOTAL" in line_txt:
                    header_y = y
                    for w in line_words:
                        txt = w['text'].strip()
                        if txt.upper() not in ["COLO/SIZE", "COLOR/SIZE", "TOTAL", "PRICE", "AMOUNT", "CURRENCY"]:
                            size_cols.append({'name': txt, 'x0': w['x0'] - 10, 'x1': w['x1'] + 10})
                    break

            if header_y != -1:
                for y in sorted_ys:
                    if y <= header_y: continue
                    
                    line_words = sorted(y_lines[y], key=lambda x: x['x0'])
                    line_txt = " ".join([w['text'] for w in line_words])
                    
                    # টেবিলের শেষ বর্ডার
                    if "Total Quantity" in line_txt or "Total Amount" in line_txt: break
                    
                    # কালার নাম (হেডারের প্রথম কলামের বামে যা থাকে)
                    color_candidate = [w['text'] for w in line_words if w['x1'] < size_cols[0]['x0']]
                    color_name = " ".join(color_candidate).replace("Spec. price", "").strip()
                    
                    if color_name and not color_name.isdigit():
                        for col in size_cols:
                            # ওই কলামের X-সীমানায় কোনো সংখ্যা আছে কি না
                            cell_text = "".join([w['text'] for w in line_words if w['x0'] >= col['x0'] and w['x1'] <= col['x1']])
                            qty_match = re.search(r'(\d+)', cell_text.replace(",", ""))
                            qty = int(qty_match.group(1)) if qty_match else 0
                            
                            rows.append({
                                'P.O NO': order_no,
                                'Color': color_name,
                                'Size': col['name'],
                                'Quantity': qty
                            })

    return rows, meta

# ==========================================
#  FLASK ROUTES
# ==========================================

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        files = request.files.getlist('pdf_files')
        combined_rows = []
        final_meta = {'buyer': 'N/A', 'booking': 'N/A', 'style': 'N/A', 'season': 'N/A', 'dept': 'N/A', 'item': 'N/A'}
        
        for f in files:
            if not f.filename: continue
            path = os.path.join(app.config['UPLOAD_FOLDER'], f.filename)
            f.save(path)
            d, m = process_kiabi_pdf(path)
            if d: combined_rows.extend(d)
            for k, v in m.items():
                if v != 'N/A': final_meta[k] = v
        
        if not combined_rows:
            return render_template_string(INDEX_HTML, message="No PO data extracted. Check PDF format.")

        df = pd.DataFrame(combined_rows)
        # ডাটা এগ্রিগেশন
        df = df.groupby(['P.O NO', 'Color', 'Size'])['Quantity'].sum().reset_index()
        
        def size_sort(s):
            order = ['XXS','XS','S','M','L','XL','XXL','3XL','4XL','TU','ONE SIZE']
            return order.index(s.upper()) if s.upper() in order else 99

        final_tables = []
        grand_total = 0

        for color in df['Color'].unique():
            c_df = df[df['Color'] == color]
            pivot = c_df.pivot_table(index='P.O NO', columns='Size', values='Quantity', aggfunc='sum', fill_value=0)
            
            # কলাম সর্ট করা
            sorted_cols = sorted(pivot.columns.tolist(), key=size_sort)
            pivot = pivot[sorted_cols]
            
            pivot['Total'] = pivot.sum(axis=1)
            grand_total += pivot['Total'].sum()

            # সামারি রো
            act = pivot.sum(); act.name = 'Actual Qty'
            p3 = (act * 1.03).round().astype(int); p3.name = '3% Order Qty'
            
            pivot = pd.concat([pivot, act.to_frame().T, p3.to_frame().T])
            pivot = pivot.reset_index().rename(columns={'index': 'P.O NO'})
            
            html = pivot.to_html(classes='table table-bordered table-striped', index=False)
            html = html.replace('<td>Actual Qty</td>', '<td class="summary-row">Actual Qty</td>')
            html = html.replace('<td>3% Order Qty</td>', '<td class="summary-row">3% Order Qty</td>')
            html = html.replace('<tr>', '<tr class="data-row">')
            html = html.replace('<tr class="data-row"><td>Actual Qty', '<tr class="summary-row"><td>Actual Qty')
            html = html.replace('<tr class="data-row"><td>3% Order Qty', '<tr class="summary-row"><td>3% Order Qty')
            
            final_tables.append({'color': color, 'table': html})

        return render_template_string(RESULT_HTML, tables=final_tables, meta=final_meta, grand_total=f"{grand_total:,}")

    return render_template_string(INDEX_HTML)

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)

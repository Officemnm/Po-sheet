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
                <div class="card">
                    <div class="header text-center">
                        <h2>KIABI PO REPORT GENERATOR</h2>
                        <p class="mb-0">Precise Column Mapping Engine</p>
                    </div>
                    <div class="card-body p-4 text-center">
                        <form action="/" method="post" enctype="multipart/form-data">
                            <div class="upload-box mb-4">
                                <h3>📂</h3>
                                <h5>সিলেক্ট করুন বুকিং এবং পিও ফাইল</h5>
                                <input class="form-control mt-3" type="file" name="pdf_files" multiple accept=".pdf" required>
                            </div>
                            <button type="submit" class="btn btn-primary btn-lg w-100">রিপোর্ট জেনারেট করুন</button>
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
        body { padding: 30px; font-family: 'Segoe UI', sans-serif; }
        .report-header { text-align: center; border-bottom: 3px solid #000; padding-bottom: 10px; margin-bottom: 20px; }
        .meta-table { width: 100%; border: 1px solid #000; margin-bottom: 20px; }
        .meta-table td { padding: 8px; font-weight: bold; border: 1px solid #000; }
        .color-title { background: #1e3a8a; color: white; padding: 10px; font-weight: 900; margin-top: 30px; }
        .table th { background: #334155 !important; color: white !important; border: 1px solid #000 !important; text-align: center; font-weight: 900; }
        .table td { border: 1px solid #000 !important; text-align: center; font-weight: 800; font-size: 1.1rem; }
        .summary-row td { background: #d1ecff !important; border-top: 2px solid #000 !important; }
        .grand-total { background: #1e3a8a; color: white; padding: 15px; text-align: center; font-size: 2rem; font-weight: 900; }
        @media print { .no-print { display: none; } }
    </style>
</head>
<body>
    <div class="container-fluid">
        <div class="no-print text-end mb-3">
            <a href="/" class="btn btn-secondary">Upload New</a>
            <button onclick="window.print()" class="btn btn-primary">Print Report</button>
        </div>

        <div class="report-header">
            <h1 class="display-5 fw-bold">Cotton Clothing BD Limited</h1>
            <h4>Purchase Order Summary</h4>
        </div>

        <table class="meta-table">
            <tr>
                <td>Buyer: {{ meta.buyer }}</td>
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

        <div class="grand-total mb-4">GRAND TOTAL: {{ grand_total }} Pcs</div>

        {% for item in tables %}
            <div class="color-title">COLOR: {{ item.color }}</div>
            <div class="table-responsive">
                {{ item.table | safe }}
            </div>
        {% endfor %}
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
        first_page_text = pdf.pages[0].extract_text() or ""
        
        # Meta extraction
        if "KIABI" in first_page_text.upper(): meta['buyer'] = "KIABI"
        m_booking = re.search(r"Booking NO\.?[:\s]*([\w-]+)", first_page_text, re.I)
        if m_booking: meta['booking'] = m_booking.group(1)
        m_style = re.search(r"Style (?:Ref|Des)\.?[:\s]*([\w-]+)", first_page_text, re.I)
        if m_style: meta['style'] = m_style.group(1)
        m_season = re.search(r"Season\s*[:\s]*([\w\d-]+)", first_page_text, re.I)
        if m_season: meta['season'] = m_season.group(1)
        
        # Order No
        m_order = re.search(r"Order no: (\d+)", first_page_text, re.I)
        if m_order: order_no = m_order.group(1)
        
        if "Main Fabric Booking" in first_page_text:
            return [], meta

        for page in pdf.pages:
            words = page.extract_words()
            if not words: continue

            # হেডার ডিটেকশন (X-Coordinates map করার জন্য)
            size_map = [] # List of {name: str, x0: float, x1: float}
            
            # রো অনুযায়ী ওয়ার্ডগুলো সাজানো
            y_groups = {}
            for w in words:
                y = round(w['top'], 0)
                if y not in y_groups: y_groups[y] = []
                y_groups[y].append(w)
            
            sorted_ys = sorted(y_groups.keys())
            
            header_y = -1
            for y in sorted_ys:
                line_words = sorted(y_groups[y], key=lambda x: x['x0'])
                line_txt = " ".join([w['text'] for w in line_words]).upper()
                
                # যদি হেডার লাইন পাওয়া যায়
                if "COLO/SIZE" in line_txt and "TOTAL" in line_txt:
                    header_y = y
                    # সাইজগুলোর X-অবস্থান সেভ করে রাখা
                    for w in line_words:
                        txt = w['text'].strip()
                        if txt.upper() not in ["COLO/SIZE", "TOTAL", "PRICE", "AMOUNT", "CURRENCY"]:
                            # বাফার হিসেবে ৫ পিক্সেল ডানে বামে নেওয়া
                            size_map.append({'name': txt, 'x0': w['x0'] - 8, 'x1': w['x1'] + 8})
                    break

            if header_y != -1:
                # হেডারের নিচের লাইনগুলো চেক করা
                for y in sorted_ys:
                    if y <= header_y: continue
                    
                    line_words = sorted(y_groups[y], key=lambda x: x['x0'])
                    line_txt = " ".join([w['text'] for w in line_words])
                    
                    if "Total Quantity" in line_txt or "Total Amount" in line_txt: break
                    
                    # কালার কলাম (সাধারণত সবার বামে থাকে)
                    color_words = [w['text'] for w in line_words if w['x1'] < size_map[0]['x0']]
                    color_name = " ".join(color_words).replace("Spec. price", "").strip()
                    
                    if color_name and not color_name.isdigit():
                        for col in size_map:
                            # ওই সাইজ কলামের X-সীমানার ভেতরে কোনো সংখ্যা আছে কি না
                            cell_data = [w['text'] for w in line_words if w['x0'] >= col['x0'] and w['x1'] <= col['x1']]
                            val_str = "".join(cell_data).replace(",", "")
                            
                            qty = 0
                            match = re.search(r'(\d+)', val_str)
                            if match: qty = int(match.group(1))
                            
                            rows.append({
                                'P.O NO': order_no,
                                'Color': color_name,
                                'Size': col['name'],
                                'Quantity': qty
                            })
                            
            # বিশেষ লজিক: Assortment tables (১০৯১৩৭০০ এর মতো ফাইলের জন্য)
            if "Quantity" in first_page_text and "Size" in first_page_text and not rows:
                # এখানেও একই X-Coordinate লজিক দিয়ে অ্যাসার্টমেন্ট টেবিল পড়া সম্ভব
                pass 

    return rows, meta

# ==========================================
#  FLASK ROUTES
# ==========================================

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        files = request.files.getlist('pdf_files')
        all_data = []
        f_meta = {'buyer': 'N/A', 'booking': 'N/A', 'style': 'N/A', 'season': 'N/A', 'dept': 'N/A', 'item': 'N/A'}
        
        for f in files:
            if not f.filename: continue
            path = os.path.join(app.config['UPLOAD_FOLDER'], f.filename)
            f.save(path)
            d, m = process_kiabi_pdf(path)
            if d: all_data.extend(d)
            for k, v in m.items():
                if v != 'N/A': f_meta[k] = v
        
        if not all_data:
            return render_template_string(INDEX_HTML, message="No data found in these PDFs.")

        df = pd.DataFrame(all_data)
        # একই কালার-সাইজ-পিও এর ডুপ্লিকেট যোগ করা
        df = df.groupby(['P.O NO', 'Color', 'Size'])['Quantity'].sum().reset_index()
        
        final_tables = []
        grand_total_pcs = 0

        # সাইজ সর্টিং এর জন্য স্ট্যান্ডার্ড অর্ডার
        def sort_size(s):
            std = ['XXS','XS','S','M','L','XL','XXL','3XL','4XL','5XL','TU','ONE SIZE']
            return std.index(s.upper()) if s.upper() in std else 99

        for color in df['Color'].unique():
            c_df = df[df['Color'] == color]
            pivot = c_df.pivot_table(index='P.O NO', columns='Size', values='Quantity', aggfunc='sum', fill_value=0)
            
            # কলাম সর্ট
            cols = sorted(pivot.columns.tolist(), key=sort_size)
            pivot = pivot[cols]
            
            pivot['Total'] = pivot.sum(axis=1)
            grand_total_pcs += pivot['Total'].sum()

            act = pivot.sum(); act.name = 'Actual Qty'
            p3 = (act * 1.03).round().astype(int); p3.name = '3% Order Qty'
            
            pivot = pd.concat([pivot, act.to_frame().T, p3.to_frame().T])
            pivot = pivot.reset_index().rename(columns={'index': 'P.O NO'})
            
            html = pivot.to_html(classes='table table-bordered', index=False)
            html = html.replace('<td>Actual Qty</td>', '<td class="summary-row">Actual Qty</td>')
            html = html.replace('<td>3% Order Qty</td>', '<td class="summary-row">3% Order Qty</td>')
            html = html.replace('<tr>', '<tr class="data-row">')
            html = html.replace('<tr class="data-row"><td>Actual Qty', '<tr class="summary-row"><td>Actual Qty')
            html = html.replace('<tr class="data-row"><td>3% Order Qty', '<tr class="summary-row"><td>3% Order Qty')
            
            final_tables.append({'color': color, 'table': html})

        return render_template_string(RESULT_HTML, tables=final_tables, meta=f_meta, grand_total=f"{grand_total_pcs:,}")

    return render_template_string(INDEX_HTML)

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)

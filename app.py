from flask import Flask, request, render_template_string
import pypdf
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
#  HTML & CSS TEMPLATES
# ==========================================

INDEX_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Purchase Order Parser</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
        * { font-family: 'Poppins', sans-serif; }
        body { 
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .main-card { 
            border: none; 
            border-radius: 24px; 
            box-shadow: 0 25px 50px rgba(0,0,0,0.3);
            backdrop-filter: blur(10px);
            background: rgba(255,255,255,0.95);
            overflow: hidden;
        }
        .card-header { 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 50%, #f093fb 100%);
            color: white; 
            padding: 40px; 
            text-align: center;
            position: relative;
        }
        .card-header::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: url("data:image/svg+xml,%3Csvg width='60' height='60' viewBox='0 0 60 60' xmlns='http://www.w3.org/2000/svg'%3E%3Cg fill='none' fill-rule='evenodd'%3E%3Cg fill='%23ffffff' fill-opacity='0.08'%3E%3Cpath d='M36 34v-4h-2v4h-4v2h4v4h2v-4h4v-2h-4zm0-30V0h-2v4h-4v2h4v4h2V6h4V4h-4zM6 34v-4H4v4H0v2h4v4h2v-4h4v-2H6zM6 4V0H4v4H0v2h4v4h2V6h4V4H6z'/%3E%3C/g%3E%3C/g%3E%3C/svg%3E");
        }
        .card-header h2 { 
            font-weight: 700; 
            font-size: 2rem;
            margin-bottom: 8px;
            position: relative;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
        }
        .card-header p {
            position: relative;
            opacity: 0.9;
            font-weight: 400;
        }
        .btn-upload { 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border: none; 
            padding: 16px 40px; 
            font-weight: 600; 
            font-size: 1.1rem;
            border-radius: 50px;
            transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
            box-shadow: 0 10px 30px rgba(102, 126, 234, 0.4);
        }
        .btn-upload:hover { 
            transform: translateY(-5px) scale(1.02);
            box-shadow: 0 20px 40px rgba(102, 126, 234, 0.5);
        }
        .upload-icon { 
            font-size: 4rem; 
            margin-bottom: 20px;
            animation: float 3s ease-in-out infinite;
        }
        @keyframes float {
            0%, 100% { transform: translateY(0); }
            50% { transform: translateY(-10px); }
        }
        .file-input-wrapper { 
            border: 3px dashed #cbd5e0; 
            border-radius: 20px; 
            padding: 50px; 
            background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
            transition: all 0.3s ease;
        }
        .file-input-wrapper:hover { 
            border-color: #764ba2; 
            background: linear-gradient(135deg, #faf5ff 0%, #f3e8ff 100%);
            transform: scale(1.01);
        }
        .file-input-wrapper h5 {
            font-weight: 600;
            color: #1e293b;
        }
        .footer-credit { 
            margin-top: 30px; 
            font-size: 0.85rem; 
            color: #64748b;
        }
        .footer-credit strong {
            color: #764ba2;
        }
    </style>
</head>
<body>
    <div class="container py-5">
        <div class="row justify-content-center">
            <div class="col-md-8 col-lg-6">
                <div class="card main-card">
                    <div class="card-header">
                        <h2>📊 PDF Report Generator</h2>
                        <p class="mb-0">Cotton Clothing BD Limited</p>
                    </div>
                    <div class="card-body p-5 text-center">
                        <form action="/" method="post" enctype="multipart/form-data">
                            <div class="file-input-wrapper mb-4">
                                <div class="upload-icon">📂</div>
                                <h5>Select PDF Files</h5>
                                <p class="text-muted small mb-3">Select both Booking File & PO Files together</p>
                                <input class="form-control form-control-lg" type="file" name="pdf_files" multiple accept=".pdf" required 
                                    style="border-radius: 12px; border: 2px solid #e2e8f0; padding: 12px;">
                            </div>
                            <button type="submit" class="btn btn-primary btn-upload btn-lg w-100">
                                🚀 Generate Report
                            </button>
                        </form>
                        <div class="footer-credit">
                            Developed by <strong>Mehedi Hasan</strong>
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
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary-gradient: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            --dark-bg: #0f172a;
            --card-bg: #ffffff;
            --text-primary: #1e293b;
            --text-secondary: #64748b;
            --border-color: #e2e8f0;
            --success-color: #10b981;
            --accent-color: #8b5cf6;
        }
        
        * { font-family: 'Inter', sans-serif; }
        
        body { 
            background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%);
            min-height: 100vh;
            padding: 40px 0; 
        }
        
        .container { max-width: 1200px; }
        
        /* ===== HEADER SECTION ===== */
        .company-header { 
            background: var(--card-bg);
            border-radius: 20px;
            padding: 30px 40px;
            margin-bottom: 25px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.08);
            border: 1px solid var(--border-color);
            position: relative;
            overflow: hidden;
        }
        
        .company-header::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 5px;
            background: var(--primary-gradient);
        }
        
        .header-content {
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .company-info {
            text-align: left;
        }
        
        .company-name { 
            font-size: 1.8rem; 
            font-weight: 800; 
            color: var(--text-primary);
            text-transform: uppercase; 
            letter-spacing: 0.5px;
            margin-bottom: 5px;
            background: var(--primary-gradient);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }
        
        .report-title { 
            font-size: 1rem; 
            color: var(--text-secondary);
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 2px;
        }
        
        .date-badge {
            background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%);
            border: 2px solid #86efac;
            border-radius: 12px;
            padding: 12px 24px;
            text-align: center;
        }
        
        .date-label {
            font-size: 0.75rem;
            color: var(--text-secondary);
            text-transform: uppercase;
            letter-spacing: 1px;
            font-weight: 600;
        }
        
        .date-value {
            font-size: 1.2rem;
            font-weight: 700;
            color: #166534;
        }
        
        /* ===== INFO SECTION ===== */
        .info-container { 
            display: grid;
            grid-template-columns: 1fr auto;
            gap: 20px;
            margin-bottom: 25px;
        }
        
        .info-box { 
            background: var(--card-bg);
            border-radius: 20px;
            padding: 25px 30px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.08);
            border: 1px solid var(--border-color);
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px 40px;
        }
        
        .info-item { 
            display: flex;
            align-items: center;
            gap: 12px;
            padding: 8px 0;
            border-bottom: 1px dashed var(--border-color);
        }
        
        .info-item:last-child {
            border-bottom: none;
        }
        
        .info-icon {
            width: 40px;
            height: 40px;
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.2rem;
            flex-shrink: 0;
        }
        
        .info-icon.buyer { background: linear-gradient(135deg, #dbeafe 0%, #bfdbfe 100%); }
        .info-icon.booking { background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%); }
        .info-icon.style { background: linear-gradient(135deg, #fce7f3 0%, #fbcfe8 100%); }
        .info-icon.season { background: linear-gradient(135deg, #d1fae5 0%, #a7f3d0 100%); }
        .info-icon.dept { background: linear-gradient(135deg, #e0e7ff 0%, #c7d2fe 100%); }
        .info-icon.item { background: linear-gradient(135deg, #fed7d7 0%, #fecaca 100%); }
        
        .info-text {
            flex: 1;
            min-width: 0;
        }
        
        .info-label { 
            font-size: 0.7rem;
            font-weight: 600;
            color: var(--text-secondary);
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 2px;
        }
        
        .info-value { 
            font-size: 1rem;
            font-weight: 700;
            color: var(--text-primary);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .total-box { 
            background: var(--primary-gradient);
            color: white; 
            padding: 25px 35px;
            border-radius: 20px;
            text-align: center;
            display: flex;
            flex-direction: column;
            justify-content: center;
            min-width: 200px;
            box-shadow: 0 10px 40px rgba(102, 126, 234, 0.4);
            position: relative;
            overflow: hidden;
        }
        
        .total-box::before {
            content: '';
            position: absolute;
            top: -50%;
            right: -50%;
            width: 100%;
            height: 100%;
            background: radial-gradient(circle, rgba(255,255,255,0.2) 0%, transparent 70%);
        }
        
        .total-label { 
            font-size: 0.8rem;
            text-transform: uppercase;
            letter-spacing: 2px;
            font-weight: 600;
            opacity: 0.9;
            margin-bottom: 5px;
            position: relative;
        }
        
        .total-value { 
            font-size: 2.8rem;
            font-weight: 800;
            line-height: 1;
            position: relative;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
        }
        
        .total-unit {
            font-size: 0.75rem;
            opacity: 0.8;
            margin-top: 5px;
            position: relative;
        }

        /* ===== TABLE SECTION ===== */
        .table-card { 
            background: var(--card-bg);
            border-radius: 20px;
            margin-bottom: 25px;
            overflow: hidden;
            box-shadow: 0 4px 20px rgba(0,0,0,0.08);
            border: 1px solid var(--border-color);
        }
        
        .color-header { 
            background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
            color: white;
            padding: 18px 25px;
            font-size: 1.1rem;
            font-weight: 700;
            display: flex;
            align-items: center;
            gap: 12px;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        
        .color-badge {
            width: 24px;
            height: 24px;
            border-radius: 6px;
            background: linear-gradient(135deg, var(--accent-color) 0%, #a78bfa 100%);
            box-shadow: 0 2px 8px rgba(139, 92, 246, 0.4);
        }

        .table { 
            margin-bottom: 0; 
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
        }
        
        .table th { 
            background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%);
            color: var(--text-primary);
            font-weight: 700;
            font-size: 0.85rem;
            text-align: center;
            padding: 14px 12px;
            border-bottom: 2px solid var(--border-color);
            text-transform: uppercase;
            letter-spacing: 0.5px;
            position: sticky;
            top: 0;
        }
        
        .table td { 
            text-align: center;
            vertical-align: middle;
            padding: 14px 12px;
            color: var(--text-primary);
            font-weight: 600;
            font-size: 1rem;
            border-bottom: 1px solid var(--border-color);
            transition: background 0.2s ease;
        }
        
        .table tbody tr:hover td {
            background: #f8fafc;
        }
        
        .table tbody tr:last-child td {
            border-bottom: none;
        }
        
        .order-col { 
            font-weight: 800 !important;
            background: linear-gradient(135deg, #faf5ff 0%, #f3e8ff 100%) !important;
            color: #7c3aed !important;
            white-space: nowrap;
            border-right: 2px solid #e9d5ff !important;
        }
        
        .total-col { 
            font-weight: 800 !important;
            background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%) !important;
            color: #166534 !important;
            border-left: 2px solid #86efac !important;
        }
        
        .total-col-header { 
            background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%) !important;
            color: #166534 !important;
            font-weight: 800 !important;
            border-left: 2px solid #86efac !important;
        }

        /* Summary Rows */
        .table tbody tr.summary-row td { 
            background: linear-gradient(135deg, #dbeafe 0%, #bfdbfe 100%) !important;
            color: #1e40af !important;
            font-weight: 800 !important;
            font-size: 1.05rem !important;
            border-top: 3px solid #3b82f6 !important;
            padding: 16px 12px !important;
        }
        
        .summary-label { 
            text-align: right !important;
            padding-right: 20px !important;
            text-transform: uppercase;
            letter-spacing: 1px;
            font-size: 0.85rem !important;
        }

        /* ===== ACTION BAR ===== */
        .action-bar { 
            margin-bottom: 25px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .btn-back {
            background: var(--card-bg);
            color: var(--text-primary);
            border: 2px solid var(--border-color);
            border-radius: 12px;
            padding: 12px 28px;
            font-weight: 600;
            transition: all 0.3s ease;
            text-decoration: none;
            display: inline-flex;
            align-items: center;
            gap: 8px;
        }
        
        .btn-back:hover {
            border-color: var(--accent-color);
            color: var(--accent-color);
            transform: translateX(-5px);
        }
        
        .btn-print { 
            background: var(--primary-gradient);
            color: white;
            border: none;
            border-radius: 12px;
            padding: 14px 32px;
            font-weight: 600;
            font-size: 1rem;
            transition: all 0.3s ease;
            display: inline-flex;
            align-items: center;
            gap: 10px;
            box-shadow: 0 8px 25px rgba(102, 126, 234, 0.35);
        }
        
        .btn-print:hover {
            transform: translateY(-3px);
            box-shadow: 0 12px 35px rgba(102, 126, 234, 0.45);
        }
        
        /* ===== FOOTER ===== */
        .footer-credit { 
            text-align: center;
            margin-top: 40px;
            padding: 20px;
            background: var(--card-bg);
            border-radius: 16px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.05);
            border: 1px solid var(--border-color);
        }
        
        .footer-text {
            font-size: 0.9rem;
            color: var(--text-secondary);
        }
        
        .footer-text strong {
            color: var(--accent-color);
            font-weight: 700;
        }

        /* ===== PRINT STYLES ===== */
        @media print {
            @page { 
                margin: 8mm;
                size: portrait;
            }
            
            body { 
                background: white !important;
                padding: 0 !important;
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
                color-adjust: exact !important;
            }
            
            .container { 
                max-width: 100% !important;
                padding: 0 !important;
                margin: 0 !important;
            }
            
            .no-print { display: none !important; }
            
            .company-header {
                border-radius: 0 !important;
                box-shadow: none !important;
                border: 2px solid #000 !important;
                margin-bottom: 15px !important;
                padding: 15px 20px !important;
            }
            
            .company-header::before {
                height: 3px !important;
                background: #000 !important;
            }
            
            .company-name {
                font-size: 1.5rem !important;
                -webkit-text-fill-color: #000 !important;
                color: #000 !important;
            }
            
            .info-container {
                margin-bottom: 15px !important;
            }
            
            .info-box {
                border-radius: 0 !important;
                box-shadow: none !important;
                border: 1px solid #000 !important;
                padding: 15px !important;
            }
            
            .info-icon { display: none !important; }
            
            .info-item {
                padding: 4px 0 !important;
            }
            
            .info-label {
                font-size: 9pt !important;
            }
            
            .info-value {
                font-size: 11pt !important;
            }
            
            .total-box {
                border-radius: 0 !important;
                box-shadow: none !important;
                border: 2px solid #000 !important;
                background: #f5f5f5 !important;
                color: #000 !important;
            }
            
            .total-box * {
                color: #000 !important;
            }
            
            .table-card {
                border-radius: 0 !important;
                box-shadow: none !important;
                border: 1px solid #000 !important;
                margin-bottom: 15px !important;
                break-inside: avoid;
            }
            
            .color-header {
                background: #e0e0e0 !important;
                color: #000 !important;
                padding: 10px 15px !important;
                font-size: 12pt !important;
                -webkit-print-color-adjust: exact !important;
            }
            
            .color-badge { display: none !important; }
            
            .table th {
                background: #f0f0f0 !important;
                color: #000 !important;
                font-size: 10pt !important;
                padding: 8px 6px !important;
                border: 1px solid #000 !important;
                -webkit-print-color-adjust: exact !important;
            }
            
            .table td {
                font-size: 11pt !important;
                padding: 8px 6px !important;
                border: 1px solid #000 !important;
            }
            
            .order-col {
                background: #f8f8f8 !important;
                color: #000 !important;
                -webkit-print-color-adjust: exact !important;
            }
            
            .total-col,
            .total-col-header {
                background: #e8f5e9 !important;
                color: #000 !important;
                -webkit-print-color-adjust: exact !important;
            }
            
            .summary-row td {
                background: #e3f2fd !important;
                color: #000 !important;
                font-size: 11pt !important;
                border-top: 2px solid #000 !important;
                -webkit-print-color-adjust: exact !important;
            }
            
            .footer-credit {
                border-radius: 0 !important;
                box-shadow: none !important;
                border-top: 1px solid #000 !important;
                margin-top: 20px !important;
                padding: 10px !important;
                background: transparent !important;
            }
            
            .footer-text {
                font-size: 9pt !important;
                color: #000 !important;
            }
            
            .footer-text strong {
                color: #000 !important;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- Action Bar -->
        <div class="action-bar no-print">
            <a href="/" class="btn-back">
                ← Back to Upload
            </a>
            <button onclick="window.print()" class="btn-print">
                🖨️ Print Report
            </button>
        </div>

        <!-- Company Header -->
        <div class="company-header">
            <div class="header-content">
                <div class="company-info">
                    <div class="company-name">Cotton Clothing BD Limited</div>
                    <div class="report-title">Purchase Order Summary Report</div>
                </div>
                <div class="date-badge">
                    <div class="date-label">Report Date</div>
                    <div class="date-value" id="date"></div>
                </div>
            </div>
        </div>

        {% if message %}
            <div class="alert alert-warning text-center no-print" style="border-radius: 16px; border: none; box-shadow: 0 4px 15px rgba(0,0,0,0.1);">
                {{ message }}
            </div>
        {% endif %}

        {% if tables %}
            <!-- Info Section -->
            <div class="info-container">
                <div class="info-box">
                    <div class="info-item">
                        <div class="info-icon buyer">👤</div>
                        <div class="info-text">
                            <div class="info-label">Buyer</div>
                            <div class="info-value">{{ meta.buyer }}</div>
                        </div>
                    </div>
                    <div class="info-item">
                        <div class="info-icon season">🌤️</div>
                        <div class="info-text">
                            <div class="info-label">Season</div>
                            <div class="info-value">{{ meta.season }}</div>
                        </div>
                    </div>
                    <div class="info-item">
                        <div class="info-icon booking">📋</div>
                        <div class="info-text">
                            <div class="info-label">Booking</div>
                            <div class="info-value">{{ meta.booking }}</div>
                        </div>
                    </div>
                    <div class="info-item">
                        <div class="info-icon dept">🏢</div>
                        <div class="info-text">
                            <div class="info-label">Department</div>
                            <div class="info-value">{{ meta.dept }}</div>
                        </div>
                    </div>
                    <div class="info-item">
                        <div class="info-icon style">✨</div>
                        <div class="info-text">
                            <div class="info-label">Style</div>
                            <div class="info-value">{{ meta.style }}</div>
                        </div>
                    </div>
                    <div class="info-item">
                        <div class="info-icon item">👕</div>
                        <div class="info-text">
                            <div class="info-label">Item</div>
                            <div class="info-value">{{ meta.item }}</div>
                        </div>
                    </div>
                </div>
                
                <div class="total-box">
                    <div class="total-label">Grand Total</div>
                    <div class="total-value">{{ grand_total }}</div>
                    <div class="total-unit">Pieces</div>
                </div>
            </div>

            <!-- Color Tables -->
            {% for item in tables %}
                <div class="table-card">
                    <div class="color-header">
                        <div class="color-badge"></div>
                        Color: {{ item.color }}
                    </div>
                    <div class="table-responsive">
                        {{ item.table | safe }}
                    </div>
                </div>
            {% endfor %}
            
            <!-- Footer -->
            <div class="footer-credit">
                <div class="footer-text">
                    Report Generated by <strong>Mehedi Hasan</strong> | Cotton Clothing BD Limited
                </div>
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
#  LOGIC PART (হুবহু আগের মতো)
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


def parse_vertical_table(lines, start_idx, sizes, order_no):
    """
    Vertical format table parse করে।
    """
    extracted_data = []
    i = start_idx
    
    while i < len(lines):
        line = lines[i].strip()
        
        if line.startswith("Total") and i + 1 < len(lines):
            next_line = lines[i + 1].strip() if i + 1 < len(lines) else ""
            if "Quantity" in next_line or "Amount" in next_line or re.match(r'^Quantity', next_line):
                break
            if re.match(r'^\d', next_line):
                break
        
        if line and re.search(r'[a-zA-Z]', line):
            if any(kw in line.lower() for kw in ['spec', 'price', 'total', 'quantity', 'amount']):
                i += 1
                continue
            
            color_name = line
            i += 1
            
            if i < len(lines) and 'spec' in lines[i].lower():
                i += 1
            
            quantities = []
            size_idx = 0
            
            while size_idx < len(sizes) and i < len(lines):
                qty_line = lines[i].strip() if i < len(lines) else ""
                price_line = lines[i + 1].strip() if i + 1 < len(lines) else ""
                
                if qty_line and re.search(r'[a-zA-Z]', qty_line):
                    if not any(kw in qty_line.lower() for kw in ['spec', 'price']):
                        while size_idx < len(sizes):
                            quantities.append(0)
                            size_idx += 1
                        break
                
                if (qty_line == "" or qty_line.isspace()) and (price_line == "" or price_line.isspace()):
                    quantities.append(0)
                    size_idx += 1
                    i += 2
                    continue
                
                if re.match(r'^\d+$', qty_line):
                    quantities.append(int(qty_line))
                    size_idx += 1
                    i += 2
                    continue
                
                i += 1
            
            if quantities:
                while len(quantities) < len(sizes):
                    quantities.append(0)
                
                for idx, size in enumerate(sizes):
                    extracted_data.append({
                        'P.O NO': order_no,
                        'Color': color_name,
                        'Size': size,
                        'Quantity': quantities[idx] if idx < len(quantities) else 0
                    })
            
            continue
        
        i += 1
    
    return extracted_data


def extract_data_dynamic(file_path):
    extracted_data = []
    metadata = {
        'buyer': 'N/A', 'booking': 'N/A', 'style': 'N/A', 
        'season': 'N/A', 'dept': 'N/A', 'item': 'N/A'
    }
    order_no = "Unknown"
    
    try:
        reader = pypdf.PdfReader(file_path)
        first_page_text = reader.pages[0].extract_text()
        
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

        for page in reader.pages:
            text = page.extract_text()
            lines = text.split('\n')
            
            for i, line in enumerate(lines):
                if ("Colo" in line or "Size" in line) and "Total" in line:
                    parts = line.split()
                    try:
                        total_idx = [idx for idx, x in enumerate(parts) if 'Total' in x][0]
                        raw_sizes = parts[:total_idx]
                        sizes = [s for s in raw_sizes if s not in ["Colo", "/", "Size", "Colo/Size", "Colo/", "Size's"]]
                        
                        valid_size_count = sum(1 for s in sizes if is_potential_size(s))
                        if sizes and valid_size_count >= len(sizes) / 2:
                            data = parse_vertical_table(lines, i + 1, sizes, order_no)
                            extracted_data.extend(data)
                    except: 
                        pass
                    break
                    
    except Exception as e: 
        print(f"Error processing file: {e}")
    
    return extracted_data, metadata


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
            
            table_html = re.sub(r'<tr>\s*<td>', '<tr><td class="order-col">', table_html)
            table_html = table_html.replace('<th>Total</th>', '<th class="total-col-header">Total</th>')
            table_html = table_html.replace('<td>Total</td>', '<td class="total-col">Total</td>')
            
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

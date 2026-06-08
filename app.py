import io
import re
import csv
import json
import tempfile
import os
from datetime import date
from flask import Flask, request, jsonify, send_file, render_template
import pdfplumber
import openpyxl

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

FIXED = {
    'vendor': 'SHNSSD03',
    'warehouse': 'S01',
    'shipment_type': 'B2B',
    'ai': '　',  # full-width space
}

HEADERS = ['日期', 'STO #', 'Line', 'QSS 料號', '數量', '廠商', '預計入庫倉', '出貨指示類型', 'AI#', '單價']


def parse_pl(pdf_bytes):
    """Extract items from Packing List PDF."""
    items = []
    sto = None
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split('\n'):
                # Item line pattern: {item_no} {material} {qty}EA {sto} {item_no} ...
                m = re.match(r'^(\d+)\s+([\w\-]+)\s+([\d,]+)EA\s+(\d+)\s+\d+', line)
                if m:
                    item_no = int(m.group(1))
                    material = m.group(2)
                    qty = int(m.group(3).replace(',', ''))
                    sto = m.group(4)
                    items.append({'line': item_no, 'material': material, 'qty': qty})
    return items, sto


def parse_ei(pdf_bytes):
    """Extract unit prices from Export Invoice PDF. Returns {material: unit_price}."""
    prices = {}
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            lines = text.split('\n')
            current_material = None
            for line in lines:
                # Item header: {seq} {material} EA {qty}
                m = re.match(r'^\d+\s+([\w\-]+)\s+EA\s+([\d,]+)', line)
                if m:
                    current_material = m.group(1)
                    continue
                # BOM row: {bom} {CN} {qty} {asm_rate} / [EA] {ext_asm} {unit_price} ...
                # "EA" is optional — some EI formats omit it after the slash
                if current_material:
                    m2 = re.match(
                        r'^\d+\s+[A-Z]{2}\s+[\d,]+\s+[\d.]+\s+/\s+(?:EA\s+)?[\d,.]+\s+([\d.]+)',
                        line
                    )
                    if m2:
                        prices[current_material] = float(m2.group(1))
                        current_material = None
    return prices


def build_rows(items, sto, prices):
    today = date.today().strftime('%-m/%-d/%Y') if os.name != 'nt' else \
        f"{date.today().month}/{date.today().day}/{date.today().year}"
    # Format: yyyy/m/d
    d = date.today()
    today_str = f"{d.year}/{d.month}/{d.day}"
    rows = []
    for item in items:
        mat = item['material']
        unit_price = prices.get(mat, '')
        rows.append([
            today_str,
            sto or '',
            item['line'],
            '7SD' + mat,
            item['qty'],
            FIXED['vendor'],
            FIXED['warehouse'],
            FIXED['shipment_type'],
            FIXED['ai'],
            unit_price,
        ])
    return rows


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/extract', methods=['POST'])
def extract():
    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': 'No files uploaded'}), 400

    pl_bytes = None
    ei_bytes = None

    for f in files:
        name_upper = f.filename.upper()
        data = f.read()
        if 'PL_' in name_upper or name_upper.startswith('PL'):
            pl_bytes = data
        elif 'EI_' in name_upper or name_upper.startswith('EI'):
            ei_bytes = data
        # Fallback: auto-detect by content
        if pl_bytes is None or ei_bytes is None:
            try:
                with pdfplumber.open(io.BytesIO(data)) as pdf:
                    first_text = (pdf.pages[0].extract_text() or '') if pdf.pages else ''
                    if 'Packing List' in first_text and pl_bytes is None:
                        pl_bytes = data
                    elif 'Unit Price' in first_text and ei_bytes is None:
                        ei_bytes = data
            except Exception:
                pass

    if not pl_bytes:
        return jsonify({'error': 'Packing List (PL) PDF not found'}), 400
    if not ei_bytes:
        return jsonify({'error': 'Export Invoice (EI) PDF not found'}), 400

    try:
        items, sto = parse_pl(pl_bytes)
        prices = parse_ei(ei_bytes)
        rows = build_rows(items, sto, prices)
    except Exception as e:
        return jsonify({'error': f'Parse error: {str(e)}'}), 500

    return jsonify({
        'headers': HEADERS,
        'rows': rows,
        'sto': sto,
        'item_count': len(rows),
    })


@app.route('/download/csv', methods=['POST'])
def download_csv():
    data = request.json
    headers = data.get('headers', HEADERS)
    rows = data.get('rows', [])
    filename = data.get('filename', 'INPL_extract.csv')

    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(headers)
    writer.writerows(rows)
    buf.seek(0)
    # Encode as UTF-8 BOM for Excel compatibility
    csv_bytes = io.BytesIO(('﻿' + buf.getvalue()).encode('utf-8'))
    return send_file(
        csv_bytes,
        mimetype='text/csv; charset=utf-8',
        as_attachment=True,
        download_name=filename,
    )


@app.route('/download/xlsx', methods=['POST'])
def download_xlsx():
    data = request.json
    headers = data.get('headers', HEADERS)
    rows = data.get('rows', [])
    filename = data.get('filename', 'INPL_extract.xlsx')

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'INPL'
    ws.append(headers)
    for row in rows:
        ws.append(row)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename,
    )


if __name__ == '__main__':
    app.run(debug=True, port=5000)

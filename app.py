from flask import Flask, render_template, request, send_file, jsonify
import os
import json
import pandas as pd
import tempfile
import shutil
from werkzeug.utils import secure_filename
import pdfkit
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'xlsx', 'xls'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

WKHTML = shutil.which("wkhtmltopdf")

if not WKHTML:
    WKHTML = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"

config = pdfkit.configuration(wkhtmltopdf=WKHTML)

print("Using wkhtmltopdf:", WKHTML)

def safe_filename(filename):
    return "".join(c for c in filename if c.isalnum() or c in (' ', '-', '_')).rstrip()

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/preview')
def preview():
    static_base = os.path.abspath("static")

    # Create sample data for preview
    items = []
    for i in range(1,28):
        items.append({
            "qty": '',
            "remark": ''
        })

    return render_template(
        "pdf_template.html",
        po='',
        store='',
        site_type='',
        address='',
        start_date='',
        end_date='',
        expiry_date='',
        items=items,
        images=[],
        serial_pages=[],
        static_base=static_base
    )

@app.route('/generate', methods=['POST'])
def generate():

    po = request.form.get('po') or ''
    store = request.form.get('store') or ''
    site_type = request.form.get('site_type') or ''
    address = request.form.get('address') or ''
    start_date = datetime.strptime(request.form.get('start_date'), '%Y-%m-%d').strftime('%d/%m/%Y') if request.form.get('start_date') else ''
    end_date = datetime.strptime(request.form.get('end_date'), '%Y-%m-%d').strftime('%d/%m/%Y') if request.form.get('end_date') else ''
    expiry_date = request.form.get('expiry_date') or ''

    items = []
    for i in range(1,28):
        items.append({
            "qty": request.form.get(f'qty{i}') or '',
            "remark": request.form.get(f'remark{i}') or ''
        })

    # ===== IMAGES =====
    image_urls = []
    photos = request.files.getlist("photos")

    for photo in photos:
        if photo.filename:
            filename = secure_filename(photo.filename)

            temp_dir = tempfile.mkdtemp()

            save_path = os.path.join(temp_dir, filename)
            photo.save(save_path)

            image_urls.append(
                "file:///" + os.path.abspath(save_path).replace("\\","/")
            )

    # ===== SERIAL EXCEL =====
    serial_pages = []
    excel = request.files.get("excel")

    if excel and excel.filename:
        temp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        excel.save(temp_excel.name)

        path = temp_excel.name

        df = pd.read_excel(path).fillna("")
        rows = df.to_dict("records")

        ROWS_PER_PAGE = 80
        HALF = 40

        for start in range(0, len(rows), ROWS_PER_PAGE):

            chunk = rows[start:start + ROWS_PER_PAGE]

            left = chunk[:HALF]
            right = chunk[HALF:HALF*2]

            serial_pages.append({
                "left": left,
                "right": right,
                "offset": start
            })

    static_base = os.path.abspath("static")

    html = render_template(
        "pdf_template.html",
        po=po,
        store=store,
        site_type=site_type,
        address=address,
        start_date=start_date,
        end_date=end_date,
        expiry_date=expiry_date,
        items=items,
        images=image_urls,
        serial_pages=serial_pages,
        static_base=static_base
    )

    output = os.path.join(tempfile.gettempdir(), f"PO_{safe_filename(po)}.pdf")

    options = {
        'page-size': 'A4',
        'encoding': 'UTF-8',
        'margin-top': '5mm',
        'margin-bottom': '5mm',
        'margin-left': '5mm',
        'margin-right': '5mm',
        'enable-local-file-access': None
    }

    pdfkit.from_string(html, output, configuration=config, options=options)

    filename = f"PO_{safe_filename(po)}.pdf"

    return jsonify({
        "download_url": f"/download/{filename}"
    })

@app.route('/download-excel', methods=['POST'])
def download_excel():
    data = request.form

    rows = []
    rows.append(["Category", "Item Description", "Unit", "Quantity", "Remarks"])

    cctvItems = [
        'SITC of 2 MP Indoor Bullet IP Camera',
        'SITC of 2 MP Indoor Dome IP Camera',
        'Camera Box',
        '12U Rack with other hardware',
        '6U Rack with other hardware',
        '9U Rack with other hardware',
        '16 Port POE Switch with 2 SFP Port',
        '24 Port POE Switch with 2 SFP Port',
        'SITC of 8Tb Surveillance HDD',
        'Patch Chord 1 Mtr',
        '24 Port Patch Panel',
        '1U Cable Manager',
        'Supply & Laying of Cat-6 Cable (Only Copper Wire)',
        'SITC of 32CH NVR 4 Sata',
        'RJ45 Connector',
        'Supply & laying 25mm PVC Conduits with accessories',
        'HDMI Switcher (Vanition Switch With 2 HDMI Cables - 5 MTR)',
        'Wall Mount Stand For LED',
        'Wireless Mouse',
        'Adjustable CEILING MOUNT STAND',
        'USB Extender',
        'Display Quality Status',
        '32" LED Monitor',
        'HDMI Cable 15 MTR',
        '8 Port POE Switch',
        'Installation Cost'
    ]

    for i in range(1, 27):
        rows.append([
            "CCTV" if i == 1 else "",
            cctvItems[i-1],
            data.get(f'unit{i}', ''),
            data.get(f'qty{i}', ''),
            data.get(f'remark{i}', '')
        ])

    rows.append([
        "Networking",
        "Installation of Access Point",
        data.get('unit27', ''),
        data.get('qty27', ''),
        data.get('remark27', '')
    ])

    df = pd.DataFrame(rows)

    filename = "cctv_form_data.xlsx"
    file_path = os.path.join(tempfile.gettempdir(), filename)

    df.to_excel(file_path, index=False, header=False)

    return jsonify({
        "download_url": f"/download-excel-file/{filename}"
    })

@app.route('/download-excel-file/<filename>')
def download_excel_file(filename):
    path = os.path.join(tempfile.gettempdir(), filename)
    return send_file(
        path,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/download/<filename>')
def download(filename):
    path = os.path.join(tempfile.gettempdir(), filename)
    return send_file(path, as_attachment=True, download_name=filename, mimetype='application/pdf')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=True)

import os
import pandas as pd
import qrcode
import barcode
from barcode.writer import ImageWriter
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from flask import Flask, render_template, request, send_file
from xhtml2pdf import pisa

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Helper: Set Cell Margins (Padding)
def set_cell_margins(cell, top=60, bottom=60, left=80, right=80):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

# Helper: Set Cell Border
def set_cell_border(cell, color="A0A0A0", sz="6", val="single"):
    tcPr = cell._element.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    for edge in ('top', 'left', 'bottom', 'right'):
        element = OxmlElement(f'w:{edge}')
        element.set(qn('w:val'), val)
        element.set(qn('w:sz'), sz)
        element.set(qn('w:space'), '0')
        element.set(qn('w:color'), color)
        tcBorders.append(element)

# Generate Word Document (.docx) - Exact Sticker Size: 9.5 cm x 7.8 cm
def build_docx(items):
    doc = Document()
    
    # Page setup for A4
    for section in doc.sections:
        section.top_margin = Cm(1.0)
        section.bottom_margin = Cm(1.0)
        section.left_margin = Cm(1.0)
        section.right_margin = Cm(1.0)

    chunk_size = 6
    chunks = [items[i:i + chunk_size] for i in range(0, len(items), chunk_size)]

    for page_idx, chunk in enumerate(chunks):
        if page_idx > 0:
            doc.add_page_break()

        grid_table = doc.add_table(rows=3, cols=2)
        grid_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        grid_table.autofit = False

        for i, row in enumerate(chunk):
            row_idx = i // 2
            col_idx = i % 2
            
            # Set exact height & width for each cell (9.5 cm x 7.8 cm)
            grid_table.rows[row_idx].height = Cm(7.8)
            cell = grid_table.cell(row_idx, col_idx)
            cell.width = Cm(9.5)

            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            set_cell_border(cell, color="808080", sz="6")

            fields = [
                ("Model", row.get('Model', 'N/A')),
                ("Serial No", row.get('Serial No', 'N/A')),
                ("Processor Details", row.get('Processor Details', 'N/A')),
                ("Storage", f"{row.get('Harddisk','N/A')} | {row.get('SSD','N/A')} | {row.get('RAM','N/A')}"),
                ("IP No", row.get('IP No', 'N/A')),
                ("Hostname", row.get('Hostname', 'N/A')),
                ("MAC address", row.get('MAC address', 'N/A'))
            ]

            qr_content = "\n".join([f"{l}: {v}" for l, v in fields])
            qr_file = os.path.abspath(f"temp_qr_{page_idx}_{i}.png")
            bar_base = os.path.abspath(f"temp_bar_{page_idx}_{i}")

            qr_img = qrcode.make(qr_content)
            qr_img.save(qr_file)

            code128 = barcode.get_barcode_class('code128')
            serial_barcode = str(row.get('Serial No', '000000'))
            barcode_obj = code128(serial_barcode, writer=ImageWriter())
            barcode_file = barcode_obj.save(bar_base, options={"write_text": False})

            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)

            run_title = p.add_run("IT ASSET TAG\n")
            run_title.bold = True
            run_title.font.size = Pt(11)

            for label, val in fields:
                p_item = cell.add_paragraph()
                p_item.paragraph_format.space_before = Pt(0)
                p_item.paragraph_format.space_after = Pt(2)

                lbl_run = p_item.add_run(f"{label}: ")
                lbl_run.bold = True
                lbl_run.font.size = Pt(8.5)

                val_run = p_item.add_run(str(val))
                val_run.font.size = Pt(8.5)

            bottom_table = cell.add_table(rows=1, cols=2)
            bottom_table.alignment = WD_TABLE_ALIGNMENT.CENTER

            cell_qr = bottom_table.cell(0, 0)
            cell_bar = bottom_table.cell(0, 1)

            p_qr = cell_qr.paragraphs[0]
            p_qr.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_qr.add_run().add_picture(qr_file, width=Cm(2.2))

            p_bar = cell_bar.paragraphs[0]
            p_bar.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_bar.add_run().add_picture(barcode_file, width=Cm(3.8), height=Cm(1.0))

            p_bar_txt = cell_bar.add_paragraph()
            p_bar_txt.alignment = WD_ALIGN_PARAGRAPH.CENTER
            txt_run = p_bar_txt.add_run(f"*{serial_barcode}*")
            txt_run.font.size = Pt(8.0)

            if os.path.exists(qr_file): os.remove(qr_file)
            if os.path.exists(barcode_file): os.remove(barcode_file)

    output_path = os.path.join(OUTPUT_FOLDER, "Asset_Tags.docx")
    doc.save(output_path)
    return output_path

# Generate PDF Document (.pdf) - Exact Sticker Size: 9.5 cm x 7.8 cm
def build_pdf(items):
    temp_images = []
    tags_html = ""

    for i, row in enumerate(items):
        fields = [
            ("Model", row.get('Model', 'N/A')),
            ("Serial No", row.get('Serial No', 'N/A')),
            ("Processor Details", row.get('Processor Details', 'N/A')),
            ("Storage", f"{row.get('Harddisk','N/A')} | {row.get('SSD','N/A')} | {row.get('RAM','N/A')}"),
            ("IP No", row.get('IP No', 'N/A')),
            ("Hostname", row.get('Hostname', 'N/A')),
            ("MAC address", row.get('MAC address', 'N/A'))
        ]

        qr_content = "\n".join([f"{l}: {v}" for l, v in fields])
        qr_file = os.path.abspath(f"temp_qr_pdf_{i}.png")
        bar_base = os.path.abspath(f"temp_bar_pdf_{i}")

        qr_img = qrcode.make(qr_content)
        qr_img.save(qr_file)

        code128 = barcode.get_barcode_class('code128')
        serial_barcode = str(row.get('Serial No', '000000'))
        barcode_obj = code128(serial_barcode, writer=ImageWriter())
        barcode_file = barcode_obj.save(bar_base, options={"write_text": False})

        temp_images.extend([qr_file, barcode_file])

        tags_html += f"""
        <div class="tag-box">
            <div class="title">IT ASSET TAG</div>
            <div class="field"><b>Model:</b> {row.get('Model', 'N/A')}</div>
            <div class="field"><b>Serial No:</b> {row.get('Serial No', 'N/A')}</div>
            <div class="field"><b>Processor Details:</b> {row.get('Processor Details', 'N/A')}</div>
            <div class="field"><b>Storage:</b> {row.get('Harddisk','N/A')} | {row.get('SSD','N/A')} | {row.get('RAM','N/A')}</div>
            <div class="field"><b>IP No:</b> {row.get('IP No', 'N/A')}</div>
            <div class="field"><b>Hostname:</b> {row.get('Hostname', 'N/A')}</div>
            <div class="field"><b>MAC address:</b> {row.get('MAC address', 'N/A')}</div>
            
            <table class="img-table">
                <tr>
                    <td style="width:35%; text-align:left; vertical-align:middle;">
                        <img src="{qr_file}" width="65" height="65" />
                    </td>
                    <td style="width:65%; text-align:center; vertical-align:middle;">
                        <img src="{barcode_file}" width="120" height="32" /><br>
                        <span style="font-size: 8px; font-weight: bold;">*{serial_barcode}*</span>
                    </td>
                </tr>
            </table>
        </div>
        """

    html_full = f"""
    <html>
    <head>
        <style>
            @page {{ size: A4; margin: 0.6cm; }}
            body {{ font-family: Helvetica, Arial, sans-serif; margin:0; padding:0; }}
            .tag-box {{
                width: 9.5cm;
                height: 7.8cm;
                border: 1px solid #333;
                padding: 8px 10px;
                margin: 4px;
                display: inline-block;
                vertical-align: top;
                box-sizing: border-box;
                overflow: hidden;
            }}
            .title {{ font-weight: bold; font-size: 11px; margin-bottom: 5px; text-decoration: underline; }}
            .field {{ font-size: 8.5px; margin-bottom: 3px; line-height: 1.1; }}
            .img-table {{ width: 100%; margin-top: 6px; border-collapse: collapse; }}
        </style>
    </head>
    <body>
        {tags_html}
    </body>
    </html>
    """

    output_path = os.path.join(OUTPUT_FOLDER, "Asset_Tags.pdf")
    with open(output_path, "w+b") as out_f:
        pisa.CreatePDF(html_full, dest=out_f)

    for img in temp_images:
        if os.path.exists(img): os.remove(img)

    return output_path

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    input_type = request.form.get('input_type')
    file_type = request.form.get('file_type')

    items = []

    # 1. Excel File Upload
    if input_type == 'excel':
        file = request.files.get('file')
        if not file or file.filename == '':
            return "No excel file selected!"
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)
        
        df = pd.read_excel(file_path)
        items = df.to_dict('records')

    # 2. Direct Web Entry
    elif input_type == 'manual':
        models = request.form.getlist('model[]')
        serials = request.form.getlist('serial[]')
        processors = request.form.getlist('processor[]')
        rams = request.form.getlist('ram[]')
        harddisks = request.form.getlist('harddisk[]')
        ssds = request.form.getlist('ssd[]')
        ips = request.form.getlist('ip[]')
        hostnames = request.form.getlist('hostname[]')
        macs = request.form.getlist('mac[]')

        for i in range(len(models)):
            if models[i].strip() or serials[i].strip():
                items.append({
                    'Model': models[i],
                    'Serial No': serials[i],
                    'Processor Details': processors[i],
                    'RAM': rams[i],
                    'Harddisk': harddisks[i],
                    'SSD': ssds[i],
                    'IP No': ips[i],
                    'Hostname': hostnames[i],
                    'MAC address': macs[i]
                })

    if not items:
        return "No data provided!"

    if file_type == 'word':
        out_file = build_docx(items)
        return send_file(out_file, as_attachment=True)
    else:
        out_file = build_pdf(items)
        return send_file(out_file, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)

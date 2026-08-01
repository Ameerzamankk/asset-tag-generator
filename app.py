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

# Helper: Word Cell Margins
def set_cell_margins(cell, top=50, bottom=50, left=70, right=70):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

# Helper: Word Cell Border (Gray)
def set_cell_border(cell, color="CCCCCC", sz="8", val="single"):
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

# Generate Word Document (.docx) - Clean Gray Border
def build_docx(items):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Cm(1.0)
        section.bottom_margin = Cm(1.0)
        section.left_margin = Cm(0.8)
        section.right_margin = Cm(0.8)

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
            
            grid_table.rows[row_idx].height = Cm(8.0)
            cell = grid_table.cell(row_idx, col_idx)
            cell.width = Cm(9.0)

            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            set_cell_border(cell, color="CCCCCC", sz="8")

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
            qr_file = os.path.abspath(f"temp_qr_docx_{page_idx}_{i}.png")
            bar_base = os.path.abspath(f"temp_bar_docx_{page_idx}_{i}")

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
            run_title.font.name = 'Arial'
            run_title.font.size = Pt(10)

            for label, val in fields:
                p_item = cell.add_paragraph()
                p_item.paragraph_format.space_before = Pt(0)
                p_item.paragraph_format.space_after = Pt(1)

                lbl_run = p_item.add_run(f"{label}: ")
                lbl_run.bold = True
                lbl_run.font.name = 'Arial'
                lbl_run.font.size = Pt(8)

                val_run = p_item.add_run(str(val))
                val_run.font.name = 'Arial'
                val_run.font.size = Pt(8)

            bottom_table = cell.add_table(rows=1, cols=2)
            bottom_table.alignment = WD_TABLE_ALIGNMENT.CENTER

            cell_qr = bottom_table.cell(0, 0)
            cell_bar = bottom_table.cell(0, 1)

            p_qr = cell_qr.paragraphs[0]
            p_qr.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_qr.add_run().add_picture(qr_file, width=Cm(1.8))

            p_bar = cell_bar.paragraphs[0]
            p_bar.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_bar.add_run().add_picture(barcode_file, width=Cm(3.8), height=Cm(0.9))

            p_bar_txt = cell_bar.add_paragraph()
            p_bar_txt.alignment = WD_ALIGN_PARAGRAPH.CENTER
            txt_run = p_bar_txt.add_run(f"*{serial_barcode}*")
            txt_run.font.name = 'Arial'
            txt_run.font.size = Pt(7)

            if os.path.exists(qr_file): os.remove(qr_file)
            if os.path.exists(barcode_file): os.remove(barcode_file)

    output_path = os.path.join(OUTPUT_FOLDER, "Asset_Tags.docx")
    doc.save(output_path)
    return output_path

# Generate PDF Document (.pdf) - Clean Layout with Gray Border & No Inner Columns
def build_pdf(items):
    temp_images = []
    
    chunk_size = 6
    chunks = [items[i:i + chunk_size] for i in range(0, len(items), chunk_size)]
    
    pages_html = ""

    for page_idx, chunk in enumerate(chunks):
        stickers_in_page = ""
        
        for i, row in enumerate(chunk):
            idx = page_idx * 6 + i
            
            # Position logic: 2 columns x 3 rows
            col = i % 2
            r = i // 2
            
            left_pos = "0.2cm" if col == 0 else "9.6cm"
            top_pos = f"{r * 8.4}cm"
            
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
            qr_file = os.path.abspath(f"temp_qr_pdf_{idx}.png")
            bar_base = os.path.abspath(f"temp_bar_pdf_{idx}")

            qr_img = qrcode.make(qr_content)
            qr_img.save(qr_file)

            code128 = barcode.get_barcode_class('code128')
            serial_barcode = str(row.get('Serial No', '000000'))
            barcode_obj = code128(serial_barcode, writer=ImageWriter())
            barcode_file = barcode_obj.save(bar_base, options={"write_text": False})

            temp_images.extend([qr_file, barcode_file])

            field_lines = ""
            for label, val in fields:
                field_lines += f"""
                <div class="line">
                    <span class="lbl">{label}:</span>
                    <span class="val">{val}</span>
                </div>
                """

            sticker_box = f"""
            <div class="sticker-card" style="left: {left_pos}; top: {top_pos};">
                <div class="sticker-header">IT ASSET TAG</div>
                <div class="fields-container">
                    {field_lines}
                </div>
                <table class="code-table">
                    <tr>
                        <td style="width: 35%; text-align: left; vertical-align: bottom;">
                            <img src="{qr_file}" width="50" height="50" />
                        </td>
                        <td style="width: 65%; text-align: center; vertical-align: bottom;">
                            <img src="{barcode_file}" width="115" height="26" /><br/>
                            <span class="serial-txt">*{serial_barcode}*</span>
                        </td>
                    </tr>
                </table>
            </div>
            """
            stickers_in_page += sticker_box

        page_break_css = 'page-break-after: always;' if page_idx < len(chunks) - 1 else ''
        pages_html += f"""
        <div class="page-container" style="{page_break_css}">
            {stickers_in_page}
        </div>
        """

    html_full = f"""
    <html>
    <head>
        <style>
            @page {{
                size: A4 portrait;
                margin: 1.0cm 0.8cm;
            }}
            body {{
                font-family: Arial, Helvetica, sans-serif;
                margin: 0;
                padding: 0;
                background-color: #ffffff;
            }}
            .page-container {{
                position: relative;
                width: 19.0cm;
                height: 26.5cm;
            }}
            
            /* Clean Sticker Styling with Gray Border & No Column Dividers */
            .sticker-card {{
                position: absolute;
                width: 9.0cm;
                height: 8.0cm;
                border: 1px solid #cccccc;
                padding: 8px 10px;
                box-sizing: border-box;
                background-color: #ffffff;
            }}
            .sticker-header {{
                font-family: Arial, Helvetica, sans-serif;
                font-weight: bold;
                font-size: 10.5pt;
                text-align: left;
                color: #000000;
                margin-bottom: 6px;
                text-decoration: underline;
            }}
            .fields-container {{
                margin-bottom: 6px;
            }}
            .line {{
                font-size: 8pt;
                line-height: 1.3;
                margin-bottom: 2px;
            }}
            .lbl {{
                font-weight: bold;
                color: #000000;
                display: inline-block;
            }}
            .val {{
                color: #222222;
            }}
            .code-table {{
                width: 100%;
                border-collapse: collapse;
                position: absolute;
                bottom: 8px;
                left: 10px;
                right: 10px;
            }}
            .serial-txt {{
                font-size: 7.5pt;
                font-weight: bold;
                color: #222222;
            }}
        </style>
    </head>
    <body>
        {pages_html}
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

    if input_type == 'excel':
        file = request.files.get('file')
        if not file or file.filename == '':
            return "No excel file selected!"
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)
        
        df = pd.read_excel(file_path)
        items = df.to_dict('records')

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

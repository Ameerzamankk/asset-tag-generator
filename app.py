import os
import pandas as pd
import qrcode
import barcode
from barcode.writer import ImageWriter
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from flask import Flask, render_template, request, send_file

app = Flask(__name__)

# ഫോൾഡറുകൾ സജ്ജമാക്കുന്നു
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def set_cell_margins(cell, top=60, bottom=60, left=80, right=80):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

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

def generate_docx_from_excel(excel_path):
    df = pd.read_excel(excel_path)
    doc = Document()

    # Page Margins
    for section in doc.sections:
        section.top_margin = Inches(0.35)
        section.bottom_margin = Inches(0.35)
        section.left_margin = Inches(0.35)
        section.right_margin = Inches(0.35)

    # 6 എണ്ണം വീതമുള്ള ഗ്രൂപ്പുകളാക്കുന്നു
    chunk_size = 6
    chunks = [df[i:i + chunk_size] for i in range(0, len(df), chunk_size)]

    for page_idx, chunk in enumerate(chunks):
        if page_idx > 0:
            doc.add_page_break()

        # 3 Rows x 2 Cols Layout
        grid_table = doc.add_table(rows=3, cols=2)
        grid_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        grid_table.autofit = False

        items = chunk.to_dict('records')
        for i, row in enumerate(items):
            row_idx = i // 2
            col_idx = i % 2
            cell = grid_table.cell(row_idx, col_idx)
            cell.width = Inches(3.65)

            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            set_cell_border(cell, color="A0A0A0", sz="6")

            def get_val(column_name):
                val = row.get(column_name, '')
                return str(val).strip() if pd.notna(val) else 'N/A'

            model = get_val('Model')
            serial = get_val('Serial No')
            ram = get_val('RAM')
            processor = get_val('Processor Details')
            harddisk = get_val('Harddisk')
            ssd = get_val('SSD')
            ip = get_val('IP No')
            hostname = get_val('Hostname')
            mac = get_val('MAC address')

            storage_combined = f"{harddisk} | {ssd} | {ram}"

            fields = [
                ("Model", model),
                ("Serial No", serial),
                ("Processor Details", processor),
                ("Storage", storage_combined),
                ("IP No", ip),
                ("Hostname", hostname),
                ("MAC address", mac)
            ]

            qr_content = "\n".join([f"{label}: {val}" for label, val in fields])
            qr_file = f"temp_qr_{page_idx}_{i}.png"
            bar_base = f"temp_bar_{page_idx}_{i}"

            qr_img = qrcode.make(qr_content)
            qr_img.save(qr_file)

            code128 = barcode.get_barcode_class('code128')
            serial_barcode = serial if serial != 'N/A' else '000000'
            barcode_obj = code128(serial_barcode, writer=ImageWriter())
            barcode_file = barcode_obj.save(bar_base, options={"write_text": False})

            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)

            run_title = p.add_run("IT ASSET TAG\n")
            run_title.bold = True
            run_title.font.size = Pt(10.5)
            run_title.font.name = 'Arial'

            for label, val in fields:
                p_item = cell.add_paragraph()
                p_item.paragraph_format.space_before = Pt(0)
                p_item.paragraph_format.space_after = Pt(1)
                p_item.paragraph_format.line_spacing = 1.0

                lbl_run = p_item.add_run(f"{label}: ")
                lbl_run.bold = True
                lbl_run.font.size = Pt(8.0)
                lbl_run.font.name = 'Arial'

                val_run = p_item.add_run(val)
                val_run.font.size = Pt(8.0)
                val_run.font.name = 'Arial'

            p_space = cell.add_paragraph()
            p_space.paragraph_format.space_before = Pt(2)
            p_space.paragraph_format.space_after = Pt(0)

            bottom_table = cell.add_table(rows=1, cols=2)
            bottom_table.alignment = WD_TABLE_ALIGNMENT.CENTER

            cell_qr = bottom_table.cell(0, 0)
            cell_bar = bottom_table.cell(0, 1)

            p_qr = cell_qr.paragraphs[0]
            p_qr.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_qr.add_run().add_picture(qr_file, width=Inches(0.85))

            p_bar = cell_bar.paragraphs[0]
            p_bar.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_bar.add_run().add_picture(barcode_file, width=Inches(1.4), height=Inches(0.35))

            p_bar_txt = cell_bar.add_paragraph()
            p_bar_txt.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_bar_txt.paragraph_format.space_before = Pt(1)
            p_bar_txt.paragraph_format.space_after = Pt(0)
            txt_run = p_bar_txt.add_run(f"*{serial}*")
            txt_run.font.size = Pt(7.5)
            txt_run.font.name = 'Arial'

            if os.path.exists(qr_file):
                os.remove(qr_file)
            if os.path.exists(barcode_file):
                os.remove(barcode_file)

    output_path = os.path.join(OUTPUT_FOLDER, "Bulk_Asset_Tags.docx")
    doc.save(output_path)
    return output_path

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'എക്സെൽ ഫയൽ അപ്‌ലോഡ് ചെയ്തിട്ടില്ല!'
    file = request.files['file']
    if file.filename == '':
        return 'ഫയൽ ഒന്നും സെലക്ട് ചെയ്തിട്ടില്ല!'

    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)

    output_docx = generate_docx_from_excel(file_path)
    return send_file(output_docx, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
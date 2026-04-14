from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from datetime import datetime
from io import BytesIO
import zipfile, re, os

app = Flask(__name__)
CORS(app)

MESES = ['enero','febrero','marzo','abril','mayo','junio','julio',
         'agosto','septiembre','octubre','noviembre','diciembre']
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'plantilla.xlsx')

# Style indices from original template
S_TEXT='31'; S_TEXT2='51'; S_NUM='15'; S_NUM2='40'; S_FRAC='23'
S_HEAD='1';  S_DATA='38';  S_PRICE='32'

# Section header rows — must NEVER be cleared
# B17=Desabollar-Reparar(46), B39=Pintar(60), B48=Desmontar y montar(57), B56=Repuestos(57)
HEADER_ROWS = {17, 39, 48, 56}

# Repair rows: 18-38 (under Desabollar), 40-47 (under Pintar), 49-55 (under Desmontar)
REPAIR_ROWS = list(range(18, 39)) + list(range(40, 48)) + list(range(49, 56))
# Parts rows: 57-106
PARTS_ROWS = list(range(57, 107))

def make_text_cell(ref, style, idx):
    return f'<c r="{ref}" s="{style}" t="s"><v>{idx}</v></c>'

def make_num_cell(ref, style, value):
    return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'

def replace_cell(xml, ref, new_cell):
    p1 = rf'<c r="{re.escape(ref)}"[^/]*/>'
    p2 = rf'<c r="{re.escape(ref)}"[^>]*>.*?</c>'
    if re.search(p1, xml):
        return re.sub(p1, new_cell, xml)
    if re.search(p2, xml, re.DOTALL):
        return re.sub(p2, new_cell, xml, flags=re.DOTALL)
    return xml

def clear_cell(xml, ref, style):
    empty = f'<c r="{ref}" s="{style}"/>'
    p1 = rf'<c r="{re.escape(ref)}"[^/]*/>'
    p2 = rf'<c r="{re.escape(ref)}"[^>]*>.*?</c>'
    if re.search(p1, xml):
        return re.sub(p1, empty, xml)
    if re.search(p2, xml, re.DOTALL):
        return re.sub(p2, empty, xml, flags=re.DOTALL)
    return xml

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

@app.route('/generar-excel', methods=['POST'])
def generar_excel():
    try:
        data = request.json or {}

        with zipfile.ZipFile(TEMPLATE_PATH, 'r') as z:
            all_files = {n: z.read(n) for n in z.namelist()}

        sheet_xml = all_files['xl/worksheets/sheet1.xml'].decode('utf-8')
        shared_xml = all_files['xl/sharedStrings.xml'].decode('utf-8')

        # Parse shared strings
        ss_items = re.findall(r'<si>.*?</si>', shared_xml, re.DOTALL)
        ss_list = []
        for item in ss_items:
            m = re.search(r'<t[^>]*>(.*?)</t>', item, re.DOTALL)
            ss_list.append(m.group(1) if m else '')

        def ss(text):
            safe = (str(text)
                .replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;'))
            for i, s in enumerate(ss_list):
                if s == safe:
                    return i
            ss_list.append(safe)
            return len(ss_list) - 1

        now = datetime.now()
        pat = (data.get('patente') or '').upper().strip()
        fecha = f" Fecha {now.day} de {MESES[now.month-1]} {now.year}"

        # Header
        sheet_xml = replace_cell(sheet_xml, 'B7',  make_text_cell('B7',  S_HEAD, ss(f'Presupuesto {pat}')))
        sheet_xml = replace_cell(sheet_xml, 'D7',  make_text_cell('D7',  S_HEAD, ss(fecha)))
        sheet_xml = replace_cell(sheet_xml, 'C13', make_text_cell('C13', S_DATA, ss(data.get('marca') or '')))
        sheet_xml = replace_cell(sheet_xml, 'F13', make_text_cell('F13', S_DATA, ss(data.get('modelo') or '')))
        sheet_xml = replace_cell(sheet_xml, 'C14', make_text_cell('C14', S_DATA, ss(pat)))

        if data.get('anio'):
            try: sheet_xml = replace_cell(sheet_xml, 'I13', make_num_cell('I13', S_NUM, int(data['anio'])))
            except: pass
        if data.get('km'):
            try: sheet_xml = replace_cell(sheet_xml, 'F14', make_num_cell('F14', S_NUM2, int(data['km'])))
            except: pass
        if data.get('combustible'):
            try: sheet_xml = replace_cell(sheet_xml, 'I14', make_num_cell('I14', S_FRAC, int(data['combustible']) / 4))
            except: pass

        # Clear repair rows (skip headers)
        for row in REPAIR_ROWS:
            sheet_xml = clear_cell(sheet_xml, f'B{row}', S_TEXT)

        # Clear parts rows
        for row in PARTS_ROWS:
            sheet_xml = clear_cell(sheet_xml, f'B{row}', S_TEXT)
            sheet_xml = clear_cell(sheet_xml, f'I{row}', S_PRICE)

        # Fill repairs across all sections
        trabajos = data.get('trabajos') or []
        for i, t in enumerate(trabajos):
            if i < len(REPAIR_ROWS) and t:
                row = REPAIR_ROWS[i]
                sheet_xml = replace_cell(sheet_xml, f'B{row}',
                    make_text_cell(f'B{row}', S_TEXT, ss(t)))

        # Fill parts
        idx_gama = ss('GAMA')
        repuestos = data.get('repuestos') or []
        for i, r in enumerate(repuestos):
            if i < len(PARTS_ROWS) and r:
                row = PARTS_ROWS[i]
                sheet_xml = replace_cell(sheet_xml, f'B{row}',
                    make_text_cell(f'B{row}', S_TEXT, ss(r)))
                sheet_xml = replace_cell(sheet_xml, f'I{row}',
                    make_text_cell(f'I{row}', '17', idx_gama))

        # Observations
        obs = data.get('observaciones') or ''
        sheet_xml = replace_cell(sheet_xml, 'B113',
            make_text_cell('B113', S_TEXT2, ss(obs)))

        # Rebuild shared strings
        new_ss = ''.join(
            f'<si><t xml:space="preserve">{s}</t></si>' for s in ss_list)
        new_shared = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'count="{len(ss_list)}" uniqueCount="{len(ss_list)}">'
            f'{new_ss}</sst>'
        )

        # Write xlsx preserving ALL original files (image, styles, formulas)
        buf = BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for name, fbytes in all_files.items():
                if name == 'xl/worksheets/sheet1.xml':
                    zout.writestr(name, sheet_xml.encode('utf-8'))
                elif name == 'xl/sharedStrings.xml':
                    zout.writestr(name, new_shared.encode('utf-8'))
                else:
                    zout.writestr(name, fbytes)

        buf.seek(0)
        return send_file(buf,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'Presupuesto_{pat}.xlsx')

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)

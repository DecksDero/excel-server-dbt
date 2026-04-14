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

# Section header rows — must NEVER be cleared
HEADER_ROWS = {17, 39, 48, 56}
REPAIR_ROWS = list(range(18, 39)) + list(range(40, 48)) + list(range(49, 56))
PARTS_ROWS  = list(range(57, 107))

def get_cell_style(xml, ref):
    """Get the original style index of a cell without changing it."""
    p = rf'<c r="{re.escape(ref)}"[^>]*(?:/>|>.*?</c>)'
    m = re.search(p, xml, re.DOTALL)
    if m:
        sm = re.search(r's="(\d+)"', m.group(0))
        return sm.group(1) if sm else '0'
    return '0'

def clear_cell_preserve_style(xml, ref):
    """Remove cell value but keep original style — preserves all borders/formatting."""
    original_style = get_cell_style(xml, ref)
    empty = f'<c r="{ref}" s="{original_style}"/>'
    # Match self-closing: <c r="X" ... />
    p1 = rf'<c r="{re.escape(ref)}"(?:\s[^>]*)?\s*/>'
    # Match with content: <c r="X" ...>...</c>
    p2 = rf'<c r="{re.escape(ref)}"[^>]*>.*?</c>'
    result = re.sub(p1, empty, xml)
    if result != xml:
        return result
    return re.sub(p2, empty, xml, flags=re.DOTALL)

def set_cell_value(xml, ref, new_cell):
    """Replace any cell form (empty or with value) with new_cell."""
    p1 = rf'<c r="{re.escape(ref)}"(?:\s[^>]*)?\s*/>'
    p2 = rf'<c r="{re.escape(ref)}"[^>]*>.*?</c>'
    result = re.sub(p1, new_cell, xml)
    if result != xml:
        return result
    return re.sub(p2, new_cell, xml, flags=re.DOTALL)

def make_text_cell(ref, style, idx):
    return f'<c r="{ref}" s="{style}" t="s"><v>{idx}</v></c>'

def make_num_cell(ref, style, value):
    return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'

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
        ss_raw = re.findall(r'<si>.*?</si>', shared_xml, re.DOTALL)
        ss_list = []
        for item in ss_raw:
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

        # Pre-read original styles for all cells we'll touch
        # This ensures clear_cell_preserve_style has the right style BEFORE any changes
        original_b_styles = {}
        original_i_styles = {}
        for row in REPAIR_ROWS:
            original_b_styles[row] = get_cell_style(sheet_xml, f'B{row}')
        for row in PARTS_ROWS:
            original_b_styles[row] = get_cell_style(sheet_xml, f'B{row}')
            original_i_styles[row] = get_cell_style(sheet_xml, f'I{row}')

        now = datetime.now()
        pat = (data.get('patente') or '').upper().strip()
        fecha = f" Fecha {now.day} de {MESES[now.month-1]} {now.year}"

        # Header styles (read once from original)
        s_b7  = get_cell_style(sheet_xml, 'B7')
        s_d7  = get_cell_style(sheet_xml, 'D7')
        s_c13 = get_cell_style(sheet_xml, 'C13')
        s_f13 = get_cell_style(sheet_xml, 'F13')
        s_i13 = get_cell_style(sheet_xml, 'I13')
        s_c14 = get_cell_style(sheet_xml, 'C14')
        s_f14 = get_cell_style(sheet_xml, 'F14')
        s_i14 = get_cell_style(sheet_xml, 'I14')
        s_b113 = get_cell_style(sheet_xml, 'B113')

        # Fill header
        sheet_xml = set_cell_value(sheet_xml, 'B7',  make_text_cell('B7',  s_b7,  ss(f'Presupuesto {pat}')))
        sheet_xml = set_cell_value(sheet_xml, 'D7',  make_text_cell('D7',  s_d7,  ss(fecha)))
        sheet_xml = set_cell_value(sheet_xml, 'C13', make_text_cell('C13', s_c13, ss(data.get('marca') or '')))
        sheet_xml = set_cell_value(sheet_xml, 'F13', make_text_cell('F13', s_f13, ss(data.get('modelo') or '')))
        sheet_xml = set_cell_value(sheet_xml, 'C14', make_text_cell('C14', s_c14, ss(pat)))

        if data.get('anio'):
            try: sheet_xml = set_cell_value(sheet_xml, 'I13', make_num_cell('I13', s_i13, int(data['anio'])))
            except: pass
        if data.get('km'):
            try: sheet_xml = set_cell_value(sheet_xml, 'F14', make_num_cell('F14', s_f14, int(data['km'])))
            except: pass
        if data.get('combustible'):
            try: sheet_xml = set_cell_value(sheet_xml, 'I14', make_num_cell('I14', s_i14, int(data['combustible']) / 4))
            except: pass

        # Clear all repair and parts cells (preserving original styles)
        for row in REPAIR_ROWS:
            style = original_b_styles[row]
            empty = f'<c r="B{row}" s="{style}"/>'
            p1 = rf'<c r="B{row}"(?:\s[^>]*)?\s*/>'
            p2 = rf'<c r="B{row}"[^>]*>.*?</c>'
            result = re.sub(p1, empty, sheet_xml)
            sheet_xml = result if result != sheet_xml else re.sub(p2, empty, sheet_xml, flags=re.DOTALL)

        for row in PARTS_ROWS:
            # Clear B
            b_style = original_b_styles[row]
            b_empty = f'<c r="B{row}" s="{b_style}"/>'
            p1b = rf'<c r="B{row}"(?:\s[^>]*)?\s*/>'
            p2b = rf'<c r="B{row}"[^>]*>.*?</c>'
            result = re.sub(p1b, b_empty, sheet_xml)
            sheet_xml = result if result != sheet_xml else re.sub(p2b, b_empty, sheet_xml, flags=re.DOTALL)
            # Clear I — preserve original style (17 for all I57-I106)
            i_style = original_i_styles[row]
            i_empty = f'<c r="I{row}" s="{i_style}"/>'
            p1i = rf'<c r="I{row}"(?:\s[^>]*)?\s*/>'
            p2i = rf'<c r="I{row}"[^>]*>.*?</c>'
            result = re.sub(p1i, i_empty, sheet_xml)
            sheet_xml = result if result != sheet_xml else re.sub(p2i, i_empty, sheet_xml, flags=re.DOTALL)

        # Fill repairs
        trabajos = data.get('trabajos') or []
        for i, t in enumerate(trabajos):
            if i < len(REPAIR_ROWS) and t:
                row = REPAIR_ROWS[i]
                style = original_b_styles[row]
                sheet_xml = set_cell_value(sheet_xml, f'B{row}',
                    make_text_cell(f'B{row}', style, ss(t)))

        # Fill parts
        idx_gama = ss('GAMA')
        repuestos = data.get('repuestos') or []
        for i, r in enumerate(repuestos):
            if i < len(PARTS_ROWS) and r:
                row = PARTS_ROWS[i]
                b_style = original_b_styles[row]
                i_style = original_i_styles[row]
                sheet_xml = set_cell_value(sheet_xml, f'B{row}',
                    make_text_cell(f'B{row}', b_style, ss(r)))
                sheet_xml = set_cell_value(sheet_xml, f'I{row}',
                    make_text_cell(f'I{row}', i_style, idx_gama))

        # Observations
        obs = data.get('observaciones') or ''
        sheet_xml = set_cell_value(sheet_xml, 'B113',
            make_text_cell('B113', s_b113, ss(obs)))

        # Rebuild shared strings
        new_ss = ''.join(
            f'<si><t xml:space="preserve">{s}</t></si>' for s in ss_list)
        new_shared = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'count="{len(ss_list)}" uniqueCount="{len(ss_list)}">'
            f'{new_ss}</sst>'
        )

        # Write xlsx preserving ALL original files
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

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

REPAIR_ROWS = list(range(18,39)) + list(range(40,48)) + list(range(49,56))
PARTS_ROWS  = list(range(57,107))
# Header rows NEVER touched: 17 (Desabollar), 39 (Pintar), 48 (Desmontar), 56 (Repuestos)

def get_style(sheet_xml, ref):
    """Read original style index for a cell."""
    m = re.search(rf'<c r="{re.escape(ref)}" s="(\d+)"', sheet_xml)
    return m.group(1) if m else '31'

def replace_cell(sheet_xml, ref, new_cell):
    """
    Replace a cell safely. Uses two non-overlapping patterns:
    - Pattern A: self-closing  <c r="REF" s="N"/>
    - Pattern B: with content  <c r="REF" s="N" ...>...</c>  ([^/]* prevents matching />)
    Only one pattern will match — they are mutually exclusive.
    """
    esc = re.escape(ref)
    # A: self-closing (ends with />)
    p_self = rf'<c r="{esc}" s="\d+"/>'
    result = re.sub(p_self, new_cell, sheet_xml)
    if result != sheet_xml:
        return result
    # B: with content ([^/]* ensures we don't match self-closing />)
    p_content = rf'<c r="{esc}" s="\d+"[^/]*>.*?</c>'
    return re.sub(p_content, new_cell, sheet_xml, flags=re.DOTALL)

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

        # Read ALL original styles before any modification
        styles = {}
        for row in REPAIR_ROWS + PARTS_ROWS:
            styles[f'B{row}'] = get_style(sheet_xml, f'B{row}')
        for row in PARTS_ROWS:
            styles[f'I{row}'] = get_style(sheet_xml, f'I{row}')
        for ref in ['B7','D7','C13','F13','I13','C14','F14','I14','B113']:
            styles[ref] = get_style(sheet_xml, ref)

        # Parse shared strings
        ss_items = re.findall(r'<si>.*?</si>', shared_xml, re.DOTALL)
        ss_list = []
        for item in ss_items:
            m = re.search(r'<t[^>]*>(.*?)</t>', item, re.DOTALL)
            ss_list.append(m.group(1) if m else '')

        def ss(text):
            safe = str(text).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
            for i, s in enumerate(ss_list):
                if s == safe: return i
            ss_list.append(safe)
            return len(ss_list) - 1

        def text_cell(ref, val):
            return f'<c r="{ref}" s="{styles[ref]}" t="s"><v>{ss(val)}</v></c>'

        def num_cell(ref, val):
            return f'<c r="{ref}" s="{styles[ref]}"><v>{val}</v></c>'

        def empty_cell(ref):
            return f'<c r="{ref}" s="{styles[ref]}"/>'

        now = datetime.now()
        pat = (data.get('patente') or '').upper().strip()
        fecha = f" Fecha {now.day} de {MESES[now.month-1]} {now.year}"

        # Fill header cells
        sheet_xml = replace_cell(sheet_xml, 'B7',  text_cell('B7',  f'Presupuesto {pat}'))
        sheet_xml = replace_cell(sheet_xml, 'D7',  text_cell('D7',  fecha))
        sheet_xml = replace_cell(sheet_xml, 'C13', text_cell('C13', data.get('marca') or ''))
        sheet_xml = replace_cell(sheet_xml, 'F13', text_cell('F13', data.get('modelo') or ''))
        sheet_xml = replace_cell(sheet_xml, 'C14', text_cell('C14', pat))
        if data.get('anio'):
            try: sheet_xml = replace_cell(sheet_xml,'I13',num_cell('I13',int(data['anio'])))
            except: pass
        if data.get('km'):
            try: sheet_xml = replace_cell(sheet_xml,'F14',num_cell('F14',int(data['km'])))
            except: pass
        if data.get('combustible'):
            try: sheet_xml = replace_cell(sheet_xml,'I14',num_cell('I14',int(data['combustible'])/4))
            except: pass

        # Clear all variable rows (preserving original styles exactly)
        for row in REPAIR_ROWS:
            sheet_xml = replace_cell(sheet_xml, f'B{row}', empty_cell(f'B{row}'))
        for row in PARTS_ROWS:
            sheet_xml = replace_cell(sheet_xml, f'B{row}', empty_cell(f'B{row}'))
            sheet_xml = replace_cell(sheet_xml, f'I{row}', empty_cell(f'I{row}'))

        # Fill repairs
        trabajos = data.get('trabajos') or []
        for i, t in enumerate(trabajos):
            if i < len(REPAIR_ROWS) and t:
                row = REPAIR_ROWS[i]
                sheet_xml = replace_cell(sheet_xml, f'B{row}', text_cell(f'B{row}', t))

        # Fill parts
        idx_gama = ss('GAMA')
        repuestos = data.get('repuestos') or []
        for i, r in enumerate(repuestos):
            if i < len(PARTS_ROWS) and r:
                row = PARTS_ROWS[i]
                sheet_xml = replace_cell(sheet_xml, f'B{row}', text_cell(f'B{row}', r))
                sheet_xml = replace_cell(sheet_xml, f'I{row}',
                    f'<c r="I{row}" s="{styles[f"I{row}"]}" t="s"><v>{idx_gama}</v></c>')

        # Observations
        sheet_xml = replace_cell(sheet_xml, 'B113', text_cell('B113', data.get('observaciones') or ''))

        # Rebuild shared strings
        new_ss = ''.join(f'<si><t xml:space="preserve">{s}</t></si>' for s in ss_list)
        new_shared = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'count="{len(ss_list)}" uniqueCount="{len(ss_list)}">'
            f'{new_ss}</sst>'
        )

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
            download_name=f'{pat}GAMA.xlsx')

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)

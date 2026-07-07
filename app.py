import os
import io
import base64
import tempfile
import sys
from flask import Flask, render_template, request, jsonify, Response

# Force reload of idoc_parser to pick up latest changes
if 'idoc_parser' in sys.modules:
    del sys.modules['idoc_parser']

from idoc_parser import parse_flat, build_excel, build_preview_data

app = Flask(__name__, static_folder='static', static_url_path='/static')
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32 MB
app.config['JSON_SORT_KEYS'] = False  # Force fresh cache

# Cache busting version
APP_VERSION = "20260707-145000"

ALLOWED_EXT = {'.txt', '.idoc', '.xml'}

print(f"DEBUG: app.py loaded from: {__file__}")
print(f"DEBUG: ALLOWED_EXT = {ALLOWED_EXT}")

def allowed(filename):
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXT


@app.route('/')
def index():
    from flask import make_response
    html = render_template('index.html')
    response = make_response(html)
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response


@app.route('/debug')
def debug_info():
    return jsonify({
        'app_file': __file__,
        'allowed_ext': list(ALLOWED_EXT),
        'debug_mode': app.debug
    })


@app.route('/convert', methods=['POST'])
def convert():
    if 'idoc_file' not in request.files:
        return jsonify(error='Nie przesłano pliku.'), 400

    f = request.files['idoc_file']
    if not f.filename:
        return jsonify(error='Nie wybrano pliku.'), 400
    if not allowed(f.filename):
        return jsonify(error='Dozwolone rozszerzenia: .txt, .idoc, .xml'), 400

    try:
        with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as tmp:
            f.save(tmp.name)
            tmp_path = tmp.name

        rows = parse_flat(tmp_path)

        # DEBUG: Check E1EDK14 data
        for seg_name, hlevel, data in rows:
            if seg_name == 'E1EDK14':
                print(f"DEBUG E1EDK14: data_len={len(data)} first 40 chars={repr(data[:40])}")
                break

        if not rows:
            os.unlink(tmp_path)
            return jsonify(error='Plik jest pusty lub nie zawiera danych IDoc.'), 400

        buf = io.BytesIO()
        build_excel(rows, buf)
        buf.seek(0)
        xlsx_b64 = base64.b64encode(buf.read()).decode('ascii')

        preview = build_preview_data(rows)
        os.unlink(tmp_path)

        base_name = os.path.splitext(f.filename)[0]
        return jsonify(
            filename=f"{base_name}_dokumentacja.xlsx",
            xlsx_b64=xlsx_b64,
            **preview
        )

    except Exception as e:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
        return jsonify(error=f'Błąd przetwarzania: {e}'), 500


if __name__ == '__main__':
    app.run(debug=False, port=5004)

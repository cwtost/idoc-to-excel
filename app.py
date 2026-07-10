import os
import io
import base64
import tempfile
import sys
from flask import Flask, render_template, request, jsonify, Response
import importlib

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
    with open('uploads/endpoint_called.txt', 'w') as df:
        df.write('convert endpoint was called\n')

    if 'idoc_file' not in request.files:
        return jsonify(error='Nie przesłano pliku.'), 400

    f = request.files['idoc_file']
    if not f.filename:
        return jsonify(error='Nie wybrano pliku.'), 400
    if not allowed(f.filename):
        return jsonify(error='Dozwolone rozszerzenia: .txt, .idoc, .xml'), 400

    try:
        # Dynamic import with reload to pick up latest changes
        import idoc_parser
        importlib.reload(idoc_parser)

        # Save to uploads directory
        os.makedirs('uploads', exist_ok=True)
        tmp_path = os.path.join('uploads', secure_filename(f.filename) + '.tmp')

        with open('uploads/debug.txt', 'w') as df:
            df.write(f"About to save to {tmp_path}\n")

        f.save(tmp_path)

        with open('uploads/debug.txt', 'a') as df:
            df.write(f"Saved successfully\n")

        rows = idoc_parser.parse_flat(tmp_path)

        # DEBUG: Show first few rows
        print(f"\n=== PARSER DEBUG ===")
        print(f"Total rows: {len(rows)}")
        if rows:
            print(f"First 3 rows:")
            for i, (seg, hlevel, data) in enumerate(rows[:3]):
                print(f"  {i+1}. seg='{seg}' hlevel={hlevel} data_len={len(data)}")
                fields = extract_fields(seg, data)
                print(f"     fields: {len(fields)} field(s)")
                if fields:
                    fname, flen, fdesc, val = fields[0]
                    print(f"     first field: {fname}={repr(val)}")
        print(f"==================\n")

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

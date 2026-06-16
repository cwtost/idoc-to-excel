import os
import io
import base64
import tempfile
from flask import Flask, render_template, request, jsonify
from idoc_parser import parse_flat, build_excel, build_preview_data

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32 MB

ALLOWED_EXT = {'.txt', '.idoc'}


def allowed(filename):
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXT


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert():
    if 'idoc_file' not in request.files:
        return jsonify(error='Nie przesłano pliku.'), 400

    f = request.files['idoc_file']
    if not f.filename:
        return jsonify(error='Nie wybrano pliku.'), 400
    if not allowed(f.filename):
        return jsonify(error='Dozwolone rozszerzenia: .txt, .idoc'), 400

    try:
        with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as tmp:
            f.save(tmp.name)
            tmp_path = tmp.name

        rows = parse_flat(tmp_path)
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
    app.run(debug=True, port=5001)

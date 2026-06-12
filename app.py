import os
import io
import tempfile
from flask import Flask, render_template, request, send_file, jsonify
from idoc_parser import parse_flat, build_excel

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
        # Save upload to a temp file (parse_flat needs a path)
        with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as tmp:
            f.save(tmp.name)
            tmp_path = tmp.name

        rows = parse_flat(tmp_path)
        if not rows:
            os.unlink(tmp_path)
            return jsonify(error='Plik jest pusty lub nie zawiera danych IDoc.'), 400

        # Build Excel into a BytesIO buffer
        buf = io.BytesIO()
        build_excel(rows, buf)
        buf.seek(0)

        os.unlink(tmp_path)

        base = os.path.splitext(f.filename)[0]
        xlsx_name = f"{base}_dokumentacja.xlsx"
        return send_file(
            buf,
            as_attachment=True,
            download_name=xlsx_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )

    except Exception as e:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
        return jsonify(error=f'Błąd przetwarzania: {e}'), 500


if __name__ == '__main__':
    app.run(debug=True, port=5001)

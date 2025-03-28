from flask import Flask, request, send_file, jsonify
from openpyxl import load_workbook
import tempfile
import os

app = Flask(__name__)

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return 'No file part', 400

    file = request.files['file']
    if file.filename == '':
        return 'No selected file', 400

    # Save uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        file.save(tmp.name)
        uploaded_path = tmp.name

    # Mock transformation logic
    try:
        # Load uploaded file and extract user inputs (this is where you’d extract real values)
        wb_old = load_workbook(uploaded_path)
        ws_old = wb_old.active
        user_name = ws_old['B2'].value if 'B2' in ws_old else 'Anonymous'

        # Load blank template to populate
        template_path = 'files/template.xlsx'  # pre-saved new version
        wb_new = load_workbook(template_path)
        ws_new = wb_new.active

        # Insert user data (mock)
        ws_new['B2'] = user_name  # Assuming B2 is where user's name goes in the new version

        # Save output file
        output_path = tempfile.mktemp(suffix='.xlsx')
        wb_new.save(output_path)

        return send_file(output_path,
                         as_attachment=True,
                         download_name='updated_template.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return f'Error processing file: {str(e)}', 500
    finally:
        os.unlink(uploaded_path)

if __name__ == '__main__':
    app.run(debug=True)

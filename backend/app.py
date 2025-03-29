from flask import Flask, request, send_file, jsonify
from openpyxl import load_workbook
import tempfile
import os
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

@app.route('/upload', methods=['POST']) # creates endpoint for file upload
def upload():

    print('DEVNOTE endpoint hit')

    #check if the file was sent correctly
    if 'file' not in request.files: # grabs file from the request via the key 'file'
        return 'No file part', 400

    file = request.files['file']
    if file.filename == '':
        return 'No selected file', 400

    # Save uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsm') as tmp:
        file.save(tmp.name)
        uploaded_path = tmp.name

    print('DEVNOTE temp file saved')

    # TRANSFORMATION LOGIC - currently a mockup
    try:

        print('DEVNOTE transformation logic started')

        # Load uploaded file and extract user inputs
        wb_old = load_workbook(uploaded_path)
        ws_old = wb_old.active
        user_name = ws_old['A1'].value # if 'A1' in ws_old else 'Anonymous'

        # Load blank template to populate
        template_path = next((f for f in os.listdir('.') if f.endswith('.xlsm')), None)
        if not template_path:
            raise FileNotFoundError("No .xlsm template file found in the current directory.")
        wb_new = load_workbook(template_path, keep_vba=True)
        ws_new = wb_new.active

        print('DEVNOTE template loaded')

        # Insert user data (mock)
        ws_new['A1'] = user_name  # Assuming B2 is where user's name goes in the new version

        # Migrate assets data
        migrate_assets_tab(wb_old, wb_new)

        # Save output file
        output_path = tempfile.mktemp(suffix='.xlsm')
        wb_new.save(output_path)

        print('DEVNOTE output file saved')

        return send_file(output_path,
                         as_attachment=True,
                         download_name='updated_template.xlsm',
                         mimetype='application/vnd.ms-excel.sheet.macroEnabled.12')
    except Exception as e:
        return f'Error processing file: {str(e)}', 500
    finally:
        os.unlink(uploaded_path)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))  # fallback to 5000 for local dev
    app.run(host='0.0.0.0', port=port)

def migrate_assets_tab(wb_old, wb_new):

    try:
        # Locate the 'Assets' tab in the old workbook
        assets_sheet_old = None
        for sheet in wb_old.worksheets:
            if sheet.title == 'Assets':
                assets_sheet_old = sheet
                break
        if not assets_sheet_old:
            raise ValueError("No 'Assets' tab found in the uploaded workbook.")

        print('DEVNOTE Assets tab found')
        
        # Locate the 'Assets' tab in the new workbook
        assets_sheet_new = None
        for sheet in wb_new.worksheets:
            if sheet.title == 'Assets':
                assets_sheet_new = sheet
                break
        if not assets_sheet_new:
            raise ValueError("No 'Assets' tab found in the new workbook.")

        print('DEVNOTE Assets tab found in new workbook')

        # Example: Copy data from the first 10 rows and 5 columns
        for row in range(3, 99):  # Adjust range as needed
            for col in range(2, 5):  # Adjust range as needed
                assets_sheet_new.cell(row=row, column=col).value = assets_sheet_old.cell(row=row, column=col).value

        print('DEVNOTE Assets data migrated successfully')

    except Exception as e:
        print(f"Error migrating assets data: {str(e)}")
        raise
from flask import Flask, request, send_file, jsonify
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
import tempfile
import os
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

def locate_sheet(workbook, sheet_name):
    for sheet in workbook.worksheets:
        if sheet.title == sheet_name:
            return sheet
    raise ValueError(f"No '{sheet_name}' tab found in {workbook}.")

def copy_data(source_sheet, target_sheet, start_row, end_row, start_col, end_col):
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            old_cell = source_sheet.cell(row=row, column=col)
            new_cell = target_sheet.cell(row=row, column=col)
            new_cell.value = old_cell.value
            new_cell.number_format = old_cell.number_format

@app.route('/upload', methods=['POST']) # creates endpoint for file upload
def upload():

    print('DEVNOTE endpoint hit')

    #--LOAD WORKBOOK--

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

    wb_old = load_workbook(uploaded_path)
    template_path = next((f for f in os.listdir('.') if f.endswith('.xlsm')), None) # loads the template file via extention search
    if not template_path:
        raise FileNotFoundError("No .xlsm template file found in the current directory.")
    
    wb_new = load_workbook(template_path, keep_vba=True)
        
    print('DEVNOTE template loaded')

    #--MIGRATION LOGIC--

    try:
        print('DEVNOTE transformation logic started')

        assets_sheet_old = locate_sheet(wb_old, 'Assets')
        assets_sheet_new = locate_sheet(wb_new, 'Assets')

        print('DEVNOTE copying Assets data')

        copy_data(assets_sheet_old, assets_sheet_new, 3, 99, 2, 5)

        print('DEVNOTE applying data validation')

        # Define the data validation rule
        data_validation = DataValidation(
            type="list",
            formula1="'Data Validation'!$A$2:$A$99",  # Reference to the source range
            allow_blank=True,
            showDropDown=False,
            showErrorMessage=True
        )

        # Add the data validation to the range B4:B99
        data_validation.add("B4:B99")

        # Apply the data validation to the new sheet
        assets_sheet_new.add_data_validation(data_validation)

        print('DEVNOTE Assets data migrated successfully')

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
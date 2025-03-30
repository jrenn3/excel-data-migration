from flask import Flask, request, send_file, jsonify
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
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
            
            # Check for the custom named formula and replace it
            if isinstance(old_cell.value, ArrayFormula) and old_cell.value.text == '=EndDayOfCurrentMonth':
                new_cell.value = '=EndOfCurrentMonth'
            else:
                new_cell.value = old_cell.value
                        
            new_cell.number_format = old_cell.number_format # copy old number format over

def apply_data_validation(sheet, validation_range, source_range):
    data_validation = DataValidation(
        type="list",
        formula1=source_range,
        allow_blank=True,
        showDropDown=False,
        showErrorMessage=True
    )
    data_validation.add(validation_range)
    sheet.add_data_validation(data_validation)

def migrate_sheet(old_workbook, new_workbook, target_sheet, start_row, end_row, start_col, end_col): #todo: add: , validation_range=None, source_range=None

    ws_old = locate_sheet(old_workbook, target_sheet)
    ws_new = locate_sheet(new_workbook, target_sheet)

    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            old_cell = ws_old.cell(row=row, column=col)
            new_cell = ws_new.cell(row=row, column=col)
            
            # Check for the custom named formula and replace it
            if isinstance(old_cell.value, ArrayFormula) and old_cell.value.text == '=EndDayOfCurrentMonth':
                new_cell.value = '=EndOfCurrentMonth'
            else:
                new_cell.value = old_cell.value
                        
            new_cell.number_format = old_cell.number_format # copy old number format over

    if target_sheet == 'Assets':
        # Apply data validation for the Assets tab
        apply_data_validation(
            ws_new,
            "$B$4:$B$99",
            "'Data Validation'!$A$2:$A$99" # todo define these and name each range
        )


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

        #--ASSETS TAB--
        migrate_sheet(wb_old, wb_new, 'Assets', 3, 99, 2, 5)        

        #--CREDIT CARDS TAB--        
        migrate_sheet(wb_old, wb_new, 'Credit Cards', 3, 24, 2, 6)

        #--LOANS TAB--
        ws_old = locate_sheet(wb_old, 'Loans')
        ws_new = locate_sheet(wb_new, 'Loans')
        copy_data(ws_old, ws_new, 3, 99, 2, 3)

        #--LOYALTY POINTS & MILES TAB--
        ws_old = locate_sheet(wb_old, 'Loyalty Points & Miles')
        ws_new = locate_sheet(wb_new, 'Loyalty Points & Miles')
        copy_data(ws_old, ws_new, 3, 53, 2, 6)
        
        #--RECURRING TAB--
        ws_old = locate_sheet(wb_old, 'Recurring')
        ws_new = locate_sheet(wb_new, 'Recurring')
        copy_data(ws_old, ws_new, 3, 53, 2, 7)
        #todo: data validation

        #--PRECEDENTS TAB--
        ws_old = locate_sheet(wb_old, 'Precedents')
        ws_new = locate_sheet(wb_new, 'Precedents')
        copy_data(ws_old, ws_new, 3, 99, 2, 5)
        #todo: data validation

        #--CHANGES TAB--
        ws_old = locate_sheet(wb_old, 'Changes')
        ws_new = locate_sheet(wb_new, 'Changes')
        copy_data(ws_old, ws_new, 3, 23, 2, 7)
        #todo: data validation

        #--PLANNED TAB--
        ws_old = locate_sheet(wb_old, 'Planned')
        ws_new = locate_sheet(wb_new, 'Planned')
        copy_data(ws_old, ws_new, 3, 24, 2, 6)
        #todo: data validation

        #--todo: BLANKET TAB--

        print('DEVNOTE data migrated successfully')

        #--SAVE AND DELIVER FILE--
        output_path = tempfile.mktemp(suffix='.xlsm')
        wb_new.save(output_path)

        print('DEVNOTE output file saved')

        return send_file(output_path,
                         as_attachment=True,
                         mimetype='application/vnd.ms-excel.sheet.macroEnabled.12')
    
    except Exception as e:
        return f'Error processing file: {str(e)}', 500
    
    finally:
        os.unlink(uploaded_path)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))  # fallback to 5000 for local dev
    app.run(host='0.0.0.0', port=port)
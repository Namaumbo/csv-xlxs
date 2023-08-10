import io
import pathlib
from flask import Flask, Response, request, jsonify, send_file
import openpyxl
import pandas as pd
import os
import pythoncom
import win32com
from win32com.client.gencache import EnsureDispatch




app = Flask(__name__)
def set_wb_password_with_win(file_path, password):
    xl_file= win32com.client.Dispatch("Excel.Application",pythoncom.CoInitialize())
    wb = xl_file.Workbooks.Open(file_path)
    xl_file.DisplayAlerts = False
    wb.Visible = False
    wb.SaveAs(file_path , Password=password)
    wb.Close()
    xl_file.Quit()


@app.route('/api/v1/convert' , methods=['POST'])
def convert():
    try:

        uploaded_file = request.files['file_name']
        file_name = uploaded_file.filename
        file = file_name.split('.csv')[0]
        exc_file_name = f"""{file}.xlsx""" 

        if uploaded_file.filename != '':
            # Read CSV data
            csv_data = uploaded_file.read()
            csv_io = io.StringIO(csv_data.decode('utf-8'))
            df = pd.read_csv(csv_io)

            # Convert to XLSX
            xlsx_io = io.BytesIO()
            password = '123'
            df.to_excel(xlsx_io, index=False)
            set_wb_password_with_win(xlsx_io, password)

            xlsx_io.seek(0)


            # Send XLSX file back to user
        return send_file(
                xlsx_io,
                download_name=f"""{exc_file_name}""",
                as_attachment=True,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        print(f'ERROR: {e} on line {e.__traceback__.tb_lineno}')
    return 'error'
if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True)


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

def unprotect_wb_with_win32 (file_path ,new_file_path , password):
    xl_file = EnsureDispatch("Excel.Application")
    wb = xl_file.Workbooks.open(file_path,False,True, None, password)
    wb.Unprotect(password)
    xl_file.DisplayAlerts = False
    wb.Visiblae = False
    wb.SaveAs(new_file_path, None ,'','')
    wb.Close()
    xl_file.quit()

       
@app.route('/api/v1/convert' , methods=['POST'])
def convert():
    # try :
        # response = {
        #     'message':'',
        #     'status':''
        # }
        # code =200
        # uploaded_file = request.files['file_name']
        # file_name = uploaded_file.filename
        # file = file_name.split('.csv')[0]
        # exc_file_name = f"""{file}.xlsx""" 
        
        # #name of the file
        # file_name = uploaded_file.filename
        # if file_name != '':
        #     csv_data = uploaded_file.read()
        #     csv_io = io.StringIO(csv_data.decode('utf-8'))

        #     df = pd.read_csv(csv_io , header=0)
        #     xlsx_io = io.BytesIO()
        #     df.to_excel(xlsx_io, index=False)
        #     xlsx_io.seek(0)
        #     password = '123'
        #     set_wb_password_with_win(str(xlsx_io), password)

        #     return send_file(
        #     xlsx_io,
        #     download_name=f"""{exc_file_name}""",
        #     as_attachment=True,
        #     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        # )

        # if os.path.exists(csv_path):
        #     # directory_path = './excel_files'
        #     csv_file_name = pathlib.Path(csv_path).name
        #     file = csv_file_name.split('.csv')[0]
        #     exc_file_name = f"""{file}.xlsx""" 


            # exc_file_path = os.path.join(directory_path, exc_file_name)

            # if os.path.exists(csv_path):
                # 
    
                # if not os.path.exists(exc_file_path):  
                #     df.to_excel(exc_file_path, index=False) 
                #     full_path = os.path.abspath(exc_file_path)
                #     full = pathlib.PurePath(full_path)

                    # change password here
                #     password = '123'

                #     if os.path.exists(full_path):
                #         set_wb_password_with_win(str(full), password)

                #     output = io.BytesIO()    
                #     response = Response(output.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                #     response.headers['Content-Disposition'] = 'attachment; filename=data.xlsx'
                #     response['message'] = 'Done Processing 😉'
                #     response['status'] = 'success'
                #     code = 200  
                # else:
                #     response['message']='file already exists'
                #     response['status'] = 'fail'
                #     code : 409     
        # else:
        #     response['message'] = 'File processing failed'

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
    
if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True)


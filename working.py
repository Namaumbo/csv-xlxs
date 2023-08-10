import pathlib
from flask import Flask, request, jsonify
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

       

@app.route('/api/v1/convert' , methods=['GET'])
def convert():
    try :
        response = {
            'message':'',
            'status':''
        }
        code =200
        uploaded_file = request.files['file_name']
        file_name = uploaded_file.filename
        files = r'./files'
        csv_path = os.path.join(files, file_name)

        if uploaded_file:
            uploaded_file.save(str(csv_path))

        if os.path.exists(csv_path):
            directory_path = './excel_files'
            csv_file_name = pathlib.Path(csv_path).name
            file = csv_file_name.split('.csv')[0]
            exc_file_name = f"""{file}.xlsx""" 
            exc_file_path = os.path.join(directory_path, exc_file_name)

            if os.path.exists(csv_path):
                df = pd.read_csv(csv_path , header=0)
                if not os.path.exists(directory_path):
                    os.makedirs(directory_path)

                if not os.path.exists(exc_file_path):  
                    df.to_excel(exc_file_path, index=False) 
                    full_path = os.path.abspath(exc_file_path)
                    full = pathlib.PurePath(full_path)

                    # change password here
                    password = '123'

                    if os.path.exists(full_path):
                        set_wb_password_with_win(str(full), password)
                
                    response['message'] = 'Done Processing'
                    response['status'] = 'success'
                    code = 200  
                else:
                    response['message']='file already exists'
                    response['status'] = 'fail'
                    code : 409     
        else:
            response['message'] = 'File processing failed'
    except Exception as e :
        response = {'error': f"""{e}"""}
        response['status'] = 'fail'
        code : 500

    return jsonify(response) , code

if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True)


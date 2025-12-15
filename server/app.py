from flask import Flask, request
from flask_cors import CORS
import os
import pandas as pd

app = Flask(__name__)
CORS(app)

EXCEL_DIR = os.path.join(os.path.dirname(__file__), 'excel-files')

@app.route('/api/excel-files', methods=['GET'])

def get_excel_files():

    files = [
        {"id": os.path.splitext(f) [0].lower(), "name": f}
        for f in os.listdir(EXCEL_DIR)
        if f.endswith('.xlsx') or f.endswith('.xls')
    ]
    print("files are")
    print({"files": files})
    return {"files": files}


import math

@app.route('/api/search-excel', methods=['POST'])

def search_excel():

    params = request.json

    file_name = params.get('fileName')
    search_params = {
    "fieldName": params.get('fieldName','').strip().lower(), 
    "fieldType": params.get('fieldType', '').strip().lower(), 
    "visibilityRules": params.get('visibilityRules', '').strip().lower(),
    "visibilityAttributes": params.get("visibilityAttributes", '').strip().lower()
    }
    column_mapping= {

    'fieldName': ['Field Name', 'FieldName', 'Field Name', 'Name'],
    'description': ['Description'],
    'fieldType': ['Field Type', 'FieldType', 'Type', 'DataType'],
    'format': ['Format'],
    'fieldLength': ['Field Length', 'FieldLength', 'Length'],
    'defaultValue': ['Default Value', 'DefaultValue', 'Default'],
    'validValues': ['Valid Values', 'Valid Value(s)', 'ValidValues'],
    'fieldBehaviour': ['Field Behaviour', 'Field Behavior', 'FieldBehaviour', 'Behavior', 'Behaviour'],
    'visibilityRules': ['Visibility Rules', 'VisibilityRules', 'Rules'],
    'visibilityAttributes': ['Visibility Attributes', 'VisibilityAttributes', 'Attributes']
    }

    results= []


    if file_name:

        files_to_search = [
            f for f in os.listdir(EXCEL_DIR)
            if os.path.splitext(f)[0].lower() == file_name.lower()
        ]

    else:
        files_to_search = [
            f for f in os.listdir(EXCEL_DIR)
            if f.endswith('.xlsx') or f.endswith('.xls')    
        ]
    for file in files_to_search:
        file_path = os.path.join(EXCEL_DIR, file)
        try:
            df = pd.read_excel(file_path)
            normalized_cols = {c.lower().replace(" ", "").replace("_", ""): c for c in df.columns}
            search_field_mappings = {}

            for key, possible_headers in column_mapping.items():
                for header in possible_headers:
                    norm_header = header.lower().replace(" ", "").replace("_", "")
                    if norm_header in normalized_cols:
                        search_field_mappings[key] = normalized_cols[norm_header]
                        break

            for _, row in df.iterrows():
                # If all search fields are empty, include all rows
                if all(not v for v in search_params.values()):
                    match = True
                else:
                    match = True
                    for key, val in search_params.items():
                        if val:
                            excel_col = search_field_mappings.get(key)
                            if excel_col:
                                cell = str(row[excel_col]).lower() if not pd.isna(row[excel_col]) else ''
                                if val not in cell:
                                    match = False
                                    break
                if match:
                    result ={}
                    for camel_key, excel_col in search_field_mappings.items():
                        value = row[excel_col] if excel_col in row else ""
                        if isinstance(value, float) and math.isnan(value):
                            value = ""
                        result[camel_key] = value
                    result['sourceFile'] = file
                    results.append(result)
        except Exception as e:
            print(f"Error processing file {file}: {e}")
    return {"results": results}

if __name__ == '__main__':
    if not os.path.exists(EXCEL_DIR):
        os.makedirs(EXCEL_DIR)
    app.run(port=3001, debug=True)
    """Endpoint to retrieve list of available Excel files."""
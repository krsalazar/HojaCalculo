from flask import Flask, jsonify, request, render_template
from marshmallow import Schema, fields, ValidationError
from asteval import Interpreter
import re

app = Flask(__name__)

# Modelo
class Cell:
    def __init__(self, value=''):
        self.value = value

class SpreadsheetModel:
    def __init__(self, rows, cols):
        self.data = [[Cell() for _ in range(cols)] for _ in range(rows)]
        self.asteval = Interpreter()

    def add_row(self):
        cols = len(self.data[0])
        self.data.append([Cell() for _ in range(cols)])

    def add_col(self):
        for row in self.data:
            row.append(Cell())

    def set_value(self, row, col, value):
        self.data[row][col].value = value

    def get_value(self, row, col):
        return self.data[row][col].value

    def evaluate_formula(self, formula):
        # Expresión regular para encontrar referencias de celda. (e.g., A1, B2, etc.)
        cell_ref_pattern = re.compile(r'([A-Z]+)([0-9]+)')
        
        def cell_ref_to_value(match):
            col_str, row_str = match.groups()
            row = int(row_str) - 1
            col = sum((ord(char) - 65 + 1) for char in col_str) - 1
            return self.get_value(row, col)

        try:
            # Reemplazar referencias de celda en la fórmula con valores reales
            formula_with_values = cell_ref_pattern.sub(
                lambda match: str(cell_ref_to_value(match)), formula)
            return self.asteval(formula_with_values)
        except Exception as e:
            return str(e)

class Workbook:
    def __init__(self):
        self.sheets = {}
        self.active_sheet = None

    def add_sheet(self, name, rows, cols):
        self.sheets[name] = SpreadsheetModel(rows, cols)
        if self.active_sheet is None:
            self.active_sheet = name

    def get_sheet(self, name):
        return self.sheets.get(name)

    def set_active_sheet(self, name):
        if name in self.sheets:
            self.active_sheet = name

    def rename_sheet(self, old_name, new_name):
        if old_name in self.sheets and new_name not in self.sheets:
            self.sheets[new_name] = self.sheets.pop(old_name)
            if self.active_sheet == old_name:
                self.active_sheet = new_name

workbook = Workbook()
workbook.add_sheet("Sheet1", 20, 26)

# Esquemas de validación
class SetValueSchema(Schema):
    sheet_name = fields.String(required=True)
    row = fields.Integer(required=True)
    col = fields.Integer(required=True)
    value = fields.String(required=True)

set_value_schema = SetValueSchema()

@app.route('/get_sheet', methods=['GET'])
def get_sheet():
    sheet_name = request.args.get('sheet_name', workbook.active_sheet)
    sheet = workbook.get_sheet(sheet_name)
    data = [[cell.value for cell in row] for row in sheet.data]
    return jsonify(data)

@app.route('/get_sheets', methods=['GET'])
def get_sheets():
    sheets = list(workbook.sheets.keys())
    return jsonify(sheets)

@app.route('/set_value', methods=['POST'])
def set_value():
    try:
        data = set_value_schema.load(request.json)
    except ValidationError as err:
        return jsonify(err.messages), 400
    
    sheet_name = data['sheet_name']
    row = data['row']
    col = data['col']
    value = data['value']
    sheet = workbook.get_sheet(sheet_name)
    if value.startswith('='):
        value = sheet.evaluate_formula(value[1:])
    sheet.set_value(row, col, value)
    return jsonify(success=True)

@app.route('/add_sheet', methods=['POST'])
def add_sheet():
    name = request.json['name']
    rows = request.json['rows']
    cols = request.json['cols']
    workbook.add_sheet(name, rows, cols)
    return jsonify(success=True)

@app.route('/rename_sheet', methods=['POST'])
def rename_sheet():
    old_name = request.json['old_name']
    new_name = request.json['new_name']
    workbook.rename_sheet(old_name, new_name)
    return jsonify(success=True)

@app.route('/add_row', methods=['POST'])
def add_row():
    sheet_name = request.json['sheet_name']
    sheet = workbook.get_sheet(sheet_name)
    sheet.add_row()
    return jsonify(success=True)

@app.route('/add_col', methods=['POST'])
def add_col():
    sheet_name = request.json['sheet_name']
    sheet = workbook.get_sheet(sheet_name)
    sheet.add_col()
    return jsonify(success=True)

@app.route('/set_active_sheet', methods=['POST'])
def set_active_sheet():
    sheet_name = request.json['sheet_name']
    workbook.set_active_sheet(sheet_name)
    return jsonify(success=True)

@app.route('/')
def index():
    return render_template('index.html')

if __name__ == "__main__":
    app.run(debug=True)
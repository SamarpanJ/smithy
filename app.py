import os
from flask import Flask, jsonify, render_template, request
from openpyxl import load_workbook
from datetime import datetime
import time

app = Flask(__name__)

COLOR_CODES = {
    'Red': 'FFFF0000',
    'Blue': 'FF00B0F0',
    'Green': 'FF00B050',
    'Normal': None
}

def get_directory():
    """Prompt for directory path, defaulting to desktop."""
    directory = os.environ.get('MYEXCEL_DIRECTORY')
    
    if not directory:
        desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        directory = input(f"Enter the directory path where 'myexcel.xlsx' is located (default: {desktop_path}): ").strip()
        if not directory:
            directory = desktop_path
        
        os.environ['MYEXCEL_DIRECTORY'] = directory

    return directory

def get_file_path(directory):
    """Get the full path to 'myexcel.xlsx' in the specified directory."""
    return os.path.join(directory, 'myexcel.xlsx')

directory = get_directory()

def get_cell_color(cell):
    """Extract color information from a cell."""
    if cell.font and cell.font.color and cell.font.color.type == 'rgb':
        return cell.font.color.rgb
    return None

def read_excel_with_colors(file_path):
    """Read Excel file and categorize cell colors."""
    wb = load_workbook(filename=file_path, data_only=True)
    sheet = wb.active

    categories = {
        'Red': [],
        'Blue': [],
        'Green': [],
        'Normal': []
    }

    for row in sheet.iter_rows(min_row=4, values_only=False):  # Start from row 4
        cell = row[3]  # Column D (0-based index 3)
        color = get_cell_color(cell)
        cell_data = [cell.value for cell in row]

        if color:
            if color == COLOR_CODES['Red']:
                categories['Red'].append(cell_data)
            elif color == COLOR_CODES['Blue']:
                categories['Blue'].append(cell_data)
            elif color == COLOR_CODES['Green']:
                categories['Green'].append(cell_data)
            else:
                categories['Normal'].append(cell_data)
        else:
            categories['Normal'].append(cell_data)

    return categories

def rotate_entries(entries, n=2, offset=0):
    """Rotate entries and return a subset."""
    if not entries:
        return []
    length = len(entries)
    return [entries[(i + offset) % length] for i in range(n)]

def parse_date(date_str):
    if isinstance(date_str, str):
        return datetime.strptime(date_str, '%d-%m-%Y').date()
    else:
        raise ValueError(f"Unsupported date type: {type(date_str)}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/data')
def data():
    file_path = get_file_path(directory)
    
    if not os.path.exists(file_path):
        return jsonify({'error': 'File not found'}), 404
    
    color_info = read_excel_with_colors(file_path)

    current_time = int(time.time())
    offset = (current_time // 10) % max(len(color_info['Red']), len(color_info['Blue']), len(color_info['Green']), 1)

    rotated_color_info = {
        'Red': rotate_entries(color_info['Red'], offset=offset),
        'Blue': rotate_entries(color_info['Blue'], offset=offset),
        'Green': rotate_entries(color_info['Green'], offset=offset),
    }

    return jsonify(rotated_color_info)

@app.route('/list')
def list_entries():
    file_path = get_file_path(directory)
    
    if not os.path.exists(file_path):
        return jsonify({'error': 'File not found'}), 404

    color_info = read_excel_with_colors(file_path)

    combined_list = []
    for category in ['Red', 'Blue', 'Green']:
        for entry in color_info[category]:
            entry_with_color = entry + [category]
            combined_list.append(entry_with_color)

    combined_list.sort(key=lambda x: parse_date(x[1]))

    list_length = len(combined_list)
    if list_length > 10:
        current_time = int(time.time())
        offset = (current_time // 5) % list_length

        rotated_list = combined_list[offset:offset + 13]

        if len(rotated_list) < 13:
            rotated_list += combined_list[:13 - len(rotated_list)]

        return jsonify(rotated_list)
    else:
        return jsonify(combined_list)

@app.route('/list2')
def list2data():
    return render_template("list2.html")

@app.route('/list1')
def list1data():
    return render_template('list1.html')

if __name__ == '__main__':
    app.run()

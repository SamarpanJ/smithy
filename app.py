from flask import Flask, jsonify, render_template, request
from openpyxl import load_workbook
from openpyxl.styles import Color
import time

app = Flask(__name__)

# Updated color codes
COLOR_CODES = {
    'Red': 'FFFF0000',
    'Blue': 'FF00B0F0',
    'Green': 'FF00B050',
    'Normal': None
}

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

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/data')
def data():
    # Load Excel data and colors
    color_info = read_excel_with_colors('data/samarpan.xlsx')
    
    # Get refresh interval from query parameter
    refresh_interval = int(request.args.get('interval', 10))  # Default to 5 seconds if not provided

    # Get the current time in seconds and use it to create an offset for rotation
    current_time = int(time.time())
    offset = (current_time // refresh_interval) % max(len(color_info['Red']), len(color_info['Blue']), len(color_info['Green']), 1)

    # Rotate entries
    rotated_color_info = {
        'Red': rotate_entries(color_info['Red'], offset=offset),
        'Blue': rotate_entries(color_info['Blue'], offset=offset),
        'Green': rotate_entries(color_info['Green'], offset=offset),
    }

    return jsonify(rotated_color_info)

if __name__ == '__main__':
    app.run(debug=True)

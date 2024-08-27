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

@app.route('/list')
def list_entries():
    # Load Excel data and colors
    color_info = read_excel_with_colors('data/samarpan.xlsx')

    # Combine all non-black entries while maintaining the original order
    combined_list = []

    # Iterate over each category and append entries with color information
    for category in ['Red', 'Blue', 'Green']:
        for entry in color_info[category]:
            entry_with_color = entry + [category]  # Add color as an additional field
            combined_list.append(entry_with_color)

    # Sort the combined list by the "number" column (assumed to be the first column)
    combined_list.sort(key=lambda x: x[0])  # Adjust the index if needed

    # Get refresh interval and number of entries per set from query parameters
    refresh_interval = int(request.args.get('interval', 10))  # Default to 10 seconds
    entries_per_set = 10  # Fixed number of entries to return

    # Calculate offset based on the current time and refresh interval
    current_time = int(time.time())
    offset = (current_time // refresh_interval) % len(combined_list)

    # Rotate the list and select the current set of entries
    rotated_list = combined_list[offset:offset + entries_per_set]

    # If the slice is smaller than the required set, wrap around the list
    if len(rotated_list) < entries_per_set:
        rotated_list += combined_list[:entries_per_set - len(rotated_list)]

    return jsonify(rotated_list)


@app.route('/listdata')
def sending_listdata():
    return render_template("list.html")

if __name__ == '__main__':
    app.run(debug=True)

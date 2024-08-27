from flask import Flask, jsonify, render_template, request
from openpyxl import load_workbook
from datetime import datetime, timedelta
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
    refresh_interval = int(request.args.get('interval', 10))  # Default to 10 seconds if not provided

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
def parse_date(date_str):
    if isinstance(date_str, str):
        print(date_str)
        try:
            return datetime.strptime(date_str, '%d-%m-%Y').date()
        except ValueError:
            raise ValueError(f"Unsupported date format: {date_str}")
    else:
        raise ValueError(f"Unsupported date type: {type(date_str)}")

@app.route('/list')
def list_entries():
    # Load Excel data and colors
    color_info = read_excel_with_colors('data/samarpan.xlsx')

    # Combine all entries across different categories
    combined_list = []
    for category in ['Red', 'Blue', 'Green']:
        for entry in color_info[category]:
            entry_with_color = entry + [category]  # Add color as an additional field
            combined_list.append(entry_with_color)

    # Get date filters from query parameters
    fromdate_str = request.args.get('fromdate')
    todate_str = request.args.get('todate')

    fromdate = parse_date(fromdate_str) if fromdate_str else None
    todate = parse_date(todate_str) if todate_str else None

    # Apply date filters only if they are provided
    if fromdate or todate:
        filtered_list = []
        for entry in combined_list:
            entry_date = parse_date(entry[1])  # Assuming the date is in the second column
            if entry_date:
                if fromdate and todate:
                    if fromdate <= entry_date <= todate:
                        filtered_list.append(entry)
                elif fromdate:
                    if fromdate <= entry_date:
                        filtered_list.append(entry)
                elif todate:
                    if entry_date <= todate:
                        filtered_list.append(entry)
        combined_list = filtered_list

    # Sort the combined list by the date column
    combined_list.sort(key=lambda x: parse_date(x[1]))

    # Get refresh interval and entries per set from query parameters
    refresh_interval = int(request.args.get('interval', 5))  # Default to 5 seconds
    entries_per_set = int(request.args.get('entries', 10))  # Default to 10 entries

    # Determine the length of the combined list
    list_length = len(combined_list)

    if list_length > 10:
        # Calculate offset based on the current time and refresh interval
        current_time = int(time.time())
        offset = (current_time // refresh_interval) % max(list_length, 1)

        # Rotate the list and select the current set of entries
        rotated_list = combined_list[offset:offset + entries_per_set]

        # If the slice is smaller than the required set, wrap around the list
        if len(rotated_list) < entries_per_set:
            rotated_list += combined_list[:entries_per_set - len(rotated_list)]

        return jsonify(rotated_list)
    else:
        # If there are 10 or fewer entries, just return them as is
        return jsonify(combined_list)





@app.route('/listdata')
def sending_listdata():
    return render_template("list.html")

if __name__ == '__main__':
    app.run(debug=True)

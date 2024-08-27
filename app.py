import tkinter as tk
from tkinter import filedialog
import os
from flask import Flask, jsonify, render_template, request
from openpyxl import load_workbook
from datetime import datetime
import time
from threading import Thread

# Flask app setup
app = Flask(__name__)

COLOR_CODES = {
    'Red': 'FFFF0000',
    'Blue': 'FF00B0F0',
    'Green': 'FF00B050',
    'Normal': None
}

directory = None  # Global variable to store directory path

def get_file_path():
    """Get the full path to 'myexcel.xlsx' in the specified directory."""
    if directory:
        return os.path.join(directory, 'myexcel.xlsx')
    return None

def get_cell_color(cell):
    """Extract color information from a cell."""
    if cell.font and cell.font.color and cell.font.color.type == 'rgb':
        return cell.font.color.rgb
    return None

def read_excel_with_colors(file_path):
    """Read Excel file and categorize cell colors."""
    if not file_path:
        return { 'Red': [], 'Blue': [], 'Green': [], 'Normal': [] }
    
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
        try:
            return datetime.strptime(date_str, '%d-%m-%Y').date()
        except ValueError:
            raise ValueError(f"Unsupported date format: {date_str}")
    else:
        raise ValueError(f"Unsupported date type: {type(date_str)}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/data')
def data():
    file_path = get_file_path()
    if not file_path or not os.path.exists(file_path):
        return jsonify({'error': 'File not found'}), 404
    
    color_info = read_excel_with_colors(file_path)
    
    refresh_interval = int(request.args.get('interval', 10))  # Default to 10 seconds if not provided
    current_time = int(time.time())
    offset = (current_time // refresh_interval) % max(len(color_info['Red']), len(color_info['Blue']), len(color_info['Green']), 1)

    rotated_color_info = {
        'Red': rotate_entries(color_info['Red'], offset=offset),
        'Blue': rotate_entries(color_info['Blue'], offset=offset),
        'Green': rotate_entries(color_info['Green'], offset=offset),
    }

    return jsonify(rotated_color_info)

@app.route('/list')
def list_entries():
    file_path = get_file_path()
    if not file_path or not os.path.exists(file_path):
        return jsonify({'error': 'File not found'}), 404
    
    color_info = read_excel_with_colors(file_path)

    combined_list = []
    for category in ['Red', 'Blue', 'Green']:
        for entry in color_info[category]:
            entry_with_color = entry + [category]
            combined_list.append(entry_with_color)

    fromdate_str = request.args.get('fromdate')
    todate_str = request.args.get('todate')

    fromdate = parse_date(fromdate_str) if fromdate_str else None
    todate = parse_date(todate_str) if todate_str else None

    if fromdate or todate:
        filtered_list = []
        for entry in combined_list:
            entry_date = parse_date(entry[1])
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

    combined_list.sort(key=lambda x: parse_date(x[1]))

    refresh_interval = int(request.args.get('interval', 5))  # Default to 5 seconds
    entries_per_set = int(request.args.get('entries', 10))  # Default to 10 entries

    list_length = len(combined_list)

    if list_length > 10:
        current_time = int(time.time())
        offset = (current_time // refresh_interval) % max(list_length, 1)

        rotated_list = combined_list[offset:offset + entries_per_set]

        if len(rotated_list) < entries_per_set:
            rotated_list += combined_list[:entries_per_set - len(rotated_list)]

        return jsonify(rotated_list)
    else:
        return jsonify(combined_list)

@app.route('/list2')
def list2data():
    return render_template("list2.html")

@app.route('/list1')
def list1data():
    return render_template('list1.html')

def start_flask_server():
    app.run(port=5000)

def select_file():
    global directory
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    
    if file_path:
        directory = os.path.dirname(file_path)
        result_label.config(text=f"Selected File Directory: {directory}")
        root.destroy()

# Tkinter setup
root = tk.Tk()
root.title("File Directory Finder")
root.geometry("400x200")

root.configure(bg="#f0f4f8")

frame = tk.Frame(root, bg="#ffffff", padx=20, pady=20, borderwidth=2, relief="groove")
frame.pack(padx=10, pady=10, expand=True, fill=tk.BOTH)

open_button = tk.Button(
    frame,
    text="Select File",
    command=select_file,
    bg="#4CAF50",
    fg="#ffffff",
    font=("Arial", 14),
    padx=10,
    pady=5,
    relief="raised",
    borderwidth=2
)
open_button.pack(pady=10)

result_label = tk.Label(
    frame,
    text="Selected File Directory:",
    bg="#ffffff",
    fg="#333333",
    font=("Arial", 12),
    wraplength=350
)
result_label.pack(pady=10)

# Run the Flask server in a separate thread
flask_thread = Thread(target=start_flask_server)
flask_thread.start()

# Run the Tkinter application
root.mainloop()

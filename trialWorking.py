import openpyxl
#version 2 working well 

def read_excel_rows(filepath, fromrow, torow):
    """
    Read and filter rows from an Excel file based on the color of the cell in column D.
    """
    COLOR_CODES = {
        'Red': 'FFFF0000',
        'Blue': 'FF00B0F0',
        'Green': 'FF00B050',
        'Normal': None
    }

    REVERSED_COLOR_CODES = {
        'FFFF0000': 'Red',
        'FF00B0F0': 'Blue',
        'FF00B050': 'Green',
    }

    def get_cell_color(cell):
        """Extract color information from a cell's font."""
        if cell.font and cell.font.color and cell.font.color.type == 'rgb':
            return cell.font.color.rgb
        return None

    # Ensure the file path is valid
    if not filepath:
        raise ValueError("File path must be provided")

    # Load the workbook and the first sheet in read-only mode
    try:
        workbook = openpyxl.load_workbook(filepath, read_only=True)
        sheet = workbook.active
    except Exception as e:
        raise RuntimeError(f"Failed to load workbook: {e}")

    data = []
    
    # Iterate through the specified row range
    for row in sheet.iter_rows(min_row=fromrow, max_row=torow, values_only=False):
        cell = row[3]  # Get the cell in column D (index 3)
        color = get_cell_color(cell)
        cell_data = [c.value for c in row]
        
        if color:
            color_name = REVERSED_COLOR_CODES.get(color, "Normal")
            cell_data.append(color_name)
            
            if color in COLOR_CODES.values():
                data.append(cell_data)
        else:
            cell_data.append("Normal")

    if not data:
        print("Warning: No data found in the specified row range.")

    return data

# Example usage
filepath = 'DAILY JOB 2023.xlsx'
fromrow = 5
torow = 7360

try:
    data = read_excel_rows(filepath, fromrow, torow)
    for index, row in enumerate(data, start=fromrow):
        print(f"Row {index}: {row}")
except Exception as e:
    print(f"An error occurred: {e}")

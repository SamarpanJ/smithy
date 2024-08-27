import tkinter as tk
from tkinter import filedialog
import os
import subprocess

def select_file():
    # Open a file dialog to select a file
    file_path = filedialog.askopenfilename()
    
    if file_path:
        # Get the directory from the file path
        directory = os.path.dirname(file_path)
        result_label.config(text=f"Selected File Directory: {directory}")
        
        # Save the directory to a file
        with open('config.txt', 'w') as f:
            f.write(directory)
        
        # Close the Tkinter app after selecting the file
        root.destroy()
        
        # Start the Flask server
        start_flask_server()

def start_flask_server():
    # Adjust the command as necessary to run your Flask app
    subprocess.Popen(['python', 'app.py'])

# Create the main window
root = tk.Tk()
root.title("File Directory Finder")
root.geometry("400x200")

# Apply a theme color and padding
root.configure(bg="#f0f4f8")

# Create a frame for better layout control
frame = tk.Frame(root, bg="#ffffff", padx=20, pady=20, borderwidth=2, relief="groove")
frame.pack(padx=10, pady=10, expand=True, fill=tk.BOTH)

# Style the button
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

# Style the label
result_label = tk.Label(
    frame,
    text="Selected File Directory:",
    bg="#ffffff",
    fg="#333333",
    font=("Arial", 12),
    wraplength=350
)
result_label.pack(pady=10)

# Run the application
root.mainloop()

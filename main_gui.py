# main_gui.py

"""
This module is responsible for creating and managing the graphical user interface of the application.
It provides a user-friendly interface for inputting chart parameters and initiating chart creation.
"""

# Import Libraries
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import re
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def create_window(create_chart_callback):
    window = tk.Tk()
    window.title("AutoChart")
    window.geometry("400x500")  # Increased height to accommodate new elements
    window.resizable(False, False)

    # Styling using ttkthemes for a more modern look
    style = ttk.Style()
    style.theme_use("clam")

    file_path = tk.StringVar()
    sheet_name = tk.StringVar()
    chart_type = tk.StringVar()
    chart_title = tk.StringVar()
    start_cell = tk.StringVar()
    end_cell = tk.StringVar()
    target_sheet_name = tk.StringVar()

    # Creating Widgets
    # Container frame for better spacing and styling
    main_frame = ttk.Frame(window, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)

    # File selection
    file_frame = ttk.LabelFrame(main_frame, text="Select Excel File")
    file_frame.pack(pady=10, fill=tk.X)
    file_button = ttk.Button(file_frame, text="Browse", command=lambda: select_file(file_path, sheet_combo, target_sheet_combo))
    file_button.pack(side=tk.RIGHT, padx=5)
    file_entry = ttk.Entry(file_frame, textvariable=file_path, state="readonly")
    file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

    # Sheet selection
    sheet_frame = ttk.LabelFrame(main_frame, text="Select Source Sheet")
    sheet_frame.pack(pady=10, fill=tk.X)
    sheet_combo = ttk.Combobox(sheet_frame, textvariable=sheet_name, state="readonly")
    sheet_combo.pack(fill=tk.X, expand=True, padx=5)

    # Target sheet selection
    target_sheet_frame = ttk.LabelFrame(main_frame, text="Select Target Sheet for Chart")
    target_sheet_frame.pack(pady=10, fill=tk.X)
    target_sheet_combo = ttk.Combobox(target_sheet_frame, textvariable=target_sheet_name, state="readonly")
    target_sheet_combo.pack(fill=tk.X, expand=True, padx=5)

    # Chart type selection
    chart_frame = ttk.LabelFrame(main_frame, text="Select Chart Type")
    chart_frame.pack(pady=10, fill=tk.X)
    chart_combo = ttk.Combobox(chart_frame, textvariable=chart_type, 
                                values=["Bar Chart", "Line Chart", "Pie Chart", "Area Chart", 
                                        "Bubble Chart", "Radar Chart", "Doughnut Chart", "Scatter Chart"], 
                                state="readonly")
    chart_combo.pack(fill=tk.X, expand=True, padx=5)

    # Chart Title input
    chart_title_frame = ttk.LabelFrame(main_frame, text="Chart Title")
    chart_title_frame.pack(pady=10, fill=tk.X)
    chart_title_entry = ttk.Entry(chart_title_frame, textvariable=chart_title)
    chart_title_entry.pack(fill=tk.X, expand=True, padx=5)

    # Cell range input
    cell_range_frame = ttk.LabelFrame(main_frame, text="Cell Range") 
    cell_range_frame.pack(pady=10, fill=tk.X)

    start_cell_label = ttk.Label(cell_range_frame, text="Start Cell:")
    start_cell_label.pack(side=tk.LEFT, padx=5)
    start_cell_entry = ttk.Entry(cell_range_frame, textvariable=start_cell)
    start_cell_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

    end_cell_label = ttk.Label(cell_range_frame, text="End Cell:")
    end_cell_label.pack(side=tk.LEFT, padx=5)
    end_cell_entry = ttk.Entry(cell_range_frame, textvariable=end_cell)
    end_cell_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

    # Run button
    run_button = ttk.Button(main_frame, text="Run", command=lambda: run_app(window, file_path, sheet_name, start_cell, end_cell, chart_type, chart_title, target_sheet_name, create_chart_callback))
    run_button.pack(pady=20)

    center_window(window)
    window.protocol("WM_DELETE_WINDOW", lambda: on_closing(window))

    return window

def select_file(file_path, sheet_combo, target_sheet_combo):
    filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if filepath:
        file_path.set(filepath)
        populate_sheets(filepath, sheet_combo, target_sheet_combo)


def populate_sheets(filepath, sheet_combo, target_sheet_combo):
    try:
        xls = pd.ExcelFile(filepath)
        sheet_names = xls.sheet_names
        sheet_combo.config(values=sheet_names)
        target_sheet_combo.config(values=sheet_names)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load Excel file: {str(e)}")

def run_app(window, file_path, sheet_name, start_cell, end_cell, chart_type, chart_title, target_sheet_name, create_chart_callback):
    try:
        if not file_path.get():
            raise ValueError("Please select an Excel File.")
        if not sheet_name.get():
            raise ValueError("Please select a source sheet.")
        if not start_cell.get() or not end_cell.get():
            raise ValueError("Please enter both Start Cell and End Cell.")
        if not chart_type.get():
            raise ValueError("Please select a Chart type.")
        if not chart_title.get():
            raise ValueError("Enter a Chart Title.")

        start = start_cell.get().upper()
        end = end_cell.get().upper()

        # Improved regex pattern to validate cell references
        cell_pattern = re.compile(r'^([A-Z]{1,3})(\d+)$')

        # Validate the Start Cell and End Cell inputs
        start_match = cell_pattern.match(start)
        end_match = cell_pattern.match(end)

        if not start_match or not end_match:
            raise ValueError("Invalid Start Cell or End Cell format. Please enter a valid cell reference (e.g., A1, BG23).")

        start_col, start_row = start_match.groups()
        end_col, end_row = end_match.groups()

        # Convert column letters to numbers
        def col_to_num(col):
            return sum((ord(c) - 64) * (26 ** i) for i, c in enumerate(reversed(col)))

        start_col_num = col_to_num(start_col)
        end_col_num = col_to_num(end_col)

        # Additional validations
        if start_col_num > 16384 or end_col_num > 16384:  # Excel's maximum column (XFD)
            raise ValueError("Column reference exceeds Excel's maximum (XFD).")

        if int(start_row) <= 0 or int(end_row) <= 0:
            raise ValueError("Row numbers must be positive integers.")

        if int(start_row) > 1048576 or int(end_row) > 1048576:  # Excel's maximum row
            raise ValueError("Row number exceeds Excel's maximum (1048576).")

        # Ensure start cell is before end cell
        if start_col_num > end_col_num or (start_col_num == end_col_num and int(start_row) > int(end_row)):
            raise ValueError("Start Cell must be before End Cell in the spreadsheet.")

        # Show a simple message
        messagebox.showinfo("Processing", "Generating chart. Please wait...")

        logging.info(f"Starting chart creation with parameters: file={file_path.get()}, sheet={sheet_name.get()}, "
                     f"start_cell={start_cell.get()}, end_cell={end_cell.get()}, chart_type={chart_type.get()}, "
                     f"chart_title={chart_title.get()}, target_sheet={target_sheet_name.get()}")

        # Call the chart creation function
        create_chart_callback(file_path.get(), sheet_name.get(), start_cell.get(), end_cell.get(), 
                              chart_type.get(), chart_title.get(), target_sheet_name.get())

        # Close the main window
        window.destroy()

    except ValueError as e:
        messagebox.showerror("Validation Error", str(e))
        logging.error(f"Validation error: {str(e)}")
    except Exception as e:
        messagebox.showerror("Unexpected Error", f"An unexpected error occurred: {str(e)}")
        logging.error(f"Unexpected error: {str(e)}", exc_info=True)

def on_closing(window):
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        logging.info("Application closed by user")
        window.destroy()

def center_window(window):
    window.update_idletasks()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    window_width = window.winfo_width()
    window_height = window.winfo_height()
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    window.geometry(f"+{x}+{y}")

if __name__ == "__main__":
    create_window(lambda: print("Chart creation function not provided"))
    tk.mainloop()
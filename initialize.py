# initialize.py

"""
This module serves as the entry point for the application. It initializes the necessary components,
checks for running Excel instances, and launches the main GUI.
"""

# Import libraries
import os
import psutil
import tkinter as tk
import sys
import ctypes
import time
import logging

# Import modules chart and main_gui
import main_gui

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def is_excel_running():
    return 'EXCEL.EXE' in (p.name().upper() for p in psutil.process_iter(['name']))

def close_excel():
    for proc in psutil.process_iter(['name']):
        if proc.info['name'].upper() == 'EXCEL.EXE':
            try:
                proc.terminate()
                proc.wait(timeout=10)
            except psutil.TimeoutExpired:
                proc.kill()
            except psutil.NoSuchProcess:
                pass

    timeout = time.time() + 30  # 30 seconds from now
    while is_excel_running():
        if time.time() > timeout:
            logging.error("Failed to close all Excel processes")
            return False
        time.sleep(0.1)
    return True


# Hide the console window
ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)


def create_chart_from_inputs(file_path, sheet_name, start_cell, end_cell, chart_type, chart_title, target_sheet_name=None):
    import chart
    import tkinter as tk
    from tkinter import messagebox
    import os
    import logging

    # Mapping of chart type names to their corresponding values used in openpyxl
    chart_type_mapping = {
        "Bar Chart": "bar", "Line Chart": "line", "Area Chart": "area",
        "Bubble Chart": "bubble", "Radar Chart": "radar", "Pie Chart": "pie",
        "Doughnut Chart": "doughnut", "Scatter Chart": "scatter"
    }
    chart_type_value = chart_type_mapping.get(chart_type, "bar")

    try:
        chart.create_chart(file_path, sheet_name, start_cell, end_cell, 
                           chart_type_value, chart_title, target_sheet_name)
        messagebox.showinfo("Success", "Chart created successfully!")
        os.startfile(file_path)
    except Exception as e:
        logging.error(f"Failed to create chart: {str(e)}")
        messagebox.showerror("Error", f"Failed to create chart: {str(e)}")
    
def main():
    if is_excel_running():
        if not close_excel():
            tk.messagebox.showerror("Error", "Failed to close all Excel instances. Please close them manually and try again.")
            sys.exit(1)

    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

    window = main_gui.create_window(create_chart_from_inputs)
    window.mainloop()

if __name__ == "__main__":
    main()
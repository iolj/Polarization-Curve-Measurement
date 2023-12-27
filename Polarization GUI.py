import time
import json
import os
import pyvisa
from datetime import date
import numpy as np
import pandas as pd
import threading
import tkinter as tk
from tkinter import ttk, Scrollbar, scrolledtext, Entry, Label, filedialog
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Initialize global variables
stop_flag = False
voltage = []
start_time = 0.0
current = [0.01, 0.1, 0.2, 0.3, 0.4, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0, 6.5, 7.0, 7.5, 
           8.0, 8.5, 9.0, 9.5, 10.0, 10.5, 11.0, 11.5, 12.0, 12.5, 13.0, 13.5, 14.0, 14.5, 15.0, 17.5, 20.0, 22.5, 25.0, 27.5, 30.0, 32.5, 35.0, 37.5, 40.0]
currentPlot = []

def plot_settings():
    ax.set_title("Polarization Curve")
    ax.set_xlabel("Current Density (mA/cm²)")
    ax.set_ylabel("Voltage (V)")
    ax.set_xlim(0, 1600)
    ax.set_ylim(1.0, 2.5)
    ax.set_xticks(np.arange(0, 1700, 200))
    ax.set_yticks(np.arange(1.0, 2.6, 0.5))
    
def load_current_from_csv():
    global current
    file_path = file_entry.get()
    try:
        with open(file_path, 'r') as file:
            values_str = file.read()
            values_list = values_str.split(',')
            current.clear()
            current = [float(value) for value in values_list]
            print("Current list updated:", current)
    except Exception as e:
        print(f"Error reading CSV file: {e}")
      
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        file_entry.delete(0, tk.END)  # Clear the current text in the entry field
        file_entry.insert(0, file_path)  # Set the selected file path
        
  
def load_default_values():
    default_values = {}
    try:
        with open("config.json", "r") as config_file:
            default_values = json.load(config_file)
    except FileNotFoundError:
        # Handle the case when the config file is not found
        pass
    
    return default_values

def update_values():
    value1 = voltageLimit_value.get()
    value2 = activateTime_value.get()
    value3 = interval_value.get()
    # Process the values here as needed
    terminal_output.insert(tk.END, f'Voltage limit = {value1}V\nActivation time = {value2}s\nInterval time = {value3}s\nValues updateted.\n')
    # Add processing for additional inputs if needed

def start_measurement():
    global stop_flag, voltage, start_time
    for item in tree.get_children():
        tree.delete(item)
    stop_flag = False
    start_time = time.time()
    # Create a thread for the measurement process
    measurement_thread = threading.Thread(target=measurement_loop)
    measurement_thread.daemon = True
    measurement_thread.start()
    ax.clear()
    # ax.plot(currentPlot, voltage, marker='o', markersize='8')
    plot_settings()
    canvas.draw()
    
def clear_terminal():
    terminal_output.delete(1.0, tk.END)  # Delete all text in the widget
        
def stop_measurement():
    global stop_flag, voltage
    terminal_output.insert(tk.END, f'Manual stop, ending the program...\n')
    terminal_output.insert(tk.END, f'Voltage = {voltage}\n')
    df = pd.DataFrame(voltage)
    excel_file = 'output.xlsx'
    df.to_excel(excel_file, index=False, header=False)
    stop_flag = True

def update_table(i):
    global voltage, current
    tree.insert("", "end", values=(f"{current[i]} A", f"{voltage[i]} V"))

def update_column_widths(event):
    # Calculate the new width for each column based on the window width
    window_width = event.width
    new_column_width = (window_width - 4) // 2  # Subtract any padding or borders
    tree.column("Current", width=new_column_width, minwidth=new_column_width, stretch=False)
    tree.column("Voltage", width=new_column_width, minwidth=new_column_width, stretch=False)

def export_voltage():
    global voltage
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile="output")
    if file_path:
        df = pd.DataFrame(voltage)
        df.to_excel(file_path, index=False, header=False)
        terminal_output.insert(tk.END, f'File Exported.')

def update_plot():
    global voltage, ax, canvas, current, currentPlot
    ax.clear()
    ax.plot(currentPlot, voltage, marker='o', markersize='8')
    plot_settings()
    canvas.draw()

def measurement_loop():
    global stop_flag, voltage, start_time, current, currentPlot
    i = 0
    while not stop_flag and i < len(current):
        measured_voltage = 1
        update_button.config(state=tk.DISABLED)
        load_button.config(state=tk.DISABLED)
        if measured_voltage < float(voltageLimit_value.get()):
            elapsed_time = time.time() - start_time
            if elapsed_time >= float(interval_value.get()):
                measured_voltage = 1
                voltage.append(measured_voltage)
                # Append the measurement results to the terminal_output widget
                terminal_output.insert(tk.END, f'{date.today()} {time.strftime("%H:%M:%S", time.localtime())}')
                terminal_output.insert(tk.END, f' {str(current[i])}A {str(voltage[i])}V\n')
                terminal_output.see(tk.END)  # Scroll to the end to show the latest output
                start_time = time.time()  # Reset the start time
                currentPlot.append(current[i] * 40)
                update_table(i)
                i += 1
                # Update the plot with the latest data in the main thread
                root.after(0, update_plot)
                
            # Sleep for a short interval to allow the "stop_flag" check
            time.sleep(0.1)
        else:
            terminal_output.insert(tk.END, f'Exceed voltage limit, ending the program...\n')
            terminal_output.insert(tk.END, f'Voltage = {voltage}')
            df = pd.DataFrame(voltage)
            excel_file = 'output.xlsx'
            df.to_excel(excel_file, index=False, header=False)
            stop_flag = True
    update_button.config(state=tk.ACTIVE)
    load_button.config(state=tk.ACTIVE)
    stop_flag = True
 
# Create the main application window
root = tk.Tk()
root.title("Measurement Program")

# File import
browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.grid(column=2, row=2)
default_file_path = os.path.abspath("current.csv")
file_label = tk.Label(root, text="CSV File Path:")
file_label.grid(column=0, row=2)

file_entry = tk.Entry(root, width=40)
file_entry.insert(0, default_file_path)  # Set the default file path
file_entry.grid(column=1, row=2)

load_button = tk.Button(root, text="Load CSV as current", command=load_current_from_csv)
load_button.grid(column=3, row=2)

# Create button to submit values
update_button = tk.Button(root, text="Update Values", command=update_values)
update_button.grid(column=0, row=7)

# Create entry widgets
default_values = load_default_values()
voltageLimit_value = tk.StringVar(value=default_values.get("Voltage Limit", ""))
activateTime_value = tk.StringVar(value=default_values.get("Activation Time", ""))
interval_value = tk.StringVar(value=default_values.get("Interval Time", ""))

voltageLimit_entry = Entry(root, textvariable=voltageLimit_value)
voltageLimit_label = Label(root, text="Voltage Limit (V):")
activateTime_entry = Entry(root, textvariable=activateTime_value)
activateTime_label = Label(root, text="Activation Time (s):")
interval_entry = Entry(root, textvariable=interval_value)
interval_label = Label(root, text="Interval (s):")

voltageLimit_label.grid(column=0, row=4)
voltageLimit_entry.grid(column=1, row=4)
activateTime_label.grid(column=0, row=5)
activateTime_entry.grid(column=1, row=5)
interval_label.grid(column=0, row=6)
interval_entry.grid(column=1, row=6)

# Create a Treeview widget to display the table
tree = ttk.Treeview(root, columns=("Current", "Voltage"), show="headings")
tree.heading("Current", text="Current")
tree.heading("Voltage", text="Voltage")
tree.column("Current", width=100, minwidth=50, stretch=False)
tree.column("Voltage", width=100, minwidth=50, stretch=False)
tree.bind("<Configure>", update_column_widths)
tree.grid(column=5, row=0, rowspan=12, sticky="nsew", padx=5, pady=5)

# Create a vertical scrollbar
v_scrollbar = Scrollbar(root, orient="vertical", command=tree.yview)

# Configure the Treeview to use the scrollbar
tree.configure(yscrollcommand=v_scrollbar.set)
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(5, weight=1)
# Pack the Treeview and scrollbar
v_scrollbar.grid(column=7, row=0, rowspan=12, sticky="nsew")

# Create Export voltage button
export_button = tk.Button(root, text="Export Voltage", command=export_voltage)
export_button.grid(column=0, row=8)

# Create clear terminal button
clearTerminal_button = tk.Button(root, text="Clear Terminal", command=clear_terminal)
clearTerminal_button.grid(column=0, row=9)

# Create Start and Stop buttons
start_button = tk.Button(root, text="Start Measurement", command=start_measurement)
start_button.grid(column=0, row=0)

stop_button = tk.Button(root, text="Stop Measurement", command=stop_measurement)
stop_button.grid(column=0, row=1)

# Create terminal view
terminal_output = scrolledtext.ScrolledText(root, wrap=tk.WORD)
terminal_output.grid(column=0, row=11, columnspan=4, sticky="nsew", padx=5, pady=5)

# Create a Matplotlib figure and axis for plotting
fig, ax = plt.subplots(figsize=(8, 6))
plot_settings()

# Create a Matplotlib canvas to embed the figure in the Tkinter GUI
canvas = FigureCanvasTkAgg(fig, master=root)
canvas_widget = canvas.get_tk_widget()
canvas_widget.grid(column=3, row=0, rowspan=10)

# Start the Tkinter main loop
root.mainloop()

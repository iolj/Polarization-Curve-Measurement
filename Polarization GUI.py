import time
from datetime import date
import numpy as np
import pandas as pd
import threading
import tkinter as tk
from tkinter import ttk, Scrollbar
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
# Initialize global variables
stop_flag = False
voltage = []
start_time = 0.0
current = [0.01, 0.1, 0.2, 0.3, 0.4, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0, 6.5, 7.0, 7.5, 8.0, 8.5, 9.0, 9.5, 10.0, 
           10.5, 11.0, 11.5, 12.0, 12.5, 13.0, 13.5, 14.0, 14.5, 15.0, 17.5, 20.0, 22.5, 25.0, 27.5, 30.0, 32.5, 35.0, 37.5, 40.0]
currentPlot = []
voltageLimit = 2.1
activationTime = 1
time_interval = 3
def start_measurement():
    global stop_flag, voltage, start_time
    stop_flag = False
    start_time = time.time()

    # Create a thread for the measurement process
    measurement_thread = threading.Thread(target=measurement_loop)
    measurement_thread.daemon = True
    measurement_thread.start()

def stop_measurement():
    global stop_flag, voltage
    print('Manual stop, ending the program...')
    print(voltage)
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
    df = pd.DataFrame(voltage)
    excel_file = 'output.xlsx'
    df.to_excel(excel_file, index=False, header=False)
    
def update_plot():
    global voltage, ax, canvas, current, currentPlot
    ax.clear()
    ax.plot(currentPlot, voltage, marker='o', markersize='8')
    ax.set_title("Polarization Measurement")
    ax.set_xlabel("Current Density (mA/cm²)")
    ax.set_ylabel("Voltage (V)")
    ax.set_xlim(0, 1600)
    ax.set_ylim(1.0, 2.5)
    ax.set_xticks(np.arange(0, 1700, 200))
    ax.set_yticks(np.arange(1.0, 2.6, 0.5))
    canvas.draw()

def measurement_loop():
    global stop_flag, voltage, start_time, current, currentPlot
    i = 0
    while not stop_flag and i < len(current):
        measured_voltage = 1
        if measured_voltage < voltageLimit:
            elapsed_time = time.time() - start_time
            if elapsed_time >= time_interval:
                measured_voltage = 1
                voltage.append(measured_voltage)
                print(date.today(), time.strftime("%H:%M:%S", time.localtime()))
                print(f'{str(current[i])}A {str(voltage[i])}V')
                start_time = time.time()  # Reset the start time
                currentPlot.append(current[i] * 40)
                update_table(i)
                i += 1
                # Update the plot with the latest data in the main thread
                root.after(0, update_plot)
                
            # Sleep for a short interval to allow the "stop_flag" check
            time.sleep(0.1)
        else:
            print('Exceed voltage limit, ending the program...')
            print(voltage)
            df = pd.DataFrame(voltage)
            excel_file = 'output.xlsx'
            df.to_excel(excel_file, index=False, header=False)
            stop_flag = True
    stop_flag = True
    # When the loop exits, you can quit the Tkinter application
    root.quit()
 
# Create the main application window
root = tk.Tk()
root.title("Measurement Program")

# Create a Treeview widget to display the table
tree = ttk.Treeview(root, columns=("Current", "Voltage"), show="headings")
tree.heading("Current", text="Current")
tree.heading("Voltage", text="Voltage")
tree.column("Current", width=100, minwidth=50, stretch=False)
tree.column("Voltage", width=100, minwidth=50, stretch=False)
tree.bind("<Configure>", update_column_widths)
tree.pack()

# Create a vertical scrollbar
v_scrollbar = Scrollbar(root, orient="vertical", command=tree.yview)

# Configure the Treeview to use the scrollbar
tree.configure(yscrollcommand=v_scrollbar.set)

# Pack the Treeview and scrollbar
tree.pack(side="right", fill="both", expand=True)
v_scrollbar.pack(side="right", fill="y")

# Create Export voltage button
export_button = tk.Button(root, text="Export Voltage", command=export_voltage)
export_button.pack()
# Create Start and Stop buttons
start_button = tk.Button(root, text="Start Measurement", command=start_measurement)
start_button.pack()

stop_button = tk.Button(root, text="Stop Measurement", command=stop_measurement)
stop_button.pack()

# Create a Matplotlib figure and axis for plotting
fig, ax = plt.subplots(figsize=(6, 4))
ax.set_title("Polarization Measurement")
ax.set_xlabel("Current Density (mA/cm²)")
ax.set_ylabel("Voltage (V)")
ax.set_xlim(0, 1600)
ax.set_ylim(1.0, 2.5)
ax.set_xticks(np.arange(0, 1700, 200))
ax.set_yticks(np.arange(1.0, 2.6, 0.5))

# Create a Matplotlib canvas to embed the figure in the Tkinter GUI
canvas = FigureCanvasTkAgg(fig, master=root)
canvas_widget = canvas.get_tk_widget()
canvas_widget.pack()

# Start the Tkinter main loop
root.mainloop()

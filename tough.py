# Ensure the required packages are installed:
# pip install minimalmodbus pyserial pandas openpyxl
from datetime import datetime  # Used for recording time
import minimalmodbus  # type: ignore
import time
import logging
import matplotlib.pyplot as plt  # Used for plotting pressure-time curves
import threading  # Used for implementing manual stop functionality
import os  # Used for creating folders
from threading import Thread, Event  # Used for thread control
import tkinter as tk  # Used for creating button interface
import matplotlib.animation as animation  # Used for real-time dynamic plotting
import matplotlib  # Used for setting backend
from tkinter import filedialog  # Import file dialog module
import pandas as pd  # Import pandas for generating Excel files
matplotlib.use("TkAgg")  # Set TkAgg backend for tkinter compatibility

# Configure logging
logging.basicConfig(filename="error.log", level=logging.ERROR, format="%(asctime)s - %(message)s")

# Configure Modbus instrument
instrument = minimalmodbus.Instrument(port="COM7", slaveaddress=1)
instrument.serial.baudrate = 9600
instrument.serial.bytesize = 8
instrument.serial.parity = minimalmodbus.serial.PARITY_NONE
instrument.serial.stopbits = 1
instrument.serial.timeout = 1  # Timeout (seconds)

# Initialise lists for storing pressure values and timestamps
pressure_data = []
time_data = []
stop_flag = False  # Flag for controlling programme stop

# Initialise stop event
stop_event = Event()

# Create a "test" folder on the desktop
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
image_folder = os.path.join(desktop_path, "test")
os.makedirs(image_folder, exist_ok=True)

def stop_program():
    global stop_flag
    input("Press Enter to stop the programme...")
    stop_flag = True

def check_serial_port():
    try:
        # Check if the serial port is available
        instrument.serial.open()
    except PermissionError as e:
        logging.error(f"Serial port access denied: {e}")
        print("Serial port access denied. Please check the following:")
        print("1. Ensure no other programme is using the serial port.")
        print("2. Check if the current user has permission to access the serial port.")
        print("   - On Windows, try running the programme as administrator.")
        print("   - On Linux, ensure the user is in the `dialout` or `tty` group.")
        exit(1)
    except Exception as e:
        logging.error(f"Unable to open serial port: {e}")
        print(f"Unable to open serial port: {e}")
        instrument.serial.close()

# Check serial port access permissions
# check_serial_port()

def send_command():
    try:
        # Send command: 01 03 00 00 00 01 84 0A
        command = bytes.fromhex("01 03 00 00 00 01 84 0A")
        instrument.serial.write(command)
        time.sleep(0.1)  # Wait to ensure the command is processed
    except Exception as e:
        logging.error(f"Failed to send command: {e}")

def read_pressure(retries=3):
    for attempt in range(retries):
        try:
            # Send command
            send_command()
            
            # Read raw byte data returned from the serial port
            raw_data = instrument.serial.read(7)  # Assume return data length is 7 bytes
            if len(raw_data) < 5:
                raise ValueError("Returned data length insufficient for parsing")
            
            # Extract the fourth and fifth bytes
            byte4 = raw_data[3]
            byte5 = raw_data[4]
            
            # Convert two bytes to float (assume big-endian)
            value = (byte4 << 8) | byte5
            pressure = value / 100.0  # Assume divide by 100 for two decimal places
            
            return pressure
        except Exception as e:
            logging.error(f"Read failed (attempt {attempt + 1}/{retries}): {e}")
            time.sleep(1)  # Wait before retrying
    return None  # Return None after multiple failed attempts

def start_realtime_plot():
    """Start real-time dynamic plotting"""
    def plot():
        fig, ax = plt.subplots()
        ax.set_title("F-Time Curve (Real-time)", color="#1a355e", fontsize=18)  # Edit bracket content
        ax.set_xlabel("Time (s)", fontsize=14)
        ax.set_ylabel("F (kg)", fontsize=14)
        ax.set_facecolor("#ffffff")
        ax.tick_params(axis='both', labelsize=13)
        ax.grid(True, color="#d0d7e5", linestyle="--", linewidth=0.7, alpha=0.7)
        line, = ax.plot([], [], label="F (kg)", color="#2176ae", linewidth=2)
        ax.legend(fontsize=13)

        def init():
            """Initialise plot"""
            ax.set_xlim(0, 10)  # Initial X axis range
            ax.set_ylim(-1, 1)  # Initial Y axis range
            return line,

        def update(frame):
            """Update plot"""
            if len(time_data) > 0:
                ax.set_xlim(0, max(time_data) + 1)  # Dynamically adjust X axis range
                ax.set_ylim(min(pressure_data) - 1, max(pressure_data) + 1)  # Dynamically adjust Y axis range
                line.set_data(time_data, pressure_data)
            return line,

        ani = animation.FuncAnimation(fig, update, init_func=init, blit=True, interval=50)
        plt.show()

    Thread(target=plot).start()  # Run plotting in a separate thread

def start_data_collection():
    """Start data collection"""
    global zero_offset, start_time, pressure_data, time_data  # Declare globals
    zero_offset = None  # Initialise zero_offset
    start_time = None
    pressure_data = []
    time_data = []
    stop_event.clear()
    print("Data collection started...")

    def collect_data():
        global zero_offset  # Declare zero_offset as global in thread
        while not stop_event.is_set():
            pressure_value = read_pressure()
            if pressure_value is not None:
                current_time = datetime.now()
                if zero_offset is None:
                    zero_offset = pressure_value  # Set value at zeroing
                    start_time = current_time
                    print(f"Zeroing complete, offset: {zero_offset:.2f} kg")
                else:
                    adjusted_pressure = pressure_value - zero_offset
                    elapsed_time = (current_time - start_time).total_seconds()
                    print(f"Real-time pressure: {adjusted_pressure:.2f} kg, Time: {elapsed_time:.2f} s")
                    pressure_data.append(adjusted_pressure)
                    time_data.append(elapsed_time)
            else:
                print("Failed to read pressure value, logged.")
            time.sleep(0.05)

        print("Data collection thread stopped.")

    # Start data collection thread
    Thread(target=collect_data, daemon=True).start()

def manual_zero():
    """Manual zeroing functionality"""
    global zero_offset
    pressure_value = read_pressure()
    if pressure_value is not None:
        zero_offset = pressure_value
        print(f"Manual zeroing complete, offset: {zero_offset:.2f} kg")
    else:
        print("Zeroing failed, unable to read pressure value.")

def select_save_path():
    """Let user select image save path"""
    global image_folder
    selected_folder = filedialog.askdirectory(initialdir=desktop_path, title="Select image save path")
    if selected_folder:  # If user selected a path
        image_folder = selected_folder
        print(f"Image save path changed to: {image_folder}")
    else:
        print("Using default path: 'test' folder on desktop")

def stop_data_collection():
    """Stop data collection"""
    stop_event.set()
    print("Data collection stopped.")
    if pressure_data and time_data:
        # Save image
        plt.figure(figsize=(10, 6))
        plt.plot(time_data, pressure_data, label="F (kg)", marker="o")
        plt.xlabel("Time (s)")
        plt.ylabel("F (kg)")
        plt.title("F-Time Curve")
        plt.legend()
        plt.gca().set_facecolor("white")  # Set axis background to white
        plt.grid(False)  # Remove grid from background

        # Display each peak value on the curve
        for i in range(1, len(pressure_data) - 1):
            if (pressure_data[i] >= pressure_data[i - 1]) and (pressure_data[i] > pressure_data[i + 1]):
                plt.text(time_data[i], pressure_data[i], f"{pressure_data[i]:.2f}", fontsize=8, ha="centre", va="bottom", colour="red")

        # Save image to user-selected or default path
        image_path = os.path.join(image_folder, f"Result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
        plt.savefig(image_path)
        print(f"Image saved to: {image_path}")
        plt.close()  # Close figure to avoid conflict with Tkinter

        # Save data to Excel file
        excel_path = os.path.join(image_folder, f"Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        try:
            data = {"Time (s)": time_data, "F (kg)": pressure_data}
            df = pd.DataFrame(data)
            df.to_excel(excel_path, index=False)
            print(f"Excel data saved to: {excel_path}")
        except Exception as e:
            print(f"Error saving Excel file: {e}")

    # Do not destroy window, only stop collection
    print("Collection stopped, but window remains available.")

def create_gui():
    """Create interface with control buttons and real-time plotting (improved layout and larger font)"""
    global root, ax, line, ani  # Declare globals for access in other functions
    root = tk.Tk()
    root.title("LoaDC")
    root.configure(bg="#f4f8fb")

    # Main frame, split into left and right areas
    main_frame = tk.Frame(root, bg="#f4f8fb")
    main_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

    # Data display area (left, with border)
    data_frame = tk.LabelFrame(
        main_frame, text="Data Display", bg="#f4f8fb", fg="#1a355e",
        bd=2, relief=tk.GROOVE, font=("Arial", 16, "bold")
    )
    data_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 12), pady=0)

    # Control button area (right, with border)
    control_frame = tk.LabelFrame(
        main_frame, text="Controls", bg="#f4f8fb", fg="#1a355e",
        bd=2, relief=tk.GROOVE, font=("Arial", 16, "bold")
    )
    control_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=0, pady=0)

    # Chart area
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    fig, ax = plt.subplots(figsize=(7, 4))
    fig.patch.set_facecolor("#f4f8fb")
    ax.set_title("F-Time Curve (Real-time)", color="#1a355e", fontsize=18)
    ax.set_xlabel("Time (s)", fontsize=14)
    ax.set_ylabel("F (kg)", fontsize=14)
    ax.set_facecolor("#ffffff")
    ax.tick_params(axis='both', labelsize=13)
    ax.grid(True, color="#d0d7e5", linestyle="--", linewidth=0.7, alpha=0.7)
    line, = ax.plot([], [], label="F (kg)", color="#2176ae", linewidth=2)
    ax.legend(fontsize=13)

    # Initial Y axis range 0-10kg
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)

    def init():
        ax.set_xlim(0, 10)
        ax.set_ylim(0, 10)
        return line,

    def update(frame):
        if len(time_data) > 0:
            ax.set_xlim(0, max(time_data) + 1)
            line.set_data(time_data, pressure_data)
        return line,

    ani = animation.FuncAnimation(fig, update, init_func=init, blit=True, interval=50)

    canvas = FigureCanvasTkAgg(fig, master=data_frame)
    canvas_widget = canvas.get_tk_widget()
    canvas_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # Beautify control button area
    def style_button(btn):
        btn.configure(
            bg="#e6eef8", fg="#1a355e", activebackground="#d0e2f2", activeforeground="#1a355e",
            relief=tk.FLAT, bd=0, width=18, height=2, cursor="hand2", font=("Arial", 14, "bold")
        )

    start_button = tk.Button(control_frame, text="Start Collection", command=start_data_collection)
    style_button(start_button)
    start_button.pack(fill=tk.X, padx=12, pady=(12, 4))

    stop_button = tk.Button(control_frame, text="Stop Collection", command=stop_data_collection)
    style_button(stop_button)
    stop_button.pack(fill=tk.X, padx=12, pady=4)

    zero_button = tk.Button(control_frame, text="Zero", command=manual_zero)
    style_button(zero_button)
    zero_button.pack(fill=tk.X, padx=12, pady=4)

    path_button = tk.Button(control_frame, text="Select Save Path", command=select_save_path)
    style_button(path_button)
    path_button.pack(fill=tk.X, padx=12, pady=4)

    def set_ylim_custom():
        top = tk.Toplevel(root)
        top.title("Set Y Axis Range")
        tk.Label(top, text="Y axis min (kg):", font=("Arial", 13)).grid(row=0, column=0, padx=5, pady=5)
        tk.Label(top, text="Y axis max (kg):", font=("Arial", 13)).grid(row=1, column=0, padx=5, pady=5)
        min_var = tk.StringVar(value="0")
        max_var = tk.StringVar(value="10")
        min_entry = tk.Entry(top, textvariable=min_var, font=("Arial", 13))
        max_entry = tk.Entry(top, textvariable=max_var, font=("Arial", 13))
        min_entry.grid(row=0, column=1, padx=5, pady=5)
        max_entry.grid(row=1, column=1, padx=5, pady=5)
        def apply():
            try:
                ymin = float(min_var.get())
                ymax = float(max_var.get())
                ax.set_ylim(ymin, ymax)
                canvas.draw_idle()
                top.destroy()
            except Exception:
                tk.messagebox.showerror("Error", "Please enter valid numbers.")
        apply_button = tk.Button(top, text="Apply", command=apply, font=("Arial", 13, "bold"))
        apply_button.grid(row=2, column=0, columnspan=2, pady=5)

    ylim10_button = tk.Button(control_frame, text="Y axis 0-10kg", command=lambda: (ax.set_ylim(0, 10), canvas.draw_idle()))
    style_button(ylim10_button)
    ylim10_button.pack(fill=tk.X, padx=12, pady=4)
    ylim50_button = tk.Button(control_frame, text="Y axis 0-50kg", command=lambda: (ax.set_ylim(0, 50), canvas.draw_idle()))
    style_button(ylim50_button)
    ylim50_button.pack(fill=tk.X, padx=12, pady=4)
    ylim_custom_button = tk.Button(control_frame, text="Custom Y axis", command=set_ylim_custom)
    style_button(ylim_custom_button)
    ylim_custom_button.pack(fill=tk.X, padx=12, pady=4)

    exit_button = tk.Button(control_frame, text="Exit", command=lambda: root.after(0, root.destroy))
    style_button(exit_button)
    exit_button.pack(fill=tk.X, padx=12, pady=(4, 12))

    # Hover effect
    def on_enter(e): e.widget.config(bg="#d0e2f2")
    def on_leave(e): e.widget.config(bg="#e6eef8")
    for btn in [start_button, stop_button, zero_button, path_button, ylim10_button, ylim50_button, ylim_custom_button, exit_button]:
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)

    root.mainloop()

# Main programme entry
if __name__ == "__main__":
    try:
        create_gui()
    except KeyboardInterrupt:
        print("Programme terminated")
    finally:
        if not stop_event.is_set():
            stop_data_collection()
        instrument.serial.close()
        if not stop_event.is_set():
            stop_data_collection()
        instrument.serial.close()

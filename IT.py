import tkinter as tk
from tkinter import ttk, PhotoImage
import pandas as pd
from datetime import datetime
import serial

class CompanyInterface:
    def __init__(self, root):
        self.root = root
        self.root.title("Company Interface")
        
        # Get screen width and height
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        # Load company logo
        self.logo_image = PhotoImage(file="company_logo.jpg")
        self.logo_label = tk.Label(root, image=self.logo_image)
        self.logo_label.place(relx=0.5, rely=0.1, anchor=tk.CENTER)

        # Company name label
        self.company_name_label = tk.Label(root, text="MK Srinivasan Systems Private Limited", font=("Helvetica", 16, "bold"))
        self.company_name_label.place(relx=0.5, rely=0.2, anchor=tk.CENTER)

        # COM port selection label
        self.com_port_label = tk.Label(root, text="Select COM Port:", font=("Helvetica", 12))
        self.com_port_label.place(relx=0.3, rely=0.3, anchor=tk.CENTER)

        # Example COM port list (replace with your actual COM ports)
        com_ports = ["COM1", "COM2", "COM3", "COM4", "COM5"]

        # COM port selection combobox
        self.com_port_combobox = ttk.Combobox(root, values=com_ports, state="readonly")
        self.com_port_combobox.current(0)  # Set default selection
        self.com_port_combobox.place(relx=0.5, rely=0.3, anchor=tk.CENTER)

        # Data dump button
        self.data_dump_button = tk.Button(root, text="Data Dump", command=self.data_dump, bg="blue", fg="white", padx=50, pady=30, font=("Helvetica", 14, "bold"))
        self.data_dump_button.place(relx=0.3, rely=0.5, anchor=tk.CENTER)

        # Live loading button
        self.live_loading_button = tk.Button(root, text="Live Loading", command=self.live_loading, bg="green", fg="white", padx=50, pady=30, font=("Helvetica", 14, "bold"))
        self.live_loading_button.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

        # Stop button
        self.stop_button = tk.Button(root, text="Stop", command=self.stop, bg="red", fg="white", padx=50, pady=30, font=("Helvetica", 14, "bold"))
        self.stop_button.place(relx=0.7, rely=0.5, anchor=tk.CENTER)

        # Flag to control data receiving
        self.receive_data = True
        self.com_port = None  # Store selected COM port

    def data_dump(self):
        print("Data Dump button clicked")
        self.com_port = self.com_port_combobox.get()  # Get selected COM port
        # Start receiving data from the selected COM port
        self.start_data_reception()

    def live_loading(self):
        print("Live Loading button clicked")
        self.com_port = self.com_port_combobox.get()  # Get selected COM port
        # Start receiving data from the selected COM port
        self.start_data_reception()

    def stop(self):
        print("Stop button clicked")
        # Set flag to stop data receiving
        self.receive_data = False

    def start_data_reception(self):
        # Check if a COM port is selected
        if self.com_port:
            # Open the COM port for communication
            ser = serial.Serial(self.com_port, 9600)  # Adjust baud rate as needed
            while self.receive_data:
                # Read data from the COM port
                data = ser.readline().decode().strip()  # Decode bytes to string and remove whitespace
                if data:
                    print("Received Data:", data)
                    self.store_data_in_excel(data)  # Store received data in Excel
                else:
                    break  # Exit loop if no data received
            ser.close()  # Close the COM port after data reception is stopped
            print("Data receiving stopped")
        else:
            print("No COM port selected.")

    def store_data_in_excel(self, data):
        # Parse data and split it into columns
        lines = data.strip().split('\n')
        data_list = [line.split(',') for line in lines]
        print("Data List:", data_list)  # Print data list for debugging

        # Create a DataFrame using pandas
        df = pd.DataFrame(data_list, columns=["Time", "Voltage", "Current", "Resistance"])

        # Add timestamp to each row
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        df.insert(0, "Timestamp", timestamp)

        # Write DataFrame to Excel file
        filename = "data_dump1.xlsx"
        df.to_excel(filename, index=False)
        print(f"Data stored in {filename}")

if __name__ == "__main__":
    root = tk.Tk()
    app = CompanyInterface(root)

    # Set window size and position it in the middle of the screen
    window_width = 800
    window_height = 600
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    root.mainloop()

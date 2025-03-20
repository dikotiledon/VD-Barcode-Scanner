import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import pystray
import numpy as np
from PIL import Image
import threading
import serial
import serial.tools.list_ports
import time

class Application:
    def __init__(self, master):
        self.master = master
        self.master.title("Barcode Reader")
        self.master.geometry('600x800+100+100')
        self.master.protocol('WM_DELETE_WINDOW', self.minimize_to_tray)
        
        # Initialize serial connections
        self.input_serial = None
        self.output_serial = None
        self.serial_thread = None
        self.running = False

        # Setup UI Components
        self.setup_ui()

    def setup_ui(self):
        # Main frame
        main_frame = tk.Frame(self.master)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Buttons frame
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        # COM port selection frames
        com_ports_frame = tk.LabelFrame(main_frame, text="COM Port Configuration")
        com_ports_frame.pack(fill=tk.X, pady=10, padx=5)
        
        # Input COM port selection
        input_frame = tk.Frame(com_ports_frame)
        input_frame.pack(fill=tk.X, pady=5, padx=5)
        tk.Label(input_frame, text="Input (Barcode Scanner):").pack(side=tk.LEFT)
        
        # Output COM port selection
        output_frame = tk.Frame(com_ports_frame)
        output_frame.pack(fill=tk.X, pady=5, padx=5)
        tk.Label(output_frame, text="Output (I/O Controller):").pack(side=tk.LEFT)
        
        # Detect available COM ports
        com_ports = self.detect_com_ports()
        
        # Refresh COM ports button
        refresh_btn = tk.Button(
            com_ports_frame, 
            text="Refresh COM Ports", 
            command=self.refresh_com_ports
        )
        refresh_btn.pack(pady=5)
        
        # Minimize to tray Button
        minimize_to_tray_btn = tk.Button(
            button_frame, 
            text="Minimize to tray", 
            command=self.minimize_to_tray,
            bg="green",
            fg="white"
        )
        minimize_to_tray_btn.pack(side=tk.RIGHT, padx=1)
        
        # Start/Stop Button
        self.start_stop_btn = tk.Button(button_frame, text="Start Listening", command=self.toggle_listen)
        self.start_stop_btn.pack(side=tk.RIGHT, padx=1)
        
        # COM port dropdowns
        self.input_com_var = tk.StringVar()
        self.input_com_var.set(com_ports[0] if com_ports else "No COM Ports Found")
        self.input_com_dropdown = ttk.Combobox(input_frame, textvariable=self.input_com_var, values=com_ports, width=30)
        self.input_com_dropdown.pack(side=tk.LEFT, padx=5)
        
        self.output_com_var = tk.StringVar()
        self.output_com_var.set(com_ports[1] if len(com_ports) > 1 else com_ports[0] if com_ports else "No COM Ports Found")
        self.output_com_dropdown = ttk.Combobox(output_frame, textvariable=self.output_com_var, values=com_ports, width=30)
        self.output_com_dropdown.pack(side=tk.LEFT, padx=5)
        
        # Input serial configuration
        input_config_frame = tk.LabelFrame(com_ports_frame, text="Input Port Configuration")
        input_config_frame.pack(fill=tk.X, pady=5, padx=5)
        
        tk.Label(input_config_frame, text="Baud Rate:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.input_baud_var = tk.StringVar(value="9600")
        input_baud_combo = ttk.Combobox(input_config_frame, textvariable=self.input_baud_var, 
                                       values=["9600", "19200", "38400", "57600", "115200"], width=10)
        input_baud_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Output serial configuration
        output_config_frame = tk.LabelFrame(com_ports_frame, text="Output Port Configuration")
        output_config_frame.pack(fill=tk.X, pady=5, padx=5)
        
        tk.Label(output_config_frame, text="Baud Rate:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.output_baud_var = tk.StringVar(value="9600")
        output_baud_combo = ttk.Combobox(output_config_frame, textvariable=self.output_baud_var, 
                                        values=["9600", "19200", "38400", "57600", "115200"], width=10)
        output_baud_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        # Results Text Area with labels
        results_frame = tk.LabelFrame(main_frame, text="Communication Log")
        results_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.results_text = tk.Text(results_frame, height=20)
        self.results_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Add scrollbar
        scrollbar = tk.Scrollbar(self.results_text)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.results_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.results_text.yview)

        # Prepare full-screen overlay
        self.master_screen = tk.Toplevel(self.master)
        self.master_screen.withdraw()
        self.master_screen.attributes("-transparent", "maroon3")
        
        self.picture_frame = tk.Frame(self.master_screen, background="maroon3")
        self.picture_frame.pack(fill=tk.BOTH, expand=True)
    
    def detect_com_ports(self):
        """Detect all available COM ports (both physical and virtual)"""
        ports = serial.tools.list_ports.comports()
        port_list = []
        
        for port in ports:
            # Include detailed information about the port
            port_info = f"{port.device} - {port.description}"
            port_list.append(port_info)
            
        return port_list
    
    def refresh_com_ports(self):
        """Refresh the list of available COM ports"""
        com_ports = self.detect_com_ports()
        
        # Update dropdown values
        self.input_com_dropdown['values'] = com_ports
        self.output_com_dropdown['values'] = com_ports
        
        if com_ports:
            if self.input_com_var.get() == "No COM Ports Found":
                self.input_com_var.set(com_ports[0])
            if self.output_com_var.get() == "No COM Ports Found":
                self.output_com_var.set(com_ports[0])
        else:
            self.input_com_var.set("No COM Ports Found")
            self.output_com_var.set("No COM Ports Found")
            
        self.update_results("COM ports refreshed.\n")
    
    def update_results(self, message):
        # Use after method to update UI from thread safely
        self.master.after(0, self._safe_text_update, message)

    def _safe_text_update(self, message):
        # Append message to results text
        self.results_text.insert(tk.END, message)
        self.results_text.see(tk.END)
    
    def minimize_to_tray(self):
        self.master.withdraw() 
        # Load icon image, handle if file is missing
        try:
            image = Image.open("app.ico")
        except FileNotFoundError:
            # Create a simple default icon
            image = Image.new('RGB', (16, 16), color = 'blue')
            
        menu = (pystray.MenuItem('Quit', self.quit_window), 
                pystray.MenuItem('Show', self.show_window))
        self.icon = pystray.Icon("name", image, "Barcode Reader", menu)
        self.icon.run_detached()
    
    def quit_window(self, icon):
        self.stop_listening()
        icon.stop()
        self.master.destroy()    
    
    def show_window(self, icon):
        icon.stop()
        self.master.after(0, self.master.deiconify)  
    
    def toggle_listen(self):
        if not self.running:
            if self.start_listening():
                self.start_stop_btn.config(text="Stop Listening")
                self.running = True
        else:
            self.stop_listening()
            self.start_stop_btn.config(text="Start Listening")
            self.running = False
    
    def extract_port_name(self, port_string):
        """Extract the actual COM port name from the dropdown string"""
        return port_string.split(' - ')[0]
    
    def start_listening(self):
        # Get selected port names
        input_port_full = self.input_com_var.get()
        output_port_full = self.output_com_var.get()
        
        if input_port_full == "No COM Ports Found" or output_port_full == "No COM Ports Found":
            messagebox.showerror("Error", "Please select valid COM ports.")
            return False
        
        # Extract actual port names
        input_port = self.extract_port_name(input_port_full)
        output_port = self.extract_port_name(output_port_full)
        
        # Get baud rates
        input_baud = int(self.input_baud_var.get())
        output_baud = int(self.output_baud_var.get())
        
        try:
            # Open input serial port
            self.input_serial = serial.Serial(input_port, baudrate=input_baud, timeout=1)
            self.update_results(f"Connected to input port {input_port} at {input_baud} baud.\n")
            
            # Open output serial port
            self.output_serial = serial.Serial(output_port, baudrate=output_baud, timeout=1)
            self.update_results(f"Connected to output port {output_port} at {output_baud} baud.\n")
            
            # Start reading thread
            self.serial_thread = threading.Thread(target=self.read_serial_data)
            self.serial_thread.daemon = True
            self.serial_thread.start()
            
            return True
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open COM ports: {e}")
            self.stop_listening()
            return False
    
    def stop_listening(self):
        self.running = False
        
        # Close input serial port
        if hasattr(self, 'input_serial') and self.input_serial and self.input_serial.is_open:
            self.input_serial.close()
            self.update_results("Input port closed.\n")
            
        # Close output serial port
        if hasattr(self, 'output_serial') and self.output_serial and self.output_serial.is_open:
            self.output_serial.close()
            self.update_results("Output port closed.\n")
    
    def read_serial_data(self):
        self.update_results("Listening for barcode scans...\n")
        
        while self.running:
            try:
                # Read data from input port (barcode scanner)
                if self.input_serial and self.input_serial.is_open:
                    data = self.input_serial.readline().decode('utf-8', errors='replace').strip()
                    
                    if data:
                        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
                        log_message = f"[{timestamp}] Received: {data}\n"
                        self.update_results(log_message)
                        
                        # Forward data to output port (I/O controller)
                        if self.output_serial and self.output_serial.is_open:
                            try:
                                # Add carriage return and line feed to ensure proper command termination
                                self.output_serial.write((data + "\r\n").encode('utf-8'))
                                self.update_results(f"Sent to I/O controller: {data}\n")
                            except Exception as e:
                                self.update_results(f"Error sending to output port: {e}\n")
                
                # Small delay to prevent CPU hogging
                time.sleep(0.01)
                
            except Exception as e:
                self.update_results(f"Error in serial communication: {e}\n")
                # If there's an error in the loop, wait a bit before retrying
                time.sleep(1)
                
            # Check if we should exit the loop
            if not self.running:
                break

def main():
    root = tk.Tk()
    app = Application(root)
    root.mainloop()

if __name__ == '__main__':
    main()
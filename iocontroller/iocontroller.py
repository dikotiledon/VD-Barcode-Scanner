import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import pystray
import numpy as np
from PIL import Image
import threading
import serial
import serial.tools.list_ports
import time
import keyboard  # Add at top of file
import sys
import os

class Application:
    def __init__(self, master):
        self.master = master
        self.master.title("Barcode Scanner Interface")
        self.master.geometry('600x800+100+100')
        self.master.protocol('WM_DELETE_WINDOW', self.on_closing)
        
        # Create main frames
        self.control_frame = ttk.Frame(self.master)
        self.control_frame.pack(fill="x", padx=5, pady=5)
        
        # Initialize barcode length variable
        self.barcode_length = tk.StringVar(value="14")
        self.switch_number = tk.StringVar(value="1")
        
        # Create UI elements
        self.create_barcode_length_input()
        
        # Initialize serial connections
        self.input_serial = None
        self.output_serial = None
        self.serial_thread = None
        self.running = False
        self.shutdown_flag = threading.Event()  # Add flag for clean thread termination
        self.interface_type = 'serial'  # or 'usb'
        self.barcode_buffer = ''
        self.available_ports = []
        self.scan_com_ports()
        self.last_barcode = ''

        # Setup UI Components
        self.setup_ui()
        
        # Initialize tray icon object
        self.icon = None
        self.icon_thread = None

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
        try:
            self.results_text.insert(tk.END, message)
            self.results_text.see(tk.END)
            
            # Limit log size for performance (keep last 1000 lines)
            line_count = int(self.results_text.index('end-1c').split('.')[0])
            if line_count > 1000:
                self.results_text.delete('1.0', f'{line_count-1000}.0')
        except tk.TclError:
            # Handle case when widget is destroyed
            pass
    
    def minimize_to_tray(self):
        """Minimize window to system tray"""
        self.master.withdraw()
        if not self.icon:
            image = Image.new('RGB', (64, 64), color='red')
            menu = pystray.Menu(
                pystray.MenuItem("Show", self.show_window),
                pystray.MenuItem("Exit", lambda: self.master.after(0, self.on_closing))
            )
            self.icon = pystray.Icon("name", image, "Barcode Scanner", menu)
            self.icon_thread = threading.Thread(target=self.run_icon, daemon=True)
            self.icon_thread.start()

    def quit_window(self, icon=None):
        # First stop listening to clean up resources
        self.stop_listening()
        
        # Signal thread to exit
        self.shutdown_flag.set()
        
        # Stop the tray icon if it exists
        if icon is not None:
            try:
                icon.stop()
            except Exception:
                pass
        elif self.icon is not None:
            try:
                self.icon.stop()
            except Exception:
                pass
        
        # Destroy all Tkinter windows
        try:
            if self.master_screen:
                self.master_screen.destroy()
            self.master.quit()
            self.master.destroy()
        except Exception:
            pass
            
        # Forcibly exit the application after a short delay
        self.master.after(100, lambda: os._exit(0))
    
    def show_window(self, icon):
        if icon:
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
            
            # Reset shutdown flag
            self.shutdown_flag.clear()
            
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
        # Send OFF commands before closing
        if hasattr(self, 'output_serial') and self.output_serial and self.output_serial.is_open:
            try:
                # Turn off switch 1
                self.output_serial.write(('@OFF01$' + "\r\n").encode('utf-8'))
                self.update_results("Sent to I/O controller: @OFF01$\n")
                time.sleep(0.1)
                
                # Turn off switch 2
                self.output_serial.write(('@OFF02$' + "\r\n").encode('utf-8'))
                self.update_results("Sent to I/O controller: @OFF02$\n")
            except Exception as e:
                self.update_results(f"Error sending OFF commands: {e}\n")

        self.running = False
        self.shutdown_flag.set()  # Signal thread to terminate
        
        # Close input serial port
        if hasattr(self, 'input_serial') and self.input_serial and self.input_serial.is_open:
            try:
                self.input_serial.close()
                self.update_results("Input port closed.\n")
            except Exception:
                pass
            
        # Close output serial port
        if hasattr(self, 'output_serial') and self.output_serial and self.output_serial.is_open:
            try:
                self.output_serial.close()
                self.update_results("Output port closed.\n")
            except Exception:
                pass
                
        # Wait for thread to terminate (with timeout)
        if self.serial_thread and self.serial_thread.is_alive():
            self.serial_thread.join(timeout=1.0)
    
    def read_serial_data(self):
        self.update_results("Listening for barcode scans...\n")
        
        while not self.shutdown_flag.is_set() and self.running:
            try:
                # Read data from input port (barcode scanner)
                if self.input_serial and self.input_serial.is_open:
                    # Check if data is available (non-blocking)
                    if self.input_serial.in_waiting > 0:
                        data = self.input_serial.readline().decode('utf-8', errors='replace').strip()
                        
                        if data:
                            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
                            log_message = f"[{timestamp}] Received: {data}\n"                        
                            self.update_results(log_message)
                            
                            # Forward data to output port (I/O controller)
                            if self.output_serial and self.output_serial.is_open:
                                try:
                                    if self.validate_barcode(data):
                                        if data.strip() != self.last_barcode:
                                            command = '@ON01$' if int(self.switch_number.get()) == 1 else '@ON02$'
                                            self.output_serial.write((command + "\r\n").encode('utf-8'))
                                            self.update_results(f"Sent to I/O controller: {command}\n")
                                            self.last_barcode = data.strip()
                                            time.sleep(2.0)
                                            self.output_serial.write((command.replace('ON', 'OFF') + "\r\n").encode('utf-8'))
                                    else:
                                        self.output_serial.write(('@OFF01$' + "\r\n").encode('utf-8'))
                                        self.output_serial.write(('@OFF02$' + "\r\n").encode('utf-8'))
                                        self.update_results(f"Sent to I/O controller: @OFF01$\n")
                                        # Use after method to safely call UI methods from thread
                                        self.master.after(0, self.toggle_listen)
                                        return
                                except Exception as e:
                                    self.update_results(f"Error sending to output port: {e}\n")
                
                # Small delay to prevent CPU hogging (more efficient than original)
                time.sleep(0.05)
                
            except Exception as e:
                self.update_results(f"Error in serial communication: {e}\n")
                # If there's an error in the loop, wait a bit before retrying
                time.sleep(1)
                
            # Check if we should exit the loop
            if not self.running or self.shutdown_flag.is_set():
                break
                
        self.update_results("Serial monitoring thread stopped.\n")

    def switch_to_usb_mode(self):
        """Switch input interface to USB barcode scanner"""
        self.interface_type = 'usb'
        # Use a try/except block to safely handle keyboard hook
        try:
            keyboard.on_press(self.handle_usb_input)
            self.update_results("Switched to USB barcode scanner mode\n")
        except Exception as e:
            self.update_results(f"Error setting up USB mode: {e}\n")
    
    def handle_usb_input(self, event):
        """Handle input from USB barcode scanner"""
        if not self.running:
            return
            
        if event.name == 'enter':
            # Process complete barcode
            if self.barcode_buffer:
                self.process_data(self.barcode_buffer)
                self.barcode_buffer = ''
        elif len(event.name) == 1:  # Single character
            self.barcode_buffer += event.name
    
    def process_data(self, data):
        """Process received data regardless of interface"""
        try:
            # Your existing data processing logic
            self.update_results(f"Received data: {data}\n")
            # Process and send to output port
            if self.output_serial and self.output_serial.is_open:
                self.output_serial.write(data.encode())
                self.update_results(f"Sent to I/O controller: {data}\n")
        except Exception as e:
            self.update_results(f"Error processing data: {e}\n")

    def scan_com_ports(self):
        """Scan and list all available COM ports including virtual ones"""
        try:
            # Get list of all COM ports
            self.available_ports = list(serial.tools.list_ports.comports())
            ports_info = "Available COM ports:\n"
            for port in self.available_ports:
                ports_info += f"Port: {port.device} - {port.description}\n"
            self.update_results(ports_info)
        except Exception as e:
            self.update_results(f"Error scanning COM ports: {e}\n")

    def refresh_ports(self):
        """Refresh the list of available COM ports"""
        self.scan_com_ports()
        # Update combobox or list with new ports
        if hasattr(self, 'port_select'):
            current_ports = [port.device for port in self.available_ports]
            self.port_select['values'] = current_ports
            if current_ports:
                self.port_select.set(current_ports[0])

    def connect_serial(self):
        """Connect to selected serial port"""
        try:
            port = self.port_select.get()
            self.ser = serial.Serial(
                port=port,
                baudrate=9600,
                timeout=1
            )
            self.update_results(f"Connected to {port}\n")
        except Exception as e:
            self.update_results(f"Error connecting to port: {e}\n")

    def create_barcode_length_input(self):
        """Create input field for barcode length"""
        length_frame = ttk.LabelFrame(self.control_frame, text="Barcode Settings")
        length_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(length_frame, text="Expected Length:").pack(side="left", padx=5)
        length_entry = ttk.Entry(length_frame, textvariable=self.barcode_length, width=5)
        length_entry.pack(side="left", padx=5)
        
        ttk.Label(length_frame, text="Switch Number:").pack(side="left", padx=5)

        # Create a Combobox for switch number
        switch_combobox = ttk.Combobox(length_frame, textvariable=self.switch_number, values=list(range(1, 3)), width=3)
        switch_combobox.pack(side="left", padx=5)
        switch_combobox.current(0) # set default value to 1.

        # Add validation for length_entry
        vcmd = (self.master.register(self.validate_length_input), '%P')
        length_entry.configure(validate='key', validatecommand=vcmd)
        
        vcmd2 = (self.master.register(self.validate_switch_input), '%P')
        switch_combobox.configure(validate='key', validatecommand=vcmd2)

    def validate_length_input(self, value):
        """Validate that length input is numeric"""
        if value == "":
            return True
        try:
            int_value = int(value)
            return int_value > 0 and int_value <= 100
        except ValueError:
            return False
            
    def validate_switch_input(self, value):
        """Validate that switch input is numeric"""
        if value == "":
            return True
        try:
            int_value = int(value)
            return int_value > 0 and int_value <= 2
        except ValueError:
            return False        

    def validate_barcode(self, data):
        """Validate barcode length against user setting"""
        try:
            expected_length = int(self.barcode_length.get())
            return len(data.strip()) == expected_length
        except ValueError:
            self.update_results("Invalid barcode length setting\n")
            return False

    def cleanup(self):
        """Clean up resources before closing"""
        # Stop the serial thread
        if self.running:
            self.running = False
            self.shutdown_flag.set()
            if self.serial_thread and self.serial_thread.is_alive():
                self.serial_thread.join(timeout=2.0)

        # Close serial connections
        if self.input_serial and self.input_serial.is_open:
            self.input_serial.close()
        if self.output_serial and self.output_serial.is_open:
            self.output_serial.close()

        # Remove system tray icon if it exists
        if self.icon:
            self.icon.stop()
            if self.icon_thread and self.icon_thread.is_alive():
                self.icon_thread.join(timeout=2.0)

    def run_icon(self):
        """Run system tray icon in separate thread"""
        self.icon.run()

    def on_closing(self):
        """Handle window closing event"""
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            self.cleanup()
            self.master.quit()
            self.master.destroy()
            os._exit(0)  # Force exit if needed

def main():
    root = tk.Tk()
    app = Application(root)
    root.mainloop()

if __name__ == '__main__':
    main()
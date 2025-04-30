import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import pystray
import numpy as np
import threading
import serial
import serial.tools.list_ports
import time
import keyboard  # Add at top of file
import sys
import os
from PIL import Image  # Add this import at the top if not present
from config_manager import ConfigManager

class Application:
    def __init__(self, master):
        self.master = master
        self.master.title("Barcode Scanner")
        self.master.geometry('700x500+100+100')
        self.master.protocol('WM_DELETE_WINDOW', self.on_closing)
        self.printer_serial = None
        
        # Set window icon
        try:
            if getattr(sys, 'frozen', False):
                # Running as compiled exe
                application_path = sys._MEIPASS
            else:
                # Running as script
                application_path = os.path.dirname(os.path.abspath(__file__))
                
            icon_path = os.path.join(application_path, "app.ico")
            if os.path.exists(icon_path):
                self.master.iconbitmap(icon_path)
                # Store icon path for tray icon
                self.icon_path = icon_path
            else:
                self.icon_path = None
        except Exception:
            self.icon_path = None
        
        # Create main frames
        self.control_frame = ttk.Frame(self.master)
        self.control_frame.pack(fill="x", padx=5, pady=5)
        
        # Initialize barcode length variable
        self.barcode_length = tk.StringVar(value="14")
        self.switch_number = tk.StringVar(value="1")
        self.barcode_list = []  # Initialize list to store barcodes
        
        # Create UI elements
        # self.create_barcode_length_input()
        
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
        
        self.config_manager = ConfigManager()
        self.config = self.config_manager.load_config()
        self.input_serials = [None, None]  # Two input serial connections

        # Setup UI Components
        self.setup_ui()
        
        # Initialize tray icon object
        self.icon = None
        self.icon_thread = None

    def setup_ui(self):
        # Create main container with background color
        self.container_frame = tk.Frame(self.master)
        self.container_frame.pack(fill=tk.BOTH, expand=True)

        # Title bar at top (already exists in __init__)
        # self.master.title("Barcode Scanner")

        # Main content frame
        main_frame = tk.Frame(self.container_frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10,0))

        # Buttons frame
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        # Add Start/Stop and Minimize buttons
        self.start_stop_btn = tk.Button(
            button_frame, 
            text="Start Listening", 
            command=self.toggle_listen,
            bg="green",
            fg="white"
        )
        self.start_stop_btn.pack(side=tk.LEFT, padx=5)
        # Detect available COM ports first
        com_ports = self.detect_com_ports()
        minimize_to_tray_btn = tk.Button(
            button_frame, 
            text="Minimize to Tray", 
            command=self.minimize_to_tray,
            bg="grey",
            fg="white"
        )
        minimize_to_tray_btn.pack(side=tk.RIGHT, padx=5)
        
        # COM port selection frames
        com_ports_frame = tk.LabelFrame(main_frame, text="COM Port Configuration")
        com_ports_frame.pack(fill=tk.X, pady=10, padx=5)
        
        # Left side: COM Port Selection
        ports_frame = tk.LabelFrame(com_ports_frame, text="Port Selection")
        ports_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Input Port 1
        input1_frame = tk.Frame(ports_frame)
        input1_frame.pack(fill=tk.X, pady=2)
        tk.Label(input1_frame, text="Input 1:").pack(side=tk.LEFT)
        self.input1_com_var = tk.StringVar(value=self.config['input_port1'])
        self.input1_com_dropdown = ttk.Combobox(input1_frame, textvariable=self.input1_com_var, 
                                              values=com_ports, width=30)
        self.input1_com_dropdown.pack(side=tk.LEFT, padx=5)
        
        # Input Port 2
        input2_frame = tk.Frame(ports_frame)
        input2_frame.pack(fill=tk.X, pady=2)
        tk.Label(input2_frame, text="Input 2:").pack(side=tk.LEFT)
        self.input2_com_var = tk.StringVar(value=self.config['input_port2'])
        self.input2_com_dropdown = ttk.Combobox(input2_frame, textvariable=self.input2_com_var, 
                                              values=com_ports, width=30)
        self.input2_com_dropdown.pack(side=tk.LEFT, padx=5)
        
        # Output Port
        output_frame = tk.Frame(ports_frame)
        output_frame.pack(fill=tk.X, pady=2)
        tk.Label(output_frame, text="Output: ").pack(side=tk.LEFT)
        self.output_com_var = tk.StringVar(value=self.config['output_port'])
        self.output_com_dropdown = ttk.Combobox(output_frame, textvariable=self.output_com_var, 
                                              values=com_ports, width=30)
        self.output_com_dropdown.pack(side=tk.LEFT, padx=5)

        # Add Printer Port
        printer_frame = tk.Frame(ports_frame)
        printer_frame.pack(fill=tk.X, pady=2)
        tk.Label(printer_frame, text="Printer:").pack(side=tk.LEFT)
        self.printer_com_var = tk.StringVar(value=self.config['printer_port'])
        self.printer_com_dropdown = ttk.Combobox(printer_frame, textvariable=self.printer_com_var, 
                                               values=com_ports, width=30)
        self.printer_com_dropdown.pack(side=tk.LEFT, padx=5)
        
        # Right side: Port Configurations
        config_frame = tk.LabelFrame(com_ports_frame, text="Port Configuration")
        config_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Input 1 configuration
        input1_config_frame = tk.Frame(config_frame)
        input1_config_frame.pack(fill=tk.X, pady=2)
        tk.Label(input1_config_frame, text="Input 1 Baud:").pack(side=tk.LEFT)
        self.input1_baud_var = tk.StringVar(value=self.config['input_baud1'])
        ttk.Combobox(input1_config_frame, textvariable=self.input1_baud_var, 
                     values=["9600", "19200", "38400", "57600", "115200"], width=7).pack(side=tk.LEFT, padx=5)
        
        # Input 2 configuration
        input2_config_frame = tk.Frame(config_frame)
        input2_config_frame.pack(fill=tk.X, pady=2)
        tk.Label(input2_config_frame, text="Input 2 Baud:").pack(side=tk.LEFT)
        self.input2_baud_var = tk.StringVar(value=self.config['input_baud2'])
        ttk.Combobox(input2_config_frame, textvariable=self.input2_baud_var, 
                     values=["9600", "19200", "38400", "57600", "115200"], width=7).pack(side=tk.LEFT, padx=5)
        
        # Output configuration
        output_config_frame = tk.Frame(config_frame)
        output_config_frame.pack(fill=tk.X, pady=2)
        tk.Label(output_config_frame, text="Output Baud:").pack(side=tk.LEFT)
        self.output_baud_var = tk.StringVar(value=self.config['output_baud'])
        ttk.Combobox(output_config_frame, textvariable=self.output_baud_var, 
                     values=["9600", "19200", "38400", "57600", "115200"], width=7).pack(side=tk.LEFT, padx=5)
        
        # Switch selection
        switch_frame = tk.Frame(config_frame)
        switch_frame.pack(fill=tk.X, pady=2)
        tk.Label(switch_frame, text="Switch:").pack(side=tk.LEFT)
        self.switch_number = tk.StringVar(value=self.config['switch_number'])
        ttk.Combobox(switch_frame, textvariable=self.switch_number, 
                     values=["1", "2"], width=7).pack(side=tk.LEFT, padx=5)
        
        # Add Printer Configuration
        printer_config_frame = tk.Frame(config_frame)
        printer_config_frame.pack(fill=tk.X, pady=2)
        tk.Label(printer_config_frame, text="Printer Baud:").pack(side=tk.LEFT)
        self.printer_baud_var = tk.StringVar(value=self.config['printer_baud'])
        ttk.Combobox(printer_config_frame, textvariable=self.printer_baud_var, 
                     values=["9600", "19200", "38400", "57600", "115200"], width=7).pack(side=tk.LEFT, padx=5)
        
        # Refresh button at the bottom
        refresh_btn = tk.Button(com_ports_frame, text="Refresh Ports", command=self.refresh_com_ports)
        refresh_btn.pack(pady=5)
        
        # Create horizontal layout frame
        horizontal_frame = tk.Frame(main_frame)
        horizontal_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Left side: Input barcodes
        input_frame = tk.LabelFrame(horizontal_frame, text="Input Barcodes")
        input_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # Add buttons frame above text box
        input_buttons_frame = tk.Frame(input_frame)
        input_buttons_frame.pack(fill=tk.X, padx=5, pady=2)
        
        # Add move up/down/clear/delete buttons
        self.move_up_btn = tk.Button(
            input_buttons_frame,
            text="▲",
            command=self.move_barcode_up,
            width=2
        )
        self.move_up_btn.pack(side=tk.LEFT, padx=2)
        
        self.move_down_btn = tk.Button(
            input_buttons_frame,
            text="▼",
            command=self.move_barcode_down,
            width=2
        )
        self.move_down_btn.pack(side=tk.LEFT, padx=2)
        
        self.delete_btn = tk.Button(
            input_buttons_frame,
            text="Delete Selected",
            command=self.delete_selected_input_barcode,
            bg="orange",
            fg="white"
        )
        self.delete_btn.pack(side=tk.RIGHT, padx=2)
        
        # Input text box (read-only)
        self.input_text = tk.Text(input_frame, width=30, height=10, wrap=tk.WORD, state="disabled", font=("Arial", 16))
        self.input_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Add scrollbar for input
        input_scrollbar = tk.Scrollbar(input_frame)
        input_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.input_text.config(yscrollcommand=input_scrollbar.set)
        input_scrollbar.config(command=self.input_text.yview)
        
        # Bind validation to input text
        self.input_text.bind('<KeyRelease>', self.validate_input_text)
        
        # Right side: Master barcode list
        master_frame = tk.LabelFrame(horizontal_frame, text="Master Barcodes")
        master_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Add buttons frame above master text box
        master_buttons_frame = tk.Frame(master_frame)
        master_buttons_frame.pack(fill=tk.X, padx=5, pady=2)
        
        # Add move up/down/clear/delete buttons for master list
        self.master_up_btn = tk.Button(
            master_buttons_frame,
            text="▲",
            command=self.move_master_up,
            width=2
        )
        self.master_up_btn.pack(side=tk.LEFT, padx=2)
        
        self.master_down_btn = tk.Button(
            master_buttons_frame,
            text="▼",
            command=self.move_master_down,
            width=2
        )
        self.master_down_btn.pack(side=tk.LEFT, padx=2)
        
        self.master_delete_btn = tk.Button(
            master_buttons_frame,
            text="Delete Selected",
            command=self.delete_selected_master_barcode,
            bg="orange",
            fg="white"
        )
        self.master_delete_btn.pack(side=tk.RIGHT, padx=2)
        
        self.master_text = tk.Text(master_frame, width=30, height=10, bg='lightgray', state="disabled", font=("Arial", 16))
        self.master_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Add scrollbar for master list
        master_scrollbar = tk.Scrollbar(master_frame)
        master_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.master_text.config(yscrollcommand=master_scrollbar.set)
        master_scrollbar.config(command=self.master_text.yview)
        
        # Configure tags for coloring
        self.master_text.tag_configure('matched', background='light green')
        self.master_text.tag_configure('unmatched', background='lightgray')
        
        # Add show/hide log button
        self.log_visible = False
        self.toggle_log_btn = tk.Button(
            main_frame,
            text="Show Log",
            command=self.toggle_log_visibility
        )
        self.toggle_log_btn.pack(fill=tk.X, pady=5)
        
        # Communication log frame (hidden by default)
        self.log_frame = tk.LabelFrame(main_frame, text="Communication Log")
        self.results_text = tk.Text(self.log_frame, height=10)
        self.results_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Add scrollbar for log
        log_scrollbar = tk.Scrollbar(self.results_text)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.results_text.config(yscrollcommand=log_scrollbar.set)
        log_scrollbar.config(command=self.results_text.yview)

        # Prepare full-screen overlay
        self.master_screen = tk.Toplevel(self.master)
        self.master_screen.withdraw()
        self.master_screen.attributes("-transparent", "maroon3")
        
        self.picture_frame = tk.Frame(self.master_screen, background="maroon3")
        self.picture_frame.pack(fill=tk.BOTH, expand=True)
        
        # Remove both old footer implementations and replace with this:
        # Create footer frame that sticks to bottom
        self.footer_frame = tk.Frame(self.master, bg='lightgray', height=25)
        self.footer_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.footer_frame.pack_propagate(False)

        # Add separator above footer
        separator = ttk.Separator(self.master, orient='horizontal')
        separator.pack(side=tk.BOTTOM, fill=tk.X)

        # Copyright text
        copyright_label = tk.Label(
            self.footer_frame,
            text="© 2025 IT & System Innovation P",
            font=("Arial", 8),
            bg='lightgray',
            fg='black'
        )
        copyright_label.pack(side=tk.RIGHT, padx=10, pady=5)

        # Make sure footer stays on top of other widgets
        self.footer_frame.lift()

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
        
        # Update all dropdown values
        self.input1_com_dropdown['values'] = com_ports
        self.input2_com_dropdown['values'] = com_ports
        self.output_com_dropdown['values'] = com_ports
        self.printer_com_dropdown['values'] = com_ports
        
        if com_ports:
            if self.input1_com_var.get() == "No COM Ports Found":
                self.input1_com_var.set(com_ports[0])
            if self.input2_com_var.get() == "No COM Ports Found":
                self.input2_com_var.set(com_ports[0])
            if self.output_com_var.get() == "No COM Ports Found":
                self.output_com_var.set(com_ports[0])
            if self.printer_com_var.get() == "No COM Ports Found":
                self.printer_com_var.set(com_ports[0])
        else:
            self.input1_com_var.set("No COM Ports Found")
            self.input2_com_var.set("No COM Ports Found")
            self.output_com_var.set("No COM Ports Found")
            self.printer_com_var.set("No COM Ports Found")
            
        self.update_results("COM ports refreshed.\n")
    
    def update_results(self, message):
        """Update the communication log with a timestamp"""
        # Add a readable timestamp to the message
        timestamp = time.strftime("[%Y-%m-%d %H:%M:%S]")
        formatted_message = f"{timestamp} {message}"
        
        # Use after method to update UI from thread safely
        self.master.after(0, self._safe_text_update, formatted_message)

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
        self.master.withdraw()  # Hide the main window
        if not self.icon:
            try:
                # Create image for tray icon
                if self.icon_path and os.path.exists(self.icon_path):
                    image = Image.open(self.icon_path)
                else:
                    # Fallback to a colored square if no icon file
                    image = Image.new('RGB', (64, 64), color='red')
                    
                # Create menu for tray icon
                menu = (
                    pystray.MenuItem("Show", self.show_window),
                    pystray.MenuItem("Exit", self.quit_window)
                )
                
                # Create the tray icon
                self.icon = pystray.Icon(
                    name="BarcodeScanner",
                    icon=image,
                    title="Barcode Scanner",
                    menu=pystray.Menu(*menu)
                )
                
                # Start the icon in a separate thread
                self.icon_thread = threading.Thread(target=self.run_icon, daemon=True)
                self.icon_thread.start()
                
            except Exception as e:
                self.update_results(f"Error creating tray icon: {e}\n")
                # Restore window if tray icon fails
                self.master.deiconify()

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
    
    def show_window(self, icon=None):
        """Show the main window and stop the tray icon"""
        if self.icon:
            self.icon.stop()
            self.icon = None
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
        try:
            # Reset all serial connections first
            self.input_serials = [None, None]
            self.output_serial = None
            self.printer_serial = None
            
            # Open output port first
            output_port = self.extract_port_name(self.output_com_var.get())
            if output_port and output_port != "None":
                self.output_serial = serial.Serial(
                    output_port,
                    baudrate=int(self.output_baud_var.get()),
                    timeout=1
                )
                self.update_results(f"Connected to output port: {output_port}\n")
                    # Open printer port
            printer_port = self.extract_port_name(self.printer_com_var.get())
            if printer_port and printer_port != "None":
                self.printer_serial = serial.Serial(
                    printer_port,
                    baudrate=int(self.printer_baud_var.get()),
                    timeout=1
                )
                self.update_results(f"Connected to printer port: {printer_port}\n")
            
            # Open input ports
            input1_port = self.extract_port_name(self.input1_com_var.get())
            if input1_port and input1_port != "None":
                self.input_serials[0] = serial.Serial(
                    input1_port,
                    baudrate=int(self.input1_baud_var.get()),
                    timeout=1
                )
                self.update_results(f"Connected to input port 1: {input1_port}\n")
            
            input2_port = self.extract_port_name(self.input2_com_var.get())
            if input2_port and input2_port != "None":
                self.input_serials[1] = serial.Serial(
                    input2_port,
                    baudrate=int(self.input2_baud_var.get()),
                    timeout=1
                )
                self.update_results(f"Connected to input port 2: {input2_port}\n")
            
            # Reset flags
            self.shutdown_flag.clear()
            self.running = True
            
            # Save configuration
            self.save_configuration()
            
            # Start reading thread only if at least one input port is open
            if any(serial_conn is not None for serial_conn in self.input_serials):
                self.serial_thread = threading.Thread(target=self.read_serial_data)
                self.serial_thread.daemon = True
                self.serial_thread.start()
                return True
            else:
                messagebox.showerror("Error", "No input ports configured")
                return False
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open ports: {e}")
            self.stop_listening()
            return False
    
    def save_configuration(self):
        """Save current configuration"""
        config = {
            'input_port1': self.input1_com_var.get(),
            'input_port2': self.input2_com_var.get(),
            'input_baud1': self.input1_baud_var.get(),
            'input_baud2': self.input2_baud_var.get(),
            'output_port': self.output_com_var.get(),
            'output_baud': self.output_baud_var.get(),
            'printer_port': self.printer_com_var.get(),
            'printer_baud': self.printer_baud_var.get(),
            'switch_number': self.switch_number.get()
        }
        self.config_manager.save_config(config)
    
    def stop_listening(self):
        """Stop listening and properly close all COM ports"""
        try:
            # Signal thread to terminate first
            self.running = False
            self.shutdown_flag.set()

            # Wait for thread to terminate
            if self.serial_thread and self.serial_thread.is_alive():
                self.serial_thread.join(timeout=1.0)
                self.serial_thread = None

            # Send OFF commands before closing
            if hasattr(self, 'output_serial') and self.output_serial and self.output_serial.is_open:
                try:
                    self.output_serial.write(('@OFF01$' + "\r\n").encode('utf-8'))
                    self.output_serial.write(('@OFF02$' + "\r\n").encode('utf-8'))
                    self.update_results("Sent OFF commands to I/O controller\n")
                    time.sleep(0.1)
                except Exception as e:
                    self.update_results(f"Error sending OFF commands: {e}\n")

            # Close all input serial ports
            for i, serial_conn in enumerate(self.input_serials):
                if serial_conn and serial_conn.is_open:
                    try:
                        serial_conn.close()
                        self.input_serials[i] = None
                        self.update_results(f"Input port {i+1} closed\n")
                    except Exception as e:
                        self.update_results(f"Error closing input port {i+1}: {e}\n")

            # Close output serial port
            if hasattr(self, 'output_serial') and self.output_serial:
                try:
                    if self.output_serial.is_open:
                        self.output_serial.close()
                    self.output_serial = None
                    self.update_results("Output port closed\n")
                except Exception as e:
                    self.update_results(f"Error closing output port: {e}\n")

            # Close printer port
            if hasattr(self, 'printer_serial') and self.printer_serial:
                try:
                    if self.printer_serial.is_open:
                        self.printer_serial.close()
                    self.printer_serial = None
                    self.update_results("Printer port closed\n")
                except Exception as e:
                    self.update_results(f"Error closing printer port: {e}\n")

        except Exception as e:
            self.update_results(f"Error in stop_listening: {e}\n")
    
    def read_serial_data(self):
        """Read data from both input ports with different behaviors"""
        self.update_results("Listening for barcode scans...\n")
        
        while not self.shutdown_flag.is_set() and self.running:
            try:
                # Check input ports
                for i, serial_conn in enumerate(self.input_serials):
                    if serial_conn and serial_conn.is_open and serial_conn.in_waiting > 0:
                        # Read data from the serial port
                        raw_data = serial_conn.readline().decode('utf-8', errors='replace').strip()
                        
                        # Remove any prefix or suffix (customize this logic as needed)
                        data = self.clean_data(raw_data)
                        
                        if data:
                            if i == 0:  # Input 1 - Master Barcodes only
                                self.update_results(f"Master Input Received: {data}\n")
                                self.add_to_master_list(data)
                                
                                # Send data to the Printer COM port with STX and ETX
                                if self.printer_serial and self.printer_serial.is_open:
                                    try:
                                        # Add STX and ETX to the data
                                        formatted_data = f"\x02{data}\x03"
                                        self.printer_serial.write(f"{formatted_data}\r\n".encode('utf-8'))
                                        self.update_results(f"Sent to Printer: {formatted_data}\n")
                                    except Exception as e:
                                        self.update_results(f"Error sending to Printer: {e}\n")
                            elif i == 1:  # Input 2 - Validation against master list
                                self.update_results(f"Input Received: {data}\n")
                                is_valid = self.validate_barcode(data)
                                self.add_to_input_barcodes(data, is_valid)
                                
                time.sleep(0.05)
                
            except Exception as e:
                self.update_results(f"Error in serial communication: {e}\n")
                time.sleep(1)
            

    def validate_barcode(self, data):
        """Validate barcode against master list in order"""
        try:
            data = data.strip()
            master_barcodes = self.master_text.get("1.0", tk.END).splitlines()
            master_barcodes = [b.strip() for b in master_barcodes if b.strip()]
            
            if not master_barcodes:
                self.master.after(0, lambda: self.show_alert_popup(False, "No master barcodes defined"))
                return False
                
            # Check if barcode matches the first unmatched barcode in master list
            for i, barcode in enumerate(master_barcodes, 1):
                if barcode.strip() == data:
                    # Check if this is the first unmatched barcode
                    if not self.is_any_unmatched_barcode_before(i):
                        # Success case
                        if data.strip() != self.last_barcode:
                            # Process successful scan
                            command = '@ON01$' if int(self.switch_number.get()) == 1 else '@ON02$'
                            if self.output_serial and self.output_serial.is_open:
                                self.output_serial.write((command + "\r\n").encode('utf-8'))
                                self.last_barcode = data.strip()
                                time.sleep(2.0)
                                self.output_serial.write((command.replace('ON', 'OFF') + "\r\n").encode('utf-8'))
                            
                            # Remove the matched barcode from the master list
                            self.master.after(0, lambda i=i: self.remove_matched_master_barcode(i))
                            
                            # Show success popup
                            self.master.after(0, lambda: self.show_alert_popup(True, 
                                f"Barcode: {data}\nCommand: {command}"))
                            return True
                    else:
                        self.master.after(0, lambda: self.show_alert_popup(False, 
                            f"Incorrect order. Please scan previous barcodes first."))
                        return False
            
            self.master.after(0, lambda: self.show_alert_popup(False, 
                f"Barcode not found in master list:\n{data}"))
            return False
                
        except Exception as e:
            self.update_results(f"Validation error: {e}\n")
            return False

    def is_any_unmatched_barcode_before(self, line_number):
        """Check if there are any unmatched barcodes before the given line"""
        try:
            for i in range(1, line_number):
                if "matched" not in self.master_text.tag_names(f"{i}.0"):
                    return True
            return False
        except tk.TclError:
            return False

    def add_to_master_list(self, data):
        """Add barcode to master list from Input 1 if not already present."""

        # Check if the barcode is already in the master list
        master_barcodes = self.master_text.get("1.0", tk.END).splitlines()
        if data.strip() in [b.strip() for b in master_barcodes if b.strip()]:
            self.update_results(f"Barcode already exists in master list: {data}\n")
            return  # Do not add duplicates

        # Check if the barcode is uppercase
        if not data.isupper():
            self.update_results(f"Barcode is not uppercase and will not be added: {data}\n")
            return

        # If the length of data is greater than 14, truncate it to 14 characters
        if len(data) > 14:
            data = data[:14]

        self.master.after(0, lambda: self._safe_master_update(data))

    def _safe_master_update(self, data):
        """Safely update master text widget"""
        try:
            # Temporarily enable the text widget
            self.master_text.config(state="normal")

            # Add to master list with unmatched tag only
            self.master_text.insert(tk.END, f"{data}\n", 'unmatched')

            # Disable the text widget again
            self.master_text.config(state="disabled")
        except tk.TclError:
            pass

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

        # Add validation for length_entry
        vcmd = (self.master.register(self.validate_length_input), '%P')
        length_entry.configure(validate='key', validatecommand=vcmd)

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


    def validate_input_text(self, event):
        """Validate input text and update master list"""
        try:
            # Get content from input text
            content = self.input_text.get("1.0", tk.END).strip()
            lines = content.split('\n')
            cleaned_lines = [line.strip() for line in lines if line.strip()]
            
            # Update barcode list
            self.barcode_list = cleaned_lines
            
            # Update master text preserving existing tags
            current_content = self.master_text.get("1.0", tk.END).splitlines()
            current_content = [line.strip() for line in current_content if line.strip()]
            
            self.master_text.delete("1.0", tk.END)
            for line in cleaned_lines:
                # Check if line was previously matched
                if line in current_content:
                    idx = current_content.index(line)
                    if "matched" in self.master_text.tag_names(f"{idx+1}.0"):
                        self.master_text.insert(tk.END, f"{line}\n", 'matched')
                        continue
                self.master_text.insert(tk.END, f"{line}\n", 'unmatched')
            
            self.update_results(f"Loaded {len(cleaned_lines)} barcodes\n")
            
        except Exception as e:
            self.update_results(f"Error validating input: {e}\n")

    def update_input_text(self):
        """Update input text widget with current barcode list"""
        self.input_text.delete("1.0", tk.END)
        self.input_text.insert("1.0", '\n'.join(self.barcode_list))

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
        """Run the tray icon"""
        if self.icon:
            self.icon.run()

    def on_closing(self):
        """Handle window closing event"""
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            self.cleanup()
            if self.icon:
                self.icon.stop()
                self.icon = None
            self.master.quit()
            self.master.destroy()
            os._exit(0)

    def show_alert_popup(self, is_success, message):
        """Show alert popup with description. NG alerts require manual closing"""
        # Create popup window
        popup = tk.Toplevel(self.master)
        popup.title("Scan Result")
        popup.geometry("400x300")
        
        # Set background color based on result
        bg_color = "green2" if is_success else "red"
        fg_color = "white"
        popup.configure(bg=bg_color)
        
        # Make window appear on top
        popup.lift()
        popup.attributes('-topmost', True)
        
        # Create result text (OK/NG)
        result_label = tk.Label(
            popup,
            text="OK" if is_success else "NG",
            font=("Arial", 72, "bold"),
            fg=fg_color,
            bg=bg_color
        )
        result_label.pack(pady=20)
        
        # Create description text
        desc_label = tk.Label(
            popup,
            text=message,
            font=("Arial", 12),
            fg=fg_color,
            bg=bg_color,
            wraplength=350
        )
        desc_label.pack(pady=20)
        
        if is_success:
            # Auto-close after 2 seconds for OK
            popup.after(2000, popup.destroy)
        else:
            # Add close button for NG
            close_btn = tk.Button(
                popup,
                text="Close",
                command=popup.destroy,
                font=("Arial", 12),
                bg="darkred",
                fg="white"
            )
            close_btn.pack(pady=10)
            
            # Center the window on screen
            popup.update_idletasks()
            width = popup.winfo_width()
            height = popup.winfo_height()
            x = (popup.winfo_screenwidth() // 2) - (width // 2)
            y = (popup.winfo_screenheight() // 2) - (height // 2)
            popup.geometry(f'+{x}+{y}')

    def move_barcode_up(self):
        """Move selected barcode up in the Input Barcodes window."""
        if self.running:
            return

        try:
            # Temporarily enable the text widget
            self.input_text.config(state="normal")

            # Get current selection
            sel_start = self.input_text.index("sel.first").split(".")[0]
            current_line = int(sel_start)

            if current_line > 1:
                # Get all lines
                lines = self.input_text.get("1.0", tk.END).splitlines()

                # Swap lines
                lines[current_line - 2], lines[current_line - 1] = lines[current_line - 1], lines[current_line - 2]

                # Update text and selection
                self.input_text.delete("1.0", tk.END)
                self.input_text.insert("1.0", "\n".join(lines) + "\n")

                # Update selection
                new_line = current_line - 1
                self.input_text.tag_add("sel", f"{new_line}.0", f"{new_line}.end")

            # Disable the text widget again
            self.input_text.config(state="disabled")
        except Exception as e:
            self.update_results(f"Error moving barcode: {e}\n")

    def move_barcode_down(self):
        """Move selected barcode down in the Input Barcodes window."""
        if self.running:
            return

        try:
            # Temporarily enable the text widget
            self.input_text.config(state="normal")

            # Get current selection
            sel_start = self.input_text.index("sel.first").split(".")[0]
            current_line = int(sel_start)
            total_lines = int(self.input_text.index("end-1c").split(".")[0])

            if current_line < total_lines:
                # Get all lines
                lines = self.input_text.get("1.0", tk.END).splitlines()

                # Swap lines
                lines[current_line - 1], lines[current_line] = lines[current_line], lines[current_line - 1]

                # Update text and selection
                self.input_text.delete("1.0", tk.END)
                self.input_text.insert("1.0", "\n".join(lines) + "\n")

                # Update selection
                new_line = current_line + 1
                self.input_text.tag_add("sel", f"{new_line}.0", f"{new_line}.end")

            # Disable the text widget again
            self.input_text.config(state="disabled")
        except Exception as e:
            self.update_results(f"Error moving barcode: {e}\n")

    def delete_selected_input_barcode(self):
        """Delete the selected barcode from the Input Barcodes window."""
        try:
            # Temporarily enable the text widget
            self.input_text.config(state="normal")

            # Get the selected text
            sel_start = self.input_text.index("sel.first")
            sel_end = self.input_text.index("sel.last")

            # Delete the selected text
            self.input_text.delete(sel_start, sel_end)

            # Update the barcode list
            lines = self.input_text.get("1.0", tk.END).splitlines()
            self.barcode_list = [line.strip() for line in lines if line.strip()]

            # Disable the text widget again
            self.input_text.config(state="disabled")
        except tk.TclError:
            pass  # No selection made

    def move_master_up(self):
        """Move selected barcode up in the master list"""
        if self.running:
            return
            
        try:
            # Get current selection
            try:
                sel_start = self.master_text.index("sel.first").split(".")[0]
                sel_end = self.master_text.index("sel.last").split(".")[0]
            except tk.TclError:
                return
                
            # Convert to line numbers
            current_line = int(sel_start)
            if current_line > 1:
                # Get all lines and their tags
                lines = []
                tags = []
                for i, line in enumerate(self.master_text.get("1.0", tk.END).splitlines(), 1):
                    lines.append(line)
                    if "matched" in self.master_text.tag_names(f"{i}.0"):
                        tags.append('matched')
                    else:
                        tags.append('unmatched')
                
                # Swap lines and tags
                lines[current_line-2], lines[current_line-1] = lines[current_line-1], lines[current_line-2]
                tags[current_line-2], tags[current_line-1] = tags[current_line-1], tags[current_line-2]
                
                # Update text and preserve tags
                self.master_text.delete("1.0", tk.END)
                for line, tag in zip(lines, tags):
                    self.master_text.insert(tk.END, f"{line}\n", tag)
                
                # Update selection
                new_line = current_line - 1
                self.master_text.tag_add("sel", f"{new_line}.0", f"{new_line}.end")
                
        except Exception as e:
            self.update_results(f"Error moving master barcode: {e}\n")

    def move_master_down(self):
        """Move selected barcode down in the master list"""
        if self.running:
            return
            
        try:
            # Get current selection
            try:
                sel_start = self.master_text.index("sel.first").split(".")[0]
                sel_end = self.master_text.index("sel.last").split(".")[0]
            except tk.TclError:
                return
                
            # Convert to line numbers
            current_line = int(sel_start)
            total_lines = int(self.master_text.index("end-1c").split(".")[0])
            
            if current_line < total_lines:
                # Get all lines and their tags
                lines = []
                tags = []
                for i, line in enumerate(self.master_text.get("1.0", tk.END).splitlines(), 1):
                    lines.append(line)
                    if "matched" in self.master_text.tag_names(f"{i}.0"):
                        tags.append('matched')
                    else:
                        tags.append('unmatched')
                
                # Swap lines and tags
                lines[current_line-1], lines[current_line] = lines[current_line], lines[current_line-1]
                tags[current_line-1], tags[current_line] = tags[current_line], tags[current_line-1]
                
                # Update text and preserve tags
                self.master_text.delete("1.0", tk.END)
                for line, tag in zip(lines, tags):
                    self.master_text.insert(tk.END, f"{line}\n", tag)
                
                # Update selection
                new_line = current_line + 1
                self.master_text.tag_add("sel", f"{new_line}.0", f"{new_line}.end")
                
        except Exception as e:
            self.update_results(f"Error moving master barcode: {e}\n")

    def delete_selected_master_barcode(self):
        """Delete the selected barcode from the Master Barcodes window and update the list."""
        try:
            # Temporarily enable the text widget
            self.master_text.config(state="normal")

            # Get the selected text
            sel_start = self.master_text.index("sel.first")
            sel_end = self.master_text.index("sel.last")

            # Delete the selected text
            self.master_text.delete(sel_start, sel_end)

            # Rebuild the master list to remove empty rows and preserve order
            lines = self.master_text.get("1.0", tk.END).splitlines()
            cleaned_lines = [line.strip() for line in lines if line.strip()]

            # Clear the text widget and reinsert cleaned lines
            self.master_text.delete("1.0", tk.END)
            for line in cleaned_lines:
                self.master_text.insert(tk.END, f"{line}\n", 'unmatched')

            # Disable the text widget again
            self.master_text.config(state="disabled")
        except tk.TclError:
            pass  # No selection made

    def clear_barcodes(self):
        """Clear all barcodes from input"""
        if self.running:
            return
            
        if messagebox.askokcancel("Clear Barcodes", "Are you sure you want to clear all barcodes?"):
            self.input_text.delete("1.0", tk.END)
            self.barcode_list.clear()
            self.update_results("Cleared all barcodes\n")

    def toggle_log_visibility(self):
        """Toggle the visibility of the communication log"""
        if self.log_visible:
            self.log_frame.pack_forget()
            self.toggle_log_btn.config(text="Show Log")
            self.log_visible = False
        else:
            self.log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
            self.toggle_log_btn.config(text="Hide Log")
            self.log_visible = True

    def update_scanned_barcodes(self, barcode):
        """Update the scanned barcodes list"""
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        self.master.after(0, lambda: self._safe_scanned_update(f"[{timestamp}] {barcode}\n"))

    def clear_master(self):
        """Clear all barcodes from master list"""
        if self.running:
            return
            
        if messagebox.askokcancel("Clear Master Barcodes", "Are you sure you want to clear all master barcodes?"):
            self.master_text.delete("1.0", tk.END)
            self.input_text.delete("1.0", tk.END)
            self.barcode_list.clear()
            self.update_results("Cleared all master barcodes\n")

    def add_to_input_barcodes(self, data, is_valid):
        """Add barcode to Input Barcodes window with appropriate color"""
        try:
            # Temporarily enable the text widget
            self.input_text.config(state="normal")

            # Insert barcode at the top of the input barcodes window
            self.input_text.insert("1.0", f"{data}\n")
            
            # Apply color tag based on validation result
            tag = 'ok' if is_valid else 'ng'
            self.input_text.tag_add(tag, "1.0", "1.0 lineend")
            
            # Configure tags for coloring
            self.input_text.tag_configure('ok', background='light green')
            self.input_text.tag_configure('ng', background='red')

            # Disable the text widget again
            self.input_text.config(state="disabled")
        except tk.TclError:
            pass

    def remove_matched_master_barcode(self, line_number):
        """Remove the matched barcode from the Master Barcodes window"""
        try:
            # Temporarily enable the text widget
            self.master_text.config(state="normal")

            # Calculate the start and end indices of the line to remove
            start_index = f"{line_number}.0"
            end_index = f"{line_number + 1}.0"

            # Delete the line
            self.master_text.delete(start_index, end_index)

            # Disable the text widget again
            self.master_text.config(state="disabled")
        except tk.TclError as e:
            self.update_results(f"Error removing matched barcode: {e}\n")

    def clean_data(self, raw_data):
        """Remove STX, ETX, and any non-alphanumeric characters from the raw data."""
        try:
            # Check and remove STX (ASCII 0x02) at the start
            if raw_data.startswith("\x02"):
                raw_data = raw_data[1:]  # Remove the first character (STX)

            # Check and remove ETX (ASCII 0x03) at the end
            if raw_data.endswith("\x03"):
                raw_data = raw_data[:-1]  # Remove the last character (ETX)

            # Remove any non-alphanumeric characters
            cleaned_data = ''.join(c for c in raw_data if c.isalnum())

            # Return the cleaned data
            return cleaned_data.strip()
        except Exception as e:
            self.update_results(f"Error cleaning data: {e}\n")
            return raw_data

def main():
    root = tk.Tk()
    app = Application(root)
    root.mainloop()

if __name__ == '__main__':
    main()
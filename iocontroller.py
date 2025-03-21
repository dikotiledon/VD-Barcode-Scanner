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

class Application:
    def __init__(self, master):
        self.master = master
        self.master.title("Barcode Scanner Interface")
        self.master.geometry('600x500+100+100')
        self.master.protocol('WM_DELETE_WINDOW', self.on_closing)
        
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
        
        # Ports selection frame (contains both input and output in same row)
        ports_frame = tk.Frame(com_ports_frame)
        ports_frame.pack(fill=tk.X, pady=5, padx=5)
        
        # Input COM port selection (left side)
        input_frame = tk.Frame(ports_frame)
        input_frame.pack(side=tk.LEFT, expand=True)
        # tk.Label(input_frame, text="Input (Barcode Scanner):").pack(side=tk.LEFT)
        
        tk.Label(input_frame, text="Input:").pack(side=tk.LEFT)
        self.input_com_var = tk.StringVar()
        self.input_com_var.set(com_ports[0] if com_ports else "No COM Ports Found")
        self.input_com_dropdown = ttk.Combobox(input_frame, textvariable=self.input_com_var, values=com_ports, width=15)
        self.input_com_dropdown.pack(side=tk.LEFT, padx=5)
        
        # Output COM port selection (right side)
        output_frame = tk.Frame(ports_frame)
        output_frame.pack(side=tk.LEFT, expand=True)
        # tk.Label(output_frame, text="Output (I/O Controller):").pack(side=tk.LEFT)
        tk.Label(output_frame, text="Output:").pack(side=tk.LEFT)
        self.output_com_var = tk.StringVar()
        self.output_com_var.set(com_ports[1] if len(com_ports) > 1 else com_ports[0] if com_ports else "No COM Ports Found")
        self.output_com_dropdown = ttk.Combobox(output_frame, textvariable=self.output_com_var, values=com_ports, width=15)
        self.output_com_dropdown.pack(side=tk.LEFT, padx=5)
        
        # Configuration frame (contains both configs and refresh button in same row)
        config_frame = tk.Frame(com_ports_frame)
        config_frame.pack(fill=tk.X, pady=5, padx=5)
        
        # Input serial configuration (left)
        input_config_frame = tk.LabelFrame(config_frame, text="Input Port Configuration")
        input_config_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        
        tk.Label(input_config_frame, text="Baud Rate:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=2)
        self.input_baud_var = tk.StringVar(value="9600")
        input_baud_combo = ttk.Combobox(input_config_frame, textvariable=self.input_baud_var, 
                                       values=["9600", "19200", "38400", "57600", "115200"], width=5)
        input_baud_combo.grid(row=0, column=1, sticky=tk.W, padx=4, pady=2)
        
        # Output serial configuration (middle)
        output_config_frame = tk.LabelFrame(config_frame, text="Output Port Configuration")
        output_config_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        tk.Label(output_config_frame, text="Baud Rate:").grid(row=0, column=0, sticky=tk.W, padx=4, pady=2)
        self.output_baud_var = tk.StringVar(value="9600")
        output_baud_combo = ttk.Combobox(output_config_frame, textvariable=self.output_baud_var, 
                                        values=["9600", "19200", "38400", "57600", "115200"], width=5)
        output_baud_combo.grid(row=0, column=1, sticky=tk.W, padx=4, pady=2)

        # Add switch number selection
        tk.Label(output_config_frame, text="Switch:").grid(row=1, column=0, sticky=tk.W, padx=4, pady=2)
        switch_combobox = ttk.Combobox(output_config_frame, textvariable=self.switch_number, 
                                      values=list(range(1, 3)), width=3)
        switch_combobox.grid(row=1, column=1, sticky=tk.W, padx=4, pady=2)
        switch_combobox.current(0)  # set default value to 1

        # Add validation for switch input
        vcmd2 = (self.master.register(self.validate_switch_input), '%P')
        switch_combobox.configure(validate='key', validatecommand=vcmd2)
        
        # Refresh button frame (right)
        refresh_frame = tk.Frame(config_frame)
        refresh_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        refresh_btn = tk.Button(
            refresh_frame, 
            text="Refresh COM Ports", 
            command=self.refresh_com_ports
        )
        refresh_btn.pack(side=tk.LEFT, padx=5, pady=10)  # Added pady to center vertically
        
        # Create horizontal layout frame
        horizontal_frame = tk.Frame(main_frame)
        horizontal_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Input column frame (left side)
        input_frame = tk.LabelFrame(horizontal_frame, text="Input Barcodes")
        input_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # Add buttons frame above text box
        input_buttons_frame = tk.Frame(input_frame)
        input_buttons_frame.pack(fill=tk.X, padx=5, pady=2)
        
        # Add move up/down/clear buttons
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
        
        self.clear_btn = tk.Button(
            input_buttons_frame,
            text="Clear All",
            command=self.clear_barcodes,
            bg="red",
            fg="white"
        )
        self.clear_btn.pack(side=tk.RIGHT, padx=2)
        
        # Input text box with multiline support
        self.input_text = tk.Text(input_frame, width=30, height=10, wrap=tk.WORD)
        self.input_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Add scrollbar for input
        input_scrollbar = tk.Scrollbar(input_frame)
        input_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.input_text.config(yscrollcommand=input_scrollbar.set)
        input_scrollbar.config(command=self.input_text.yview)
        
        # Bind validation to input text
        self.input_text.bind('<KeyRelease>', self.validate_input_text)
        
        # Results frame (right side)
        results_frame = tk.LabelFrame(horizontal_frame, text="Communication Log")
        results_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.results_text = tk.Text(results_frame, height=20)
        self.results_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Add scrollbar for results
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
                if self.input_serial and self.input_serial.is_open:
                    if self.input_serial.in_waiting > 0:
                        data = self.input_serial.readline().decode('utf-8', errors='replace').strip()
                        
                        if data:
                            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
                            log_message = f"[{timestamp}] Received: {data}\n"                        
                            self.update_results(log_message)
                            
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
                                            # Show success popup
                                            self.master.after(0, lambda: self.show_alert_popup(True, f"Barcode: {data}\nCommand: {command}"))
                                    else:
                                        self.output_serial.write(('@OFF01$' + "\r\n").encode('utf-8'))
                                        self.output_serial.write(('@OFF02$' + "\r\n").encode('utf-8'))
                                        self.update_results(f"Sent to I/O controller: @OFF01$\n")
                                        # Show failure popup
                                        self.master.after(0, lambda: self.show_alert_popup(False, 
                                            f"Expected: {self.barcode_list[0] if self.barcode_list else 'No barcode in list'}\nGot: {data}"))
                                        # Use after method to safely call UI methods from thread
                                        self.master.after(0, self.toggle_listen)
                                        return
                                except Exception as e:
                                    self.update_results(f"Error sending to output port: {e}\n")
                
                time.sleep(0.05)
                
            except Exception as e:
                self.update_results(f"Error in serial communication: {e}\n")
                time.sleep(1)
                
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

    def validate_barcode(self, data):
        """Validate barcode against input list"""
        try:
            if not self.barcode_list:
                self.update_results("No barcodes in input list\n")
                return False
                
            data = data.strip()
            if data == self.barcode_list[0]:
                # Remove the matched barcode from the list
                self.barcode_list.pop(0)
                
                # Update input text
                self.master.after(0, self.update_input_text)
                return True
            else:
                self.update_results(f"Barcode mismatch. Expected: {self.barcode_list[0]}, Got: {data}\n")
                return False
                
        except Exception as e:
            self.update_results(f"Validation error: {e}\n")
            return False

    def validate_input_text(self, event):
        """Validate input text to ensure proper barcode format per line"""
        try:
            # Get all text content
            content = self.input_text.get("1.0", tk.END).strip()
            
            # Split into lines and clean
            lines = content.split('\n')
            cleaned_lines = []
            
            for line in lines:
                # Remove extra whitespace but keep valid lines
                cleaned = line.strip()
                if cleaned:
                    cleaned_lines.append(cleaned)
            
            # Update barcode list
            self.barcode_list = cleaned_lines
            
            # Only update text if content has changed
            new_content = '\n'.join(cleaned_lines)
            if content != new_content:
                # Save cursor position
                cursor_pos = self.input_text.index(tk.INSERT)
                
                # Update text content
                self.input_text.delete("1.0", tk.END)
                self.input_text.insert("1.0", new_content)
                
                # Restore cursor position if possible
                try:
                    self.input_text.mark_set(tk.INSERT, cursor_pos)
                except tk.TclError:
                    pass
                    
            # Update results display
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
        """Run system tray icon in separate thread"""
        self.icon.run()

    def on_closing(self):
        """Handle window closing event"""
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            self.cleanup()
            self.master.quit()
            self.master.destroy()
            os._exit(0)  # Force exit if needed

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
        """Move selected barcode up in the list"""
        if self.running:
            return
            
        try:
            # Get current selection
            try:
                sel_start = self.input_text.index("sel.first").split(".")[0]
                sel_end = self.input_text.index("sel.last").split(".")[0]
            except tk.TclError:
                return
                
            # Convert to line numbers
            current_line = int(sel_start)
            if current_line > 1:
                # Get all lines
                lines = self.input_text.get("1.0", tk.END).splitlines()
                
                # Swap lines
                lines[current_line-2], lines[current_line-1] = lines[current_line-1], lines[current_line-2]
                
                # Update text and selection
                self.input_text.delete("1.0", tk.END)
                self.input_text.insert("1.0", "\n".join(lines))
                
                # Update selection
                new_line = current_line - 1
                self.input_text.tag_add("sel", f"{new_line}.0", f"{new_line}.end")
                
                # Update barcode list
                self.barcode_list = [line.strip() for line in lines if line.strip()]
                
        except Exception as e:
            self.update_results(f"Error moving barcode: {e}\n")

    def move_barcode_down(self):
        """Move selected barcode down in the list"""
        if self.running:
            return
            
        try:
            # Get current selection
            try:
                sel_start = self.input_text.index("sel.first").split(".")[0]
                sel_end = self.input_text.index("sel.last").split(".")[0]
            except tk.TclError:
                return
                
            # Convert to line numbers
            current_line = int(sel_start)
            total_lines = int(self.input_text.index("end-1c").split(".")[0])
            
            if current_line < total_lines:
                # Get all lines
                lines = self.input_text.get("1.0", tk.END).splitlines()
                
                # Swap lines
                lines[current_line-1], lines[current_line] = lines[current_line], lines[current_line-1]
                
                # Update text and selection
                self.input_text.delete("1.0", tk.END)
                self.input_text.insert("1.0", "\n".join(lines))
                
                # Update selection
                new_line = current_line + 1
                self.input_text.tag_add("sel", f"{new_line}.0", f"{new_line}.end")
                
                # Update barcode list
                self.barcode_list = [line.strip() for line in lines if line.strip()]
                
        except Exception as e:
            self.update_results(f"Error moving barcode: {e}\n")

    def clear_barcodes(self):
        """Clear all barcodes from input"""
        if self.running:
            return
            
        if messagebox.askokcancel("Clear Barcodes", "Are you sure you want to clear all barcodes?"):
            self.input_text.delete("1.0", tk.END)
            self.barcode_list.clear()
            self.update_results("Cleared all barcodes\n")

def main():
    root = tk.Tk()
    root.iconbitmap("app.ico")
    app = Application(root)
    root.mainloop()

if __name__ == '__main__':
    main()
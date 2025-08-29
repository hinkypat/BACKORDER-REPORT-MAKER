"""
GUI Module for Back Order Report Generator
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import logging
import os
from datetime import datetime
from .data_processor import DataProcessor
from .excel_generator import ExcelGenerator

class BackOrderReportGUI:
    def __init__(self, root, config):
        self.root = root
        self.config = config
        self.logger = logging.getLogger(__name__)
        
        self.input_file_path = tk.StringVar()
        self.output_directory = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar(value="Ready")
        
        self.data_processor = DataProcessor(config)
        self.excel_generator = ExcelGenerator(config)
        
        self.setup_gui()
        
    def setup_gui(self):
        """Initialize and setup the GUI components"""
        self.root.title("Back Order Report Generator v1.0")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="wens")
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Input file selection
        ttk.Label(main_frame, text="Input Database Export File:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.input_file_path, width=50).grid(row=0, column=1, sticky="we", pady=5, padx=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_input_file).grid(row=0, column=2, pady=5)
        
        # Output directory selection
        ttk.Label(main_frame, text="Output Directory:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_directory, width=50).grid(row=1, column=1, sticky="we", pady=5, padx=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_output_directory).grid(row=1, column=2, pady=5)
        
        # Processing options frame
        options_frame = ttk.LabelFrame(main_frame, text="Processing Options", padding="10")
        options_frame.grid(row=2, column=0, columnspan=3, sticky="we", pady=10)
        options_frame.columnconfigure(1, weight=1)
        
        # Report type selection
        ttk.Label(options_frame, text="Report Type:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.report_type = tk.StringVar(value="standard")
        report_combo = ttk.Combobox(options_frame, textvariable=self.report_type, 
                                   values=["standard", "detailed", "summary"], state="readonly")
        report_combo.grid(row=0, column=1, sticky="we", pady=2, padx=5)
        
        # Include charts option
        self.include_charts = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Include Charts and Graphs", 
                       variable=self.include_charts).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=2)
        
        # Data validation option
        self.validate_data = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Enable Data Validation", 
                       variable=self.validate_data).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=2)
        
        # Progress section
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="10")
        progress_frame.grid(row=3, column=0, columnspan=3, sticky="we", pady=10)
        progress_frame.columnconfigure(0, weight=1)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=0, column=0, sticky="we", pady=2)
        
        # Status label
        ttk.Label(progress_frame, textvariable=self.status_var).grid(row=1, column=0, sticky=tk.W, pady=2)
        
        # Log display
        log_frame = ttk.LabelFrame(main_frame, text="Processing Log", padding="10")
        log_frame.grid(row=4, column=0, columnspan=3, sticky="wens", pady=10)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # Log text widget with scrollbar
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky="wens")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Control buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=10)
        
        self.process_button = ttk.Button(button_frame, text="Generate Report", 
                                       command=self.start_processing, style="Accent.TButton")
        self.process_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Exit", command=self.root.quit).pack(side=tk.LEFT, padx=5)
        
        # Set default output directory
        self.output_directory.set(os.path.expanduser("~/Desktop"))
        
    def browse_input_file(self):
        """Browse for input database export file"""
        file_types = [
            ("All Supported", "*.csv;*.xlsx;*.xls;*.txt"),
            ("CSV files", "*.csv"),
            ("Excel files", "*.xlsx;*.xls"),
            ("Text files", "*.txt"),
            ("All files", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="Select Database Export File",
            filetypes=file_types
        )
        
        if filename:
            self.input_file_path.set(filename)
            
    def browse_output_directory(self):
        """Browse for output directory"""
        directory = filedialog.askdirectory(
            title="Select Output Directory"
        )
        
        if directory:
            self.output_directory.set(directory)
            
    def log_message(self, message):
        """Add message to log display"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        
        self.log_text.insert(tk.END, formatted_message)
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def clear_log(self):
        """Clear the log display"""
        self.log_text.delete(1.0, tk.END)
        
    def update_progress(self, value, status):
        """Update progress bar and status"""
        self.progress_var.set(value)
        self.status_var.set(status)
        self.root.update_idletasks()
        
    def start_processing(self):
        """Start the report generation process in a separate thread"""
        if not self.input_file_path.get():
            messagebox.showerror("Error", "Please select an input file")
            return
            
        if not self.output_directory.get():
            messagebox.showerror("Error", "Please select an output directory")
            return
            
        # Disable the process button during processing
        self.process_button.config(state="disabled")
        
        # Start processing in a separate thread
        processing_thread = threading.Thread(target=self.process_report)
        processing_thread.daemon = True
        processing_thread.start()
        
    def process_report(self):
        """Process the report generation"""
        try:
            self.logger.info("Starting report generation process")
            self.log_message("Starting report generation...")
            
            # Update progress
            self.update_progress(10, "Loading input file...")
            
            # Load and validate data
            data = self.data_processor.load_data(
                self.input_file_path.get(),
                validate=self.validate_data.get()
            )
            
            self.update_progress(30, "Processing data...")
            self.log_message(f"Loaded {len(data)} records from input file")
            
            # Process the data
            processed_data = self.data_processor.process_data(data)
            
            self.update_progress(60, "Generating Excel report...")
            self.log_message("Data processing completed successfully")
            
            # Generate output filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"backorder_report_{timestamp}.xlsx"
            output_path = os.path.join(self.output_directory.get(), output_filename)
            
            # Generate Excel report
            self.excel_generator.generate_report(
                processed_data,
                output_path,
                report_type=self.report_type.get(),
                include_charts=self.include_charts.get()
            )
            
            self.update_progress(100, "Report generation completed!")
            self.log_message(f"Report saved to: {output_path}")
            
            # Show success message
            messagebox.showinfo("Success", f"Report generated successfully!\n\nSaved to: {output_path}")
            
        except Exception as e:
            error_msg = f"Error during processing: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            self.log_message(f"ERROR: {error_msg}")
            self.update_progress(0, "Error occurred")
            messagebox.showerror("Processing Error", error_msg)
            
        finally:
            # Re-enable the process button
            self.process_button.config(state="normal")

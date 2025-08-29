#!/usr/bin/env python3
"""
Daily Backorder Report Generator - GUI Application
Simple interface for generating daily backorder reports
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
from datetime import datetime
import sys

# Import your existing BackorderReportGenerator
from backorder_generator import BackorderReportGenerator

class BackorderReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Daily Backorder Report Generator")
        self.root.geometry("700x500")
        self.root.resizable(True, True)
        
        # Variables
        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar(value="Ready to generate report")
        
        self.setup_gui()
        
    def setup_gui(self):
        """Setup the GUI interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky="nsew")
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Daily Backorder Report Generator", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="Select Raw Data File", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 15))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(file_frame, text="Raw Data File:").grid(row=0, column=0, sticky="w", pady=5)
        
        file_entry = ttk.Entry(file_frame, textvariable=self.input_file_path, width=60)
        file_entry.grid(row=0, column=1, sticky="ew", pady=5, padx=5)
        
        browse_btn = ttk.Button(file_frame, text="Browse", command=self.browse_file)
        browse_btn.grid(row=0, column=2, pady=5)
        
        # Output file selection
        ttk.Label(file_frame, text="Save Report As:").grid(row=1, column=0, sticky="w", pady=5)
        
        output_entry = ttk.Entry(file_frame, textvariable=self.output_file_path, width=60)
        output_entry.grid(row=1, column=1, sticky="ew", pady=5, padx=5)
        
        save_btn = ttk.Button(file_frame, text="Browse", command=self.browse_save_location)
        save_btn.grid(row=1, column=2, pady=5)
        
        # Info section
        info_frame = ttk.LabelFrame(main_frame, text="Instructions", padding="10")
        info_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 15))
        
        info_text = """1. Click 'Browse' next to 'Raw Data File' to select your daily data file (.xls or .xlsx)
2. Click 'Browse' next to 'Save Report As' to choose where to save the report
3. Click 'Generate Report' to create the backorder report
4. Military orders (Manuel Ortega + DLA/DFAS/NAVSUP customers) and Commercial orders will be separated"""
        
        info_label = ttk.Label(info_frame, text=info_text, justify="left", wraplength=650)
        info_label.grid(row=0, column=0, sticky="w")
        
        # Progress section
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="10")
        progress_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(0, 15))
        progress_frame.columnconfigure(0, weight=1)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=0, column=0, sticky="ew", pady=2)
        
        # Status label
        status_label = ttk.Label(progress_frame, textvariable=self.status_var)
        status_label.grid(row=1, column=0, sticky="w", pady=2)
        
        # Output section
        output_frame = ttk.LabelFrame(main_frame, text="Output", padding="10")
        output_frame.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(0, 15))
        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(0, weight=1)
        
        # Output log
        self.output_text = tk.Text(output_frame, height=8, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(output_frame, orient=tk.VERTICAL, command=self.output_text.yview)
        self.output_text.configure(yscrollcommand=scrollbar.set)
        
        self.output_text.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Button section
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=20)
        
        self.generate_btn = ttk.Button(button_frame, text="Generate Report", 
                                     command=self.start_processing, style="Accent.TButton")
        self.generate_btn.pack(side=tk.LEFT, padx=10)
        
        clear_btn = ttk.Button(button_frame, text="Clear Output", command=self.clear_output)
        clear_btn.pack(side=tk.LEFT, padx=5)
        
        exit_btn = ttk.Button(button_frame, text="Exit", command=self.root.quit)
        exit_btn.pack(side=tk.LEFT, padx=5)
        
        # Set default input file if it exists
        default_file = "back orders by salesperson report.xls"
        if os.path.exists(default_file):
            self.input_file_path.set(default_file)
            
        # Set default output file name
        default_output = f"BACKORDER REPORT {datetime.now().strftime('%m%d%y')}.xlsx"
        self.output_file_path.set(default_output)
        
        main_frame.rowconfigure(4, weight=1)
        
    def browse_save_location(self):
        """Browse for output file save location"""
        filename = filedialog.asksaveasfilename(
            title="Save Report As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if filename:
            self.output_file_path.set(filename)
            
    def browse_file(self):
        """Browse for input file"""
        file_types = [
            ("Excel files", "*.xlsx;*.xls"),
            ("All files", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="Select Raw Data File",
            filetypes=file_types
        )
        
        if filename:
            self.input_file_path.set(filename)
            
    def log_message(self, message):
        """Add message to output log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        
        self.output_text.insert(tk.END, formatted_message)
        self.output_text.see(tk.END)
        self.root.update_idletasks()
        
    def clear_output(self):
        """Clear the output log"""
        self.output_text.delete(1.0, tk.END)
        
    def update_progress(self, value, status):
        """Update progress bar and status"""
        self.progress_var.set(value)
        self.status_var.set(status)
        self.root.update_idletasks()
        
    def start_processing(self):
        """Start the report generation in a separate thread"""
        if not self.input_file_path.get():
            messagebox.showerror("Error", "Please select an input file")
            return
            
        if not self.output_file_path.get():
            messagebox.showerror("Error", "Please select where to save the output file")
            return
            
        if not os.path.exists(self.input_file_path.get()):
            messagebox.showerror("Error", "Selected file does not exist")
            return
            
        # Disable button during processing
        self.generate_btn.config(state="disabled")
        
        # Start processing thread
        processing_thread = threading.Thread(target=self.process_report)
        processing_thread.daemon = True
        processing_thread.start()
        
    def process_report(self):
        """Process the report generation"""
        try:
            self.log_message("Starting report generation...")
            self.update_progress(10, "Initializing...")
            
            # Create generator instance
            generator = BackorderReportGenerator(self.input_file_path.get())
            
            self.update_progress(20, "Validating input file...")
            self.log_message(f"Processing file: {self.input_file_path.get()}")
            
            self.update_progress(50, "Generating report...")
            
            # Generate the report using your existing code
            output_file = generator.generate_report(self.output_file_path.get())
            
            self.update_progress(100, "Report generated successfully!")
            self.log_message(f"✅ Report saved: {output_file}")
            self.log_message(f"📁 Location: {os.path.abspath(output_file)}")
            
            # Show success message
            messagebox.showinfo("Success", 
                              f"Report generated successfully!\n\n"
                              f"File: {output_file}\n"
                              f"Location: {os.path.abspath(output_file)}")
            
        except Exception as e:
            error_msg = f"Error generating report: {str(e)}"
            self.log_message(f"❌ {error_msg}")
            self.update_progress(0, "Error occurred")
            messagebox.showerror("Error", error_msg)
            
        finally:
            # Re-enable button
            self.generate_btn.config(state="normal")

def main():
    """Main function"""
    root = tk.Tk()
    app = BackorderReportApp(root)
    
    # Handle window close
    def on_closing():
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()
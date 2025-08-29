#!/usr/bin/env python3
"""
Back Order Report Generator - Main Application Entry Point
"""

import sys
import tkinter as tk
from tkinter import messagebox
import logging
from src.gui import BackOrderReportGUI
from src.logger import setup_logging
from src.config import Config

def main():
    """Main application entry point"""
    try:
        # Initialize logging
        setup_logging()
        logger = logging.getLogger(__name__)
        logger.info("Starting Back Order Report Generator")
        
        # Load configuration
        config = Config()
        
        # Create and run the GUI application
        root = tk.Tk()
        app = BackOrderReportGUI(root, config)
        
        # Handle application closing
        def on_closing():
            logger.info("Application closing")
            root.destroy()
        
        root.protocol("WM_DELETE_WINDOW", on_closing)
        
        # Start the GUI main loop
        root.mainloop()
        
    except Exception as e:
        error_msg = f"Failed to start application: {str(e)}"
        logging.error(error_msg, exc_info=True)
        messagebox.showerror("Application Error", error_msg)
        sys.exit(1)

if __name__ == "__main__":
    main()

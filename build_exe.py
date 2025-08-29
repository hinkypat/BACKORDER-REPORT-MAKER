#!/usr/bin/env python3
"""
Build script to create Windows executable from the Daily Backorder App
"""

import os
import subprocess
import sys

def build_executable():
    """Build the Windows executable using PyInstaller"""
    
    print("🔨 Building Windows Executable...")
    print("=" * 50)
    
    # PyInstaller command to create executable
    cmd = [
        "pyinstaller",
        "--onefile",  # Create single executable file
        "--windowed",  # Hide console window (GUI app)
        "--name=Daily_Backorder_Report_Generator",  # Executable name
        "--icon=NONE",  # No icon for now
        "--add-data=backorder_generator.py;.",  # Include the core module
        "daily_backorder_app.py"  # Main GUI file
    ]
    
    try:
        # Run PyInstaller
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        
        print("✅ Build successful!")
        print("\nExecutable created at:")
        print("  📁 dist/Daily_Backorder_Report_Generator.exe")
        
        # Check if file was created
        exe_path = "dist/Daily_Backorder_Report_Generator.exe"
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"  📊 File size: {size_mb:.1f} MB")
        
        print("\n🎉 Ready to distribute!")
        print("Copy the .exe file to any Windows computer and run it directly.")
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Build failed: {e}")
        print(f"Output: {e.stdout}")
        print(f"Error: {e.stderr}")
        return False
    
    return True

if __name__ == "__main__":
    if build_executable():
        print("\n✅ SUCCESS: Your Windows executable is ready!")
    else:
        print("\n❌ FAILED: Could not create executable")
        sys.exit(1)
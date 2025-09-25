# Back Order Report Generator - REV6

A Python desktop application that automates the transformation of raw database exports into formatted Excel back order reports for management.

## NEW in REV6
- **Sort By Section**: Choose your preferred sorting method with checkboxes
  - ORDER # (default)
  - PCX DOCK
  - SALESMAN NAME  
  - DUE DATE
- **Smart Sorting Logic**: Uses selected column if exactly one is checked, defaults to ORDER # otherwise
- **Enhanced GUI**: Updated interface with improved instructions and feedback

## Features

- **Multi-format Input Support**: Handles CSV, Excel (XLSX/XLS), and text files
- **Intelligent Data Processing**: Automatically detects and validates data structure
- **Professional Excel Reports**: Generates formatted workbooks with multiple sheets
- **Data Visualization**: Includes charts and graphs for management insights
- **User-friendly GUI**: Intuitive interface for non-technical staff
- **Robust Error Handling**: Comprehensive validation and error reporting
- **Progress Tracking**: Real-time progress indicators and status updates
- **Flexible Configuration**: Customizable settings via configuration file

## Installation

### Requirements

The application requires Python 3.7 or higher and the following packages:
- pandas
- openpyxl
- tkinter (usually included with Python)
- configparser (included with Python)

### Setup

1. **Clone or download the application files**
2. **Install required packages**:
   ```bash
   pip install pandas openpyxl
   ```
3. **Run the REV6 application**:
   ```bash
   python daily_backorder_app.py
   ```
   
   Or run the alternative interface:
   ```bash
   python main.py
   ```

## Usage

### Basic Operation

1. **Launch the application** by running `python daily_backorder_app.py` (REV6 interface)
2. **Select input file**: Click "Browse" next to "Raw Data File" and select your data file
3. **Choose output location**: Click "Browse" next to "Save Report As" to select where the report will be saved
4. **Select sorting preference**: Choose your preferred sorting method in the "Sort By" section:
   - Check exactly one option for custom sorting
   - Leave default (ORDER #) or select multiple to use ORDER # sorting
5. **Generate report**: Click "Generate Report" to start processing

### Input File Requirements

The application can process files with the following formats:
- **CSV files** (`.csv`)
- **Excel files** (`.xlsx`, `.xls`)
- **Text files** (`.txt`) with various delimiters

#### Required Columns

Your input file should contain at least these columns (names are flexible):
- **Item Code/SKU**: Product identifier
- **Quantity**: Back order quantity
- **Order Date**: When the back order was created

#### Optional Columns

Additional columns that enhance reporting:
- **Customer**: Customer name or identifier
- **Supplier**: Supplier/vendor information
- **Expected Date**: Expected delivery date
- **Unit Price**: Price per unit
- **Category**: Product category

### Output

The application generates Excel (.xlsx) reports with multiple sheets:

#### Standard Report Includes:
- **Summary**: Overview statistics and top items
- **By Item**: Analysis grouped by product
- **By Customer**: Analysis grouped by customer (if available)
- **Aging Analysis**: Back orders grouped by age
- **Charts**: Visual representations of key metrics

#### Detailed Report Includes:
All standard sheets plus:
- **By Supplier**: Analysis grouped by supplier
- **By Date**: Monthly trend analysis
- **By Category**: Analysis grouped by product category
- **Raw Data**: Complete processed dataset

#### Summary Report Includes:
- **Summary**: Key statistics and metrics
- **Charts**: Essential visualizations only

## Configuration

The application uses a `config.ini` file for customization:

### Processing Settings
```ini
[PROCESSING]
validate_data = true          # Enable data validation
remove_duplicates = false     # Remove duplicate records
default_report_type = standard # Default report type
include_charts = true         # Include charts by default

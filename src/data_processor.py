"""
Data Processing Module for Back Order Report Generator
"""

import pandas as pd
import logging
import os
from datetime import datetime
import numpy as np

class DataProcessor:
    def __init__(self, config):
        self.config = config
        self.logger = logging.getLogger(__name__)
        
    def load_data(self, file_path, validate=True):
        """
        Load data from various file formats
        
        Args:
            file_path (str): Path to the input file
            validate (bool): Whether to perform data validation
            
        Returns:
            pandas.DataFrame: Loaded data
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Input file not found: {file_path}")
            
        file_extension = os.path.splitext(file_path)[1].lower()
        
        try:
            if file_extension == '.csv':
                data = pd.read_csv(file_path, encoding='utf-8')
            elif file_extension in ['.xlsx', '.xls']:
                data = pd.read_excel(file_path)
            elif file_extension == '.txt':
                # Try to determine delimiter
                with open(file_path, 'r', encoding='utf-8') as f:
                    first_line = f.readline()
                    if '\t' in first_line:
                        delimiter = '\t'
                    elif '|' in first_line:
                        delimiter = '|'
                    else:
                        delimiter = ','
                data = pd.read_csv(file_path, delimiter=delimiter, encoding='utf-8')
            else:
                raise ValueError(f"Unsupported file format: {file_extension}")
                
            self.logger.info(f"Successfully loaded {len(data)} rows from {file_path}")
            
            if validate:
                data = self._validate_data(data)
                
            return data
            
        except Exception as e:
            error_msg = f"Failed to load data from {file_path}: {str(e)}"
            self.logger.error(error_msg)
            raise Exception(error_msg)
            
    def _validate_data(self, data):
        """
        Validate and clean the input data
        
        Args:
            data (pandas.DataFrame): Raw input data
            
        Returns:
            pandas.DataFrame: Validated and cleaned data
        """
        self.logger.info("Starting data validation")
        original_rows = len(data)
        
        # Check for empty dataset
        if data.empty:
            raise ValueError("Input file is empty or contains no valid data")
            
        # Standardize column names (remove spaces, convert to lowercase)
        data.columns = data.columns.str.strip().str.lower().str.replace(' ', '_')
        
        # Required columns for back order reports
        required_columns = ['item_code', 'quantity', 'order_date']
        optional_columns = ['customer', 'supplier', 'expected_date', 'unit_price', 'category']
        
        # Check for required columns (flexible matching)
        column_mapping = {}
        for req_col in required_columns:
            found = False
            for col in data.columns:
                if any(keyword in col for keyword in self._get_column_keywords(req_col)):
                    column_mapping[req_col] = col
                    found = True
                    break
            if not found:
                raise ValueError(f"Required column '{req_col}' not found in input data. "
                               f"Available columns: {list(data.columns)}")
        
        # Rename columns to standard names
        data = data.rename(columns=column_mapping)
        
        # Remove completely empty rows
        data = data.dropna(how='all')
        
        # Validate data types and handle missing values
        data = self._clean_data_types(data)
        
        # Remove invalid records
        initial_count = len(data)
        data = data[data['quantity'] > 0]  # Remove zero or negative quantities
        data = data.dropna(subset=['item_code'])  # Remove rows without item codes
        
        if len(data) == 0:
            raise ValueError("No valid records found after data validation")
            
        validation_summary = {
            'original_rows': original_rows,
            'valid_rows': len(data),
            'removed_rows': original_rows - len(data)
        }
        
        self.logger.info(f"Data validation completed: {validation_summary}")
        return data
        
    def _get_column_keywords(self, column_type):
        """Get potential keywords for column identification"""
        keywords = {
            'item_code': ['item', 'product', 'sku', 'part', 'code', 'id'],
            'quantity': ['quantity', 'qty', 'amount', 'count'],
            'order_date': ['date', 'order', 'created', 'timestamp'],
            'customer': ['customer', 'client', 'buyer'],
            'supplier': ['supplier', 'vendor', 'manufacturer'],
            'expected_date': ['expected', 'due', 'delivery', 'eta'],
            'unit_price': ['price', 'cost', 'value', 'amount'],
            'category': ['category', 'type', 'class', 'group']
        }
        return keywords.get(column_type, [])
        
    def _clean_data_types(self, data):
        """Clean and convert data types"""
        # Convert quantity to numeric
        data['quantity'] = pd.to_numeric(data['quantity'], errors='coerce')
        
        # Convert dates
        if 'order_date' in data.columns:
            data['order_date'] = pd.to_datetime(data['order_date'], errors='coerce')
            
        if 'expected_date' in data.columns:
            data['expected_date'] = pd.to_datetime(data['expected_date'], errors='coerce')
            
        # Convert price to numeric if available
        if 'unit_price' in data.columns:
            data['unit_price'] = pd.to_numeric(data['unit_price'], errors='coerce')
            
        # Clean text fields
        text_columns = ['item_code', 'customer', 'supplier', 'category']
        for col in text_columns:
            if col in data.columns:
                data[col] = data[col].astype(str).str.strip()
                
        return data
        
    def process_data(self, data):
        """
        Process the validated data for report generation
        
        Args:
            data (pandas.DataFrame): Validated input data
            
        Returns:
            dict: Processed data ready for Excel generation
        """
        self.logger.info("Starting data processing")
        
        processed_data = {}
        
        # Summary statistics
        processed_data['summary'] = self._generate_summary(data)
        
        # Back order analysis by item
        processed_data['by_item'] = self._analyze_by_item(data)
        
        # Back order analysis by customer (if available)
        if 'customer' in data.columns:
            processed_data['by_customer'] = self._analyze_by_customer(data)
            
        # Back order analysis by supplier (if available)
        if 'supplier' in data.columns:
            processed_data['by_supplier'] = self._analyze_by_supplier(data)
            
        # Time-based analysis
        if 'order_date' in data.columns:
            processed_data['by_date'] = self._analyze_by_date(data)
            
        # Category analysis (if available)
        if 'category' in data.columns:
            processed_data['by_category'] = self._analyze_by_category(data)
            
        # Aging analysis
        processed_data['aging'] = self._analyze_aging(data)
        
        # Raw data for detailed view
        processed_data['raw_data'] = data
        
        self.logger.info("Data processing completed")
        return processed_data
        
    def _generate_summary(self, data):
        """Generate summary statistics"""
        summary = {
            'total_items': len(data),
            'unique_items': data['item_code'].nunique(),
            'total_quantity': data['quantity'].sum(),
            'avg_quantity': data['quantity'].mean(),
        }
        
        if 'unit_price' in data.columns:
            data['total_value'] = data['quantity'] * data['unit_price']
            summary['total_value'] = data['total_value'].sum()
            summary['avg_value'] = data['total_value'].mean()
            
        if 'customer' in data.columns:
            summary['unique_customers'] = data['customer'].nunique()
            
        return summary
        
    def _analyze_by_item(self, data):
        """Analyze back orders by item"""
        by_item = data.groupby('item_code').agg({
            'quantity': ['sum', 'count', 'mean'],
            'order_date': ['min', 'max'] if 'order_date' in data.columns else None
        }).round(2)
        
        # Flatten column names
        by_item.columns = ['_'.join(col).strip() if col[1] else col[0] for col in by_item.columns]
        by_item = by_item.reset_index()
        by_item = by_item.sort_values('quantity_sum', ascending=False)
        
        return by_item
        
    def _analyze_by_customer(self, data):
        """Analyze back orders by customer"""
        by_customer = data.groupby('customer').agg({
            'quantity': ['sum', 'count'],
            'item_code': 'nunique'
        }).round(2)
        
        by_customer.columns = ['_'.join(col).strip() for col in by_customer.columns]
        by_customer = by_customer.reset_index()
        by_customer = by_customer.sort_values('quantity_sum', ascending=False)
        
        return by_customer
        
    def _analyze_by_supplier(self, data):
        """Analyze back orders by supplier"""
        by_supplier = data.groupby('supplier').agg({
            'quantity': ['sum', 'count'],
            'item_code': 'nunique'
        }).round(2)
        
        by_supplier.columns = ['_'.join(col).strip() for col in by_supplier.columns]
        by_supplier = by_supplier.reset_index()
        by_supplier = by_supplier.sort_values('quantity_sum', ascending=False)
        
        return by_supplier
        
    def _analyze_by_date(self, data):
        """Analyze back orders by date"""
        if 'order_date' not in data.columns:
            return pd.DataFrame()
            
        # Remove rows with invalid dates
        date_data = data.dropna(subset=['order_date'])
        
        # Group by month
        date_data['year_month'] = date_data['order_date'].dt.to_period('M')
        by_date = date_data.groupby('year_month').agg({
            'quantity': ['sum', 'count'],
            'item_code': 'nunique'
        }).round(2)
        
        by_date.columns = ['_'.join(col).strip() for col in by_date.columns]
        by_date = by_date.reset_index()
        by_date['year_month'] = by_date['year_month'].astype(str)
        
        return by_date
        
    def _analyze_by_category(self, data):
        """Analyze back orders by category"""
        by_category = data.groupby('category').agg({
            'quantity': ['sum', 'count', 'mean'],
            'item_code': 'nunique'
        }).round(2)
        
        by_category.columns = ['_'.join(col).strip() for col in by_category.columns]
        by_category = by_category.reset_index()
        by_category = by_category.sort_values('quantity_sum', ascending=False)
        
        return by_category
        
    def _analyze_aging(self, data):
        """Analyze aging of back orders"""
        if 'order_date' not in data.columns:
            return pd.DataFrame()
            
        current_date = datetime.now()
        date_data = data.dropna(subset=['order_date']).copy()
        date_data['days_old'] = (current_date - date_data['order_date']).dt.days
        
        # Create aging buckets
        bins = [0, 7, 14, 30, 60, 90, float('inf')]
        labels = ['0-7 days', '8-14 days', '15-30 days', '31-60 days', '61-90 days', '90+ days']
        date_data['age_bucket'] = pd.cut(date_data['days_old'], bins=bins, labels=labels, right=False)
        
        aging = date_data.groupby('age_bucket').agg({
            'quantity': ['sum', 'count'],
            'item_code': 'nunique'
        }).round(2)
        
        aging.columns = ['_'.join(col).strip() for col in aging.columns]
        aging = aging.reset_index()
        
        return aging

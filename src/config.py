"""
Configuration Management Module
"""

import configparser
import os
import logging

class Config:
    def __init__(self, config_file='config.ini'):
        self.config_file = config_file
        self.config = configparser.ConfigParser()
        self.logger = logging.getLogger(__name__)
        
        # Load configuration
        self.load_config()
        
    def load_config(self):
        """Load configuration from file or create default"""
        if os.path.exists(self.config_file):
            try:
                self.config.read(self.config_file)
                self.logger.info(f"Configuration loaded from {self.config_file}")
            except Exception as e:
                self.logger.warning(f"Failed to load config: {e}, using defaults")
                self._create_default_config()
        else:
            self.logger.info("Config file not found, creating default configuration")
            self._create_default_config()
            
    def _create_default_config(self):
        """Create default configuration"""
        self.config['DEFAULT'] = {
            'log_level': 'INFO',
            'max_log_files': '5',
            'max_log_size_mb': '10'
        }
        
        self.config['PROCESSING'] = {
            'validate_data': 'true',
            'remove_duplicates': 'false',
            'default_report_type': 'standard',
            'include_charts': 'true'
        }
        
        self.config['EXCEL'] = {
            'auto_adjust_columns': 'true',
            'freeze_header_row': 'true',
            'apply_styling': 'true',
            'max_chart_items': '10'
        }
        
        self.config['UI'] = {
            'window_width': '600',
            'window_height': '500',
            'theme': 'default'
        }
        
        # Save default configuration
        self.save_config()
        
    def save_config(self):
        """Save configuration to file"""
        try:
            with open(self.config_file, 'w') as f:
                self.config.write(f)
            self.logger.info(f"Configuration saved to {self.config_file}")
        except Exception as e:
            self.logger.error(f"Failed to save configuration: {e}")
            
    def get(self, section, key, fallback=None):
        """Get configuration value"""
        try:
            return self.config.get(section, key)
        except (configparser.NoSectionError, configparser.NoOptionError):
            return fallback
            
    def getboolean(self, section, key, fallback=False):
        """Get boolean configuration value"""
        try:
            return self.config.getboolean(section, key)
        except (configparser.NoSectionError, configparser.NoOptionError, ValueError):
            return fallback
            
    def getint(self, section, key, fallback=0):
        """Get integer configuration value"""
        try:
            return self.config.getint(section, key)
        except (configparser.NoSectionError, configparser.NoOptionError, ValueError):
            return fallback
            
    def set(self, section, key, value):
        """Set configuration value"""
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, key, str(value))

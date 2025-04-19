def _create_customer_group(self):
    """Create the Customer & Salesperson section."""
    group = QGroupBox("Customer & Salesperson")
    layout = QGridLayout(group)
    
    # Create two full-width fields side by side
    layout.addWidget(QLabel("Customer Name:"), 0, 0)
    self.customer_name = QLineEdit()
    self.customer_name.setPlaceholderText("Start typing customer name...")
    layout.addWidget(self.customer_name, 0, 1)
    
    layout.addWidget(QLabel("Salesperson:"), 0, 2)
    self.salesperson = QLineEdit()
    self.salesperson.setPlaceholderText("Start typing salesperson name...")
    layout.addWidget(self.salesperson, 0, 3)
    
    # Make columns stretch properly
    layout.setColumnStretch(1, 1)
    layout.setColumnStretch(3, 1)
    return group

def _create_machine_group(self):
    """Create the Equipment section."""
    group = QGroupBox("Equipment")
    layout = QVBoxLayout(group)
    
    # Input row with fields laid out horizontally
    input_layout = QHBoxLayout()
    
    self.equipment_product_name = QLineEdit()
    self.equipment_product_name.setPlaceholderText("Product Name")
    
    self.equipment_product_code = QLineEdit()
    self.equipment_product_code.setPlaceholderText("Product Code")
    self.equipment_product_code.setReadOnly(True)
    
    self.equipment_manual_stock = QLineEdit()
    self.equipment_manual_stock.setPlaceholderText("Stock # (Manual)")
    
    self.equipment_price = QLineEdit()
    self.equipment_price.setPlaceholderText("$0.00")
    validator_eq = QDoubleValidator(0.0, 9999999.99, 2)
    validator_eq.setNotation(QDoubleValidator.StandardNotation)
    self.equipment_price.setValidator(validator_eq)
    
    self.add_equipment_btn = QPushButton("Add")
    self.add_equipment_btn.setObjectName("add_equipment_btn")
    self.add_equipment_btn.setFixedWidth(60)  # Make the button smaller
    
    # Add widgets with specific stretch factors to match the image
    input_layout.addWidget(self.equipment_product_name, 4)
    input_layout.addWidget(self.equipment_product_code, 2)
    input_layout.addWidget(self.equipment_manual_stock, 2)
    input_layout.addWidget(self.equipment_price, 1)
    input_layout.addWidget(self.add_equipment_btn, 0)
    
    # Equipment list
    self.equipment_list = QListWidget()
    self.equipment_list.setObjectName("equipment_list_widget")
    self.equipment_list.setSelectionMode(QListWidget.SingleSelection)
    self.equipment_list.setMinimumHeight(120)  # Set a minimum height for the list
    
    layout.addLayout(input_layout)
    layout.addWidget(self.equipment_list)
    return group

def _create_trades_group(self):
    """Create the Trades section."""
    group = QGroupBox("Trades")
    layout = QVBoxLayout(group)
    
    # Input row
    input_layout = QHBoxLayout()
    
    self.trade_name = QLineEdit()
    self.trade_name.setPlaceholderText("Trade Item")
    
    self.trade_stock = QLineEdit()
    self.trade_stock.setPlaceholderText("Stock #")
    
    self.trade_amount = QLineEdit()
    self.trade_amount.setPlaceholderText("$0.00")
    validator_tr = QDoubleValidator(0.0, 9999999.99, 2)
    validator_tr.setNotation(QDoubleValidator.StandardNotation)
    self.trade_amount.setValidator(validator_tr)
    
    self.add_trade_btn = QPushButton("Add")
    self.add_trade_btn.setObjectName("add_trade_btn")
    self.add_trade_btn.setFixedWidth(60)  # Make the button smaller
    
    # Add widgets with proportional widths
    input_layout.addWidget(self.trade_name, 4)
    input_layout.addWidget(self.trade_stock, 3)
    input_layout.addWidget(self.trade_amount, 2)
    input_layout.addWidget(self.add_trade_btn, 0)
    
    # Trades list
    self.trade_list = QListWidget()
    self.trade_list.setObjectName("trade_list_widget")
    self.trade_list.setSelectionMode(QListWidget.SingleSelection)
    self.trade_list.setMinimumHeight(120)  # Set a minimum height for the list
    
    # Delete button below the list
    remove_layout = QHBoxLayout()
    self.remove_trade_btn = QPushButton("Remove Selected Trade")
    self.remove_trade_btn.setObjectName("remove_trade_btn")
    remove_layout.addWidget(self.remove_trade_btn)
    remove_layout.addStretch()
    
    layout.addLayout(input_layout)
    layout.addWidget(self.trade_list)
    layout.addLayout(remove_layout)
    
    return group

def _create_parts_group(self):
    """Create the Parts section."""
    group = QGroupBox("Parts")
    layout = QVBoxLayout(group)
    
    # Input row
    input_layout = QHBoxLayout()
    
    # Quantity spinner with up/down buttons
    self.part_quantity = QSpinBox()
    self.part_quantity.setValue(1)
    self.part_quantity.setMinimum(1)
    self.part_quantity.setMaximum(999)
    self.part_quantity.setFixedWidth(60)
    
    self.part_number = QLineEdit()
    self.part_number.setPlaceholderText("Part #")
    
    self.part_name = QLineEdit()
    self.part_name.setPlaceholderText("Part Name")
    
    self.part_location = QComboBox()
    self.part_location.addItems(["", "Camrose", "Killam", "Wainwright", "Provost"])
    
    self.part_charge_to = QLineEdit()
    self.part_charge_to.setPlaceholderText("Charge to:")
    
    self.add_part_btn = QPushButton("Add Part")
    self.add_part_btn.setObjectName("add_part_btn")
    self.add_part_btn.setFixedWidth(80)  # Make the button slightly wider than others
    
    # Add widgets with proportional widths
    input_layout.addWidget(self.part_quantity)
    input_layout.addWidget(self.part_number, 2)
    input_layout.addWidget(self.part_name, 3)
    input_layout.addWidget(self.part_location, 2)
    input_layout.addWidget(self.part_charge_to, 2)
    input_layout.addWidget(self.add_part_btn, 0)
    
    # Parts list
    self.part_list = QListWidget()
    self.part_list.setObjectName("part_list_widget")
    self.part_list.setSelectionMode(QListWidget.SingleSelection)
    self.part_list.setMinimumHeight(120)  # Set a minimum height for the list
    
    # Delete button below the list
    remove_layout = QHBoxLayout()
    self.remove_part_btn = QPushButton("Remove Selected Part")
    self.remove_part_btn.setObjectName("remove_part_btn")
    remove_layout.addWidget(self.remove_part_btn)
    remove_layout.addStretch()
    
    layout.addLayout(input_layout)
    layout.addWidget(self.part_list)
    layout.addLayout(remove_layout)
    
    return group

def _create_work_order_group(self):
    """Create the Work Order & Options section."""
    group = QGroupBox("Work Order & Options")
    layout = QHBoxLayout(group)
    
    # Work Order checkbox
    self.work_order_required = QCheckBox("Work Order Req'd?")
    layout.addWidget(self.work_order_required)
    
    # Charge to field
    self.work_order_charge_to = QLineEdit()
    self.work_order_charge_to.setPlaceholderText("Charge to:")
    layout.addWidget(self.work_order_charge_to)
    
    # Hours field
    self.work_order_hours = QLineEdit()
    self.work_order_hours.setPlaceholderText("Duration (hours)")
    layout.addWidget(self.work_order_hours)
    
    # Multiple CSV lines checkbox
    self.multi_line_csv = QCheckBox("Multiple CSV Lines")
    layout.addWidget(self.multi_line_csv)
    
    # Add stretch to push Paid checkbox to the right
    layout.addStretch()
    
    # Paid checkbox on the far right
    self.paid_checkbox = QCheckBox("Paid")
    self.paid_checkbox.setStyleSheet("font-size: 16px; color: #333;")
    self.paid_checkbox.setChecked(False)
    layout.addWidget(self.paid_checkbox)
    
    return group

def _create_action_buttons_layout(self):
    """Create the bottom action buttons."""
    layout = QHBoxLayout()
    
    # Delete selected line button on the left
    self.delete_line_btn = QPushButton("Delete Selected Line")
    self.delete_line_btn.setObjectName("delete_line_btn")
    layout.addWidget(self.delete_line_btn)
    
    # Add stretch to push other buttons to the right
    layout.addStretch(1)
    
    # Draft buttons
    self.save_draft_btn = QPushButton("Save Draft")
    self.load_draft_btn = QPushButton("Load Draft")
    layout.addWidget(self.save_draft_btn)
    layout.addWidget(self.load_draft_btn)
    
    # Add spacing before the next button group
    layout.addSpacing(20)
    
    # Generate buttons
    self.generate_csv_btn = QPushButton("Gen. CSV & Save SP")
    self.generate_email_btn = QPushButton("Gen. Email")
    self.generate_both_btn = QPushButton("Generate All")
    layout.addWidget(self.generate_csv_btn)
    layout.addSpacing(10)
    layout.addWidget(self.generate_email_btn)
    layout.addSpacing(10)
    layout.addWidget(self.generate_both_btn)
    
    # Add spacing before Reset button
    layout.addSpacing(10)
    
    # Reset button
    self.reset_btn = QPushButton("Reset Form")
    self.reset_btn.setObjectName("reset_btn")
    layout.addWidget(self.reset_btn)
    
    return layout

def init_ui(self):
    """Initialize the UI layout for the Deal Form."""
    self.logger.debug("Initializing DealForm UI...")
    
    # Create main layout
    outer_layout = QVBoxLayout(self)
    outer_layout.setContentsMargins(0, 0, 0, 0)
    
    # Header with green background and logo
    header = QWidget()
    header.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #367C2B, stop:1 #4a9c3e); padding: 15px; border-bottom: 2px solid #2a5d24;")
    header_layout = QHBoxLayout(header)
    
    # Logo on the left
    logo_label = QLabel(self)
    logo_path = get_resource_path('logo.png', getattr(self.config, 'resources_dir', None))
    logo_pixmap = QPixmap(logo_path)
    if not logo_pixmap.isNull():
        logo_label.setPixmap(logo_pixmap.scaled(80, 80, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        self.logger.debug(f"Logo loaded: {logo_path}")
    else:
        logo_label.setText("Logo Missing")
        logo_label.setStyleSheet("color: white;")
        self.logger.warning(f"Logo missing: {logo_path}")
    
    # Title text
    title_label = QLabel("AMS Deal Form")
    title_label.setStyleSheet("color: white; font-size: 40px; font-weight: bold; font-family: Arial;")
    
    header_layout.addWidget(logo_label)
    header_layout.addWidget(title_label)
    header_layout.addStretch()
    
    outer_layout.addWidget(header)
    
    # Create scrollable content area
    scroll_area = QScrollArea(self)
    scroll_area.setWidgetResizable(True)
    scroll_area.setStyleSheet("QScrollArea { border: none; background-color: #f1f3f5; }")
    
    content_container_widget = QWidget()
    content_layout = QVBoxLayout(content_container_widget)
    content_layout.setSpacing(15)
    content_layout.setContentsMargins(15, 15, 15, 15)
    
    # Customer & Salesperson section
    customer_group = QGroupBox("Customer & Salesperson")
    customer_layout = QGridLayout(customer_group)
    
    customer_layout.addWidget(QLabel("Customer Name:"), 0, 0)
    self.customer_name = QLineEdit()
    self.customer_name.setPlaceholderText("Customer Name")
    customer_layout.addWidget(self.customer_name, 0, 1)
    
    customer_layout.addWidget(QLabel("Salesperson:"), 0, 2)
    self.salesperson = QLineEdit()
    self.salesperson.setPlaceholderText("Salesperson")
    customer_layout.addWidget(self.salesperson, 0, 3)
    
    customer_layout.setColumnStretch(1, 1)
    customer_layout.setColumnStretch(3, 1)
    
    content_layout.addWidget(customer_group)
    
    # Equipment section
    equipment_group = QGroupBox("Equipment")
    equipment_layout = QVBoxLayout(equipment_group)
    
    equipment_input = QHBoxLayout()
    self.equipment_product_name = QLineEdit()
    self.equipment_product_name.setPlaceholderText("Product Name")
    
    self.equipment_product_code = QLineEdit()
    self.equipment_product_code.setPlaceholderText("Product Code")
    self.equipment_product_code.setReadOnly(True)
    
    self.equipment_manual_stock = QLineEdit()
    self.equipment_manual_stock.setPlaceholderText("Stock # (Manual)")
    
    self.equipment_price = QLineEdit()
    self.equipment_price.setPlaceholderText("$0.00")
    validator_eq = QDoubleValidator(0.0, 9999999.99, 2)
    validator_eq.setNotation(QDoubleValidator.StandardNotation)
    self.equipment_price.setValidator(validator_eq)
    
    self.add_equipment_btn = QPushButton("Add")
    self.add_equipment_btn.setObjectName("add_equipment_btn")
    
    equipment_input.addWidget(self.equipment_product_name, 4)
    equipment_input.addWidget(self.equipment_product_code, 2)
    equipment_input.addWidget(self.equipment_manual_stock, 2)
    equipment_input.addWidget(self.equipment_price, 1)
    equipment_input.addWidget(self.add_equipment_btn, 0)
    
    self.equipment_list = QListWidget()
    self.equipment_list.setObjectName("equipment_list_widget")
    self.equipment_list.setSelectionMode(QListWidget.SingleSelection)
    
    equipment_layout.addLayout(equipment_input)
    equipment_layout.addWidget(self.equipment_list)
    content_layout.addWidget(equipment_group)
    
    # Trades section
    trades_group = QGroupBox("Trades")
    trades_layout = QVBoxLayout(trades_group)
    
    trades_input = QHBoxLayout()
    self.trade_name = QLineEdit()
    self.trade_name.setPlaceholderText("Trade Item")
    
    self.trade_stock = QLineEdit()
    self.trade_stock.setPlaceholderText("Stock #")
    
    self.trade_amount = QLineEdit()
    self.trade_amount.setPlaceholderText("$0.00")
    validator_tr = QDoubleValidator(0.0, 9999999.99, 2)
    validator_tr.setNotation(QDoubleValidator.StandardNotation)
    self.trade_amount.setValidator(validator_tr)
    
    self.add_trade_btn = QPushButton("Add")
    self.add_trade_btn.setObjectName("add_trade_btn")
    
    trades_input.addWidget(self.trade_name, 4)
    trades_input.addWidget(self.trade_stock, 2)
    trades_input.addWidget(self.trade_amount, 1)
    trades_input.addWidget(self.add_trade_btn, 0)
    
    self.trade_list = QListWidget()
    self.trade_list.setObjectName("trade_list_widget")
    self.trade_list.setSelectionMode(QListWidget.SingleSelection)
    
    remove_trade_layout = QHBoxLayout()
    self.remove_trade_btn = QPushButton("Remove Selected Trade")
    self.remove_trade_btn.setObjectName("remove_trade_btn")
    remove_trade_layout.addWidget(self.remove_trade_btn)
    remove_trade_layout.addStretch()
    
    trades_layout.addLayout(trades_input)
    trades_layout.addWidget(self.trade_list)
    trades_layout.addLayout(remove_trade_layout)
    content_layout.addWidget(trades_group)
    
    # Parts section
    parts_group = QGroupBox("Parts")
    parts_layout = QVBoxLayout(parts_group)
    
    parts_input = QHBoxLayout()
    self.part_quantity = QSpinBox()
    self.part_quantity.setValue(1)
    self.part_quantity.setMinimum(1)
    self.part_quantity.setMaximum(999)
    self.part_quantity.setFixedWidth(60)
    
    self.part_number = QLineEdit()
    self.part_number.setPlaceholderText("Part #")
    
    self.part_name = QLineEdit()
    self.part_name.setPlaceholderText("Part Name")
    
    self.part_location = QComboBox()
    self.part_location.addItems(["", "Camrose", "Killam", "Wainwright", "Provost"])
    
    self.part_charge_to = QLineEdit()
    self.part_charge_to.setPlaceholderText("Charge to:")
    
    self.add_part_btn = QPushButton("Add Part")
    self.add_part_btn.setObjectName("add_part_btn")
    
    parts_input.addWidget(self.part_quantity)
    parts_input.addWidget(self.part_number, 2)
    parts_input.addWidget(self.part_name, 2)
    parts_input.addWidget(self.part_location, 1)
    parts_input.addWidget(self.part_charge_to, 2)
    parts_input.addWidget(self.add_part_btn, 0)
    
    self.part_list = QListWidget()
    self.part_list.setObjectName("part_list_widget")
    self.part_list.setSelectionMode(QListWidget.SingleSelection)
    
    remove_part_layout = QHBoxLayout()
    self.remove_part_btn = QPushButton("Remove Selected Part")
    self.remove_part_btn.setObjectName("remove_part_btn")
    remove_part_layout.addWidget(self.remove_part_btn)
    remove_part_layout.addStretch()
    
    parts_layout.addLayout(parts_input)
    parts_layout.addWidget(self.part_list)
    parts_layout.addLayout(remove_part_layout)
    content_layout.addWidget(parts_group)
    
    # Work Order & Options section
    work_order_group = QGroupBox("Work Order & Options")
    work_order_layout = QHBoxLayout(work_order_group)
    
    self.work_order_required = QCheckBox("Work Order Req'd?")
    work_order_layout.addWidget(self.work_order_required)
    
    self.work_order_charge_to = QLineEdit()
    self.work_order_charge_to.setPlaceholderText("Charge to:")
    work_order_layout.addWidget(self.work_order_charge_to)
    
    self.work_order_hours = QLineEdit()
    self.work_order_hours.setPlaceholderText("Duration (hours)")
    work_order_layout.addWidget(self.work_order_hours)
    
    self.multi_line_csv = QCheckBox("Multiple CSV Lines")
    work_order_layout.addWidget(self.multi_line_csv)
    
    work_order_layout.addStretch()
    
    self.paid_checkbox = QCheckBox("Paid")
    self.paid_checkbox.setStyleSheet("font-size: 16px; color: #333;")
    self.paid_checkbox.setChecked(False)
    work_order_layout.addWidget(self.paid_checkbox)
    
    content_layout.addWidget(work_order_group)
    
    # Add a stretch to push buttons to the bottom
    content_layout.addStretch(1)
    
    # Action buttons at the bottom
    buttons_layout = QHBoxLayout()
    
    self.delete_line_btn = QPushButton("Delete Selected Line")
    self.delete_line_btn.setObjectName("delete_line_btn")
    buttons_layout.addWidget(self.delete_line_btn)
    
    buttons_layout.addStretch(1)
    
    self.save_draft_btn = QPushButton("Save Draft")
    self.load_draft_btn = QPushButton("Load Draft")
    buttons_layout.addWidget(self.save_draft_btn)
    buttons_layout.addWidget(self.load_draft_btn)
    
    buttons_layout.addSpacing(20)
    
    self.generate_csv_btn = QPushButton("Gen. CSV & Save SP")
    self.generate_email_btn = QPushButton("Gen. Email")
    self.generate_both_btn = QPushButton("Generate All")
    buttons_layout.addWidget(self.generate_csv_btn)
    buttons_layout.addSpacing(10)
    buttons_layout.addWidget(self.generate_email_btn)
    buttons_layout.addSpacing(10)
    buttons_layout.addWidget(self.generate_both_btn)
    
    buttons_layout.addSpacing(10)
    
    self.reset_btn = QPushButton("Reset Form")
    self.reset_btn.setObjectName("reset_btn")
    buttons_layout.addWidget(self.reset_btn)
    
    content_layout.addLayout(buttons_layout)
    
    scroll_area.setWidget(content_container_widget)
    outer_layout.addWidget(scroll_area)
    
    self.logger.debug(f"{self.MODULE_DISPLAY_NAME} UI created.")# File: modules/deal_form_module.py
# Desc: Module for creating and managing equipment deals

import os
import re
import json
import logging
import traceback
from datetime import datetime
from functools import partial
import csv
import io
import webbrowser
from urllib.parse import quote

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QFormLayout, 
                             QMainWindow, QDockWidget, QGridLayout, QGroupBox, 
                             QLabel, QLineEdit, QPushButton, QComboBox, QCheckBox, 
                             QTextEdit, QPlainTextEdit, QMessageBox, QApplication, 
                             QCompleter, QListWidget, QListWidgetItem, QSpinBox, 
                             QDoubleSpinBox, QFileDialog, QTableWidget, QTableWidgetItem,
                             QHeaderView, QAbstractItemView, QScrollArea, QSizePolicy, 
                             QFrame, QDateEdit, QInputDialog, QDialog, QDialogButtonBox)
from PyQt5.QtCore import (Qt, QDate, pyqtSignal, QVariant, QAbstractTableModel, 
                          pyqtSlot, QTimer, QSize, QMimeData, QStringListModel)
from PyQt5.QtGui import (QFont, QPalette, QColor, QIcon, QDoubleValidator, 
                         QPixmap, QClipboard)

# Try importing BaseModule
try:
    from ui.base_module import BaseModule
except ImportError:
    print("CRITICAL ERROR: BaseModule could not be imported in DealFormModule.")
    # Fallback dummy class
    class BaseModule(QWidget):
        def __init__(self, main_window=None): 
            super().__init__()
            self.main_window = main_window
        def get_title(self): return "Deal Form (Import Error)"
        def get_icon_name(self): return None
        def save_state(self): pass
        def refresh(self): pass
        def close(self): pass

# Try importing Worker
try:
    from utils.worker import Worker
except ImportError:
    Worker = None
    print("WARNING: Worker class not found. Background saving to SharePoint will fail.")

# Try importing get_resource_path
try:
    from utils.general_utils import get_resource_path
except ImportError:
    print("WARNING: get_resource_path not found in general_utils. Using dummy function.")
    # Fallback
    def get_resource_path(filename, res_dir):
        if res_dir: return os.path.join(res_dir, filename)
        return filename

# Constants
DRAFT_FILENAME = "deal_draft.json"
RECENT_DEALS_FILENAME = "recent_deals.json"
MAX_RECENT_DEALS = 10

# --- Data Loading Helper Class ---
class DataLoader:
    """Handles loading data from CSV files."""
    def __init__(self, data_path, logger):
        self.data_path = data_path
        self.logger = logger.getChild("DataLoader") if logger else logging.getLogger("DataLoader")
        # Data storage (initially empty)
        self.products_dict = {}
        self.parts_dict = {}
        self.customers_list = []
        self.salesmen_emails = {}
        # Flags to track if data has been loaded
        self._products_loaded = False
        self._parts_loaded = False
        self._customers_loaded = False
        self._salesmen_loaded = False

    def _load_csv_generic(self, filename, key_column, value_column=None, is_dict=True):
        """Load data from a CSV file."""
        data = {} if is_dict else []
        if not self.data_path: 
            self.logger.error(f"{filename}: Data path not set.")
            return data
            
        file_path = os.path.join(self.data_path, filename)
        encodings = ['utf-8-sig', 'utf-8', 'latin1', 'windows-1252']
        loaded = False
        rows_processed = 0
        
        self.logger.debug(f"Attempting to load {filename}, expecting key '{key_column}'...")
        for encoding in encodings:
            try:
                with open(file_path, mode='r', newline='', encoding=encoding) as infile:
                    reader = csv.DictReader(infile)
                    actual_headers = reader.fieldnames or []
                    if not actual_headers:
                        self.logger.warning(f"{filename}: No headers found with encoding {encoding}.")
                        continue
                        
                    key_header = next((h for h in actual_headers if h.lower().strip() == key_column.lower()), None)
                    value_header = None
                    if value_column: 
                        value_header = next((h for h in actual_headers if h.lower().strip() == value_column.lower()), None)
                    
                    if not key_header:
                        self.logger.warning(f"{filename}: Key column '{key_column}' not found in headers {actual_headers} with encoding {encoding}.")
                        continue
                        
                    if value_column and not value_header and is_dict:
                        self.logger.warning(f"{filename}: Specified value column '{value_column}' not found, but loading full row dict for key '{key_header}'.")
                    elif value_column and not value_header and not is_dict:
                        self.logger.warning(f"{filename}: Specified value column '{value_column}' not found for list output. Skipping encoding {encoding}.")
                        continue

                    temp_data = {}
                    temp_list = []
                    rows_processed = 0
                    
                    for row_num, row in enumerate(reader):
                        key = row.get(key_header, "").strip()
                        if not key:
                            self.logger.debug(f"{filename}: Skipping empty key in row {row_num+1}.")
                            continue
                            
                        rows_processed += 1
                        if is_dict:
                            if value_header:
                                temp_data[key] = row.get(value_header, "").strip()
                            elif value_column is None: # Specific handling for products_dict (whole row)
                                code_h = next((k for k in row if k.lower().strip() == 'productcode'), None)
                                price_h = next((k for k in row if k.lower().strip() == 'price'), None)
                                temp_data[key] = (row.get(code_h, ""), row.get(price_h, "0.00"))
                        else: # is_list
                            temp_list.append(key)

                    if is_dict: 
                        data = temp_data
                    else: 
                        data = temp_list

                    loaded = True
                    self.logger.info(f"Successfully loaded {rows_processed} rows from {filename} with encoding {encoding}.")
                    break
                    
            except FileNotFoundError: 
                self.logger.error(f"{filename}: Not found at {file_path}")
                return data
            except UnicodeDecodeError: 
                continue
            except Exception as e: 
                self.logger.error(f"Error loading {filename} ({encoding}): {e}")
                continue
                
        if not loaded: 
            self.logger.error(f"{filename}: Failed to load with any encoding or required columns missing.")
            
        self.logger.debug(f"Finished loading {filename}. Loaded items: {len(data)}")
        return data

    def get_products(self, force_reload=False):
        """Get product data, loading from CSV if necessary."""
        if not self._products_loaded or force_reload:
            self.logger.info("Loading products data...")
            self.products_dict = self._load_csv_generic('products.csv', "ProductName", None, is_dict=True)
            self._products_loaded = True
            self.logger.info(f"Products data loaded: {len(self.products_dict)} items.")
        return self.products_dict

    def get_parts(self, force_reload=False):
        """Get parts data, loading from CSV if necessary."""
        if not self._parts_loaded or force_reload:
            self.logger.info("Loading parts data...")
            self.parts_dict = self._load_csv_generic('parts.csv', "Part Name", "Part Number", is_dict=True)
            self._parts_loaded = True
            self.logger.info(f"Parts data loaded: {len(self.parts_dict)} items.")
        return self.parts_dict

    def get_customers(self, force_reload=False):
        """Get customer data, loading from CSV if necessary."""
        if not self._customers_loaded or force_reload:
            self.logger.info("Loading customers data...")
            self.customers_list = self._load_csv_generic('customers.csv', "Name", None, is_dict=False)
            self._customers_loaded = True
            self.logger.info(f"Customers data loaded: {len(self.customers_list)} items.")
        return self.customers_list

    def get_salesmen(self, force_reload=False):
        """Get salesmen data, loading from CSV if necessary."""
        if not self._salesmen_loaded or force_reload:
            self.logger.info("Loading salesmen data...")
            self.salesmen_emails = self._load_csv_generic('salesmen.csv', "Name", "Email", is_dict=True)
            self._salesmen_loaded = True
            self.logger.info(f"Salesmen data loaded: {len(self.salesmen_emails)} items.")
        return self.salesmen_emails

    def ensure_all_loaded(self, force_reload=False):
        """Load all data if not already loaded."""
        self.get_customers(force_reload)
        self.get_products(force_reload)
        self.get_parts(force_reload)
        self.get_salesmen(force_reload)


# Main module class
class DealFormModule(BaseModule):
    """Module for creating and managing equipment deals."""
    MODULE_DISPLAY_NAME = "Deal Form"
    MODULE_ICON_NAME = "dealform_icon.png"
    deal_saved = pyqtSignal(dict)

    def __init__(self, main_window=None, sharepoint_manager=None):
        super().__init__(main_window=main_window)
        self.setObjectName("DealFormModule")

        parent_logger = getattr(self.main_window, 'logger', None)
        self.logger = parent_logger.getChild("DealForm") if parent_logger else logging.getLogger(__name__).getChild("DealForm")
        self.logger.info(f"Initializing {self.MODULE_DISPLAY_NAME}...")

        self.sharepoint_manager = sharepoint_manager
        self.config = getattr(self.main_window, 'config', None)
        self.data_path = getattr(self.config, 'data_dir', None)
        self.cache_path = getattr(self.config, 'cache_dir', None)

        if not self.config: self.logger.error("Config object not found!")
        if not self.data_path: self.logger.error("Data directory path not found!")
        if not self.cache_path: self.logger.error("Cache directory path not found!")
        if not self.sharepoint_manager: self.logger.warning("SharePointManager instance not provided!")

        self.current_form_data = {}
        self.data_loader = DataLoader(self.data_path, self.logger)

        # Initialize UI elements to None
        self._init_ui_elements_to_none()
        self.init_ui()
        self._update_completers(initial_load=True)
        self._connect_internal_signals()
        self._connect_button_signals()

        self.csv_lines = []
        self.last_charge_to = ""
        self.apply_styles()

        QTimer.singleShot(100, self._attempt_load_draft)

        self.logger.info(f"{self.MODULE_DISPLAY_NAME} initialization complete.")

    def _init_ui_elements_to_none(self):
        """Initialize all UI elements to None."""
        self.customer_name = None
        self.salesperson = None
        self.equipment_product_name = None
        self.equipment_product_code = None
        self.equipment_manual_stock = None
        self.equipment_price = None
        self.equipment_list = None
        self.trade_name = None
        self.trade_stock = None
        self.trade_amount = None
        self.trade_list = None
        self.part_quantity = None
        self.part_number = None
        self.part_name = None
        self.part_location = None
        self.part_charge_to = None
        self.part_list = None
        self.work_order_required = None
        self.work_order_charge_to = None
        self.work_order_hours = None
        self.multi_line_csv = None
        self.paid_checkbox = None
        self.save_draft_btn = None
        self.load_draft_btn = None
        self.generate_csv_btn = None
        self.generate_email_btn = None
        self.generate_both_btn = None
        self.reset_btn = None
        self.notes_edit = None
        self.add_equipment_btn = None
        self.add_part_btn = None
        self.add_trade_btn = None
        self.remove_part_btn = None
        self.remove_trade_btn = None
        # Completers
        self.customer_completer = None
        self.salesperson_completer = None
        self.equipment_completer = None
        self.trade_completer = None
        self.part_name_completer = None
        self.part_number_completer = None

    def init_ui(self):
        """Initialize the UI layout for the Deal Form."""
        self.logger.debug("Initializing DealForm UI...")
        
        # Create main layout
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        
        # Header with green background and logo
        header = QWidget()
        header.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #367C2B, stop:1 #4a9c3e); padding: 15px; border-bottom: 2px solid #2a5d24;")
        header_layout = QHBoxLayout(header)
        
        # Logo on the left
        logo_label = QLabel(self)
        logo_path = get_resource_path('logo.png', getattr(self.config, 'resources_dir', None))
        logo_pixmap = QPixmap(logo_path)
        if not logo_pixmap.isNull():
            logo_label.setPixmap(logo_pixmap.scaled(80, 80, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            self.logger.debug(f"Logo loaded: {logo_path}")
        else:
            logo_label.setText("Logo Missing")
            logo_label.setStyleSheet("color: white;")
            self.logger.warning(f"Logo missing: {logo_path}")
        
        # Title text
        title_label = QLabel("AMS Deal Form")
        title_label.setStyleSheet("color: white; font-size: 40px; font-weight: bold; font-family: Arial;")
        
        header_layout.addWidget(logo_label)
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        
        outer_layout.addWidget(header)
        
        # Create scrollable content area
        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("QScrollArea { border: none; background-color: #f1f3f5; }")
        
        content_container_widget = QWidget()
        content_layout = QVBoxLayout(content_container_widget)
        content_layout.setSpacing(15)
        content_layout.setContentsMargins(15, 15, 15, 15)
        
        # Customer & Salesperson section
        customer_group = QGroupBox("Customer & Salesperson")
        customer_layout = QGridLayout(customer_group)
        
        customer_layout.addWidget(QLabel("Customer Name:"), 0, 0)
        self.customer_name = QLineEdit()
        self.customer_name.setPlaceholderText("Customer Name")
        customer_layout.addWidget(self.customer_name, 0, 1)
        
        customer_layout.addWidget(QLabel("Salesperson:"), 0, 2)
        self.salesperson = QLineEdit()
        self.salesperson.setPlaceholderText("Salesperson")
        customer_layout.addWidget(self.salesperson, 0, 3)
        
        customer_layout.setColumnStretch(1, 1)
        customer_layout.setColumnStretch(3, 1)
        
        content_layout.addWidget(customer_group)
        
        # Equipment section
        equipment_group = QGroupBox("Equipment")
        equipment_layout = QVBoxLayout(equipment_group)
        
        equipment_input = QHBoxLayout()
        self.equipment_product_name = QLineEdit()
        self.equipment_product_name.setPlaceholderText("Product Name")
        
        self.equipment_product_code = QLineEdit()
        self.equipment_product_code.setPlaceholderText("Product Code")
        self.equipment_product_code.setReadOnly(True)
        
        self.equipment_manual_stock = QLineEdit()
        self.equipment_manual_stock.setPlaceholderText("Stock # (Manual)")
        
        self.equipment_price = QLineEdit()
        self.equipment_price.setPlaceholderText("$0.00")
        validator_eq = QDoubleValidator(0.0, 9999999.99, 2)
        validator_eq.setNotation(QDoubleValidator.StandardNotation)
        self.equipment_price.setValidator(validator_eq)
        
        self.add_equipment_btn = QPushButton("Add")
        self.add_equipment_btn.setObjectName("add_equipment_btn")
        
        equipment_input.addWidget(self.equipment_product_name, 4)
        equipment_input.addWidget(self.equipment_product_code, 2)
        equipment_input.addWidget(self.equipment_manual_stock, 2)
        equipment_input.addWidget(self.equipment_price, 1)
        equipment_input.addWidget(self.add_equipment_btn, 0)
        
        self.equipment_list = QListWidget()
        self.equipment_list.setObjectName("equipment_list_widget")
        self.equipment_list.setSelectionMode(QListWidget.SingleSelection)
        
        equipment_layout.addLayout(equipment_input)
        equipment_layout.addWidget(self.equipment_list)
        content_layout.addWidget(equipment_group)
        
        # Trades section
        trades_group = QGroupBox("Trades")
        trades_layout = QVBoxLayout(trades_group)
        
        trades_input = QHBoxLayout()
        self.trade_name = QLineEdit()
        self.trade_name.setPlaceholderText("Trade Item")
        
        self.trade_stock = QLineEdit()
        self.trade_stock.setPlaceholderText("Stock #")
        
        self.trade_amount = QLineEdit()
        self.trade_amount.setPlaceholderText("$0.00")
        validator_tr = QDoubleValidator(0.0, 9999999.99, 2)
        validator_tr.setNotation(QDoubleValidator.StandardNotation)
        self.trade_amount.setValidator(validator_tr)
        
        self.add_trade_btn = QPushButton("Add")
        self.add_trade_btn.setObjectName("add_trade_btn")
        
        trades_input.addWidget(self.trade_name, 4)
        trades_input.addWidget(self.trade_stock, 2)
        trades_input.addWidget(self.trade_amount, 1)
        trades_input.addWidget(self.add_trade_btn, 0)
        
        self.trade_list = QListWidget()
        self.trade_list.setObjectName("trade_list_widget")
        self.trade_list.setSelectionMode(QListWidget.SingleSelection)
        
        remove_trade_layout = QHBoxLayout()
        self.remove_trade_btn = QPushButton("Remove Selected Trade")
        self.remove_trade_btn.setObjectName("remove_trade_btn")
        remove_trade_layout.addWidget(self.remove_trade_btn)
        remove_trade_layout.addStretch()
        
        trades_layout.addLayout(trades_input)
        trades_layout.addWidget(self.trade_list)
        trades_layout.addLayout(remove_trade_layout)
        content_layout.addWidget(trades_group)
        
        # Parts section
        parts_group = QGroupBox("Parts")
        parts_layout = QVBoxLayout(parts_group)
        
        parts_input = QHBoxLayout()
        self.part_quantity = QSpinBox()
        self.part_quantity.setValue(1)
        self.part_quantity.setMinimum(1)
        self.part_quantity.setMaximum(999)
        self.part_quantity.setFixedWidth(60)
        
        self.part_number = QLineEdit()
        self.part_number.setPlaceholderText("Part #")
        
        self.part_name = QLineEdit()
        self.part_name.setPlaceholderText("Part Name")
        
        self.part_location = QComboBox()
        self.part_location.addItems(["", "Camrose", "Killam", "Wainwright", "Provost"])
        
        self.part_charge_to = QLineEdit()
        self.part_charge_to.setPlaceholderText("Charge to:")
        
        self.add_part_btn = QPushButton("Add Part")
        self.add_part_btn.setObjectName("add_part_btn")
        
        parts_input.addWidget(self.part_quantity)
        parts_input.addWidget(self.part_number, 2)
        parts_input.addWidget(self.part_name, 2)
        parts_input.addWidget(self.part_location, 1)
        parts_input.addWidget(self.part_charge_to, 2)
        parts_input.addWidget(self.add_part_btn, 0)
        
        self.part_list = QListWidget()
        self.part_list.setObjectName("part_list_widget")
        self.part_list.setSelectionMode(QListWidget.SingleSelection)
        
        remove_part_layout = QHBoxLayout()
        self.remove_part_btn = QPushButton("Remove Selected Part")
        self.remove_part_btn.setObjectName("remove_part_btn")
        remove_part_layout.addWidget(self.remove_part_btn)
        remove_part_layout.addStretch()
        
        parts_layout.addLayout(parts_input)
        parts_layout.addWidget(self.part_list)
        parts_layout.addLayout(remove_part_layout)
        content_layout.addWidget(parts_group)
        
        # Notes section
        notes_group = QGroupBox("Notes")
        notes_layout = QVBoxLayout(notes_group)
        self.notes_edit = QPlainTextEdit()
        self.notes_edit.setPlaceholderText("Enter deal notes here...")
        notes_layout.addWidget(self.notes_edit)
        content_layout.addWidget(notes_group)
        
        # Work Order & Options section
        work_order_group = QGroupBox("Work Order & Options")
        work_order_layout = QHBoxLayout(work_order_group)
        
        self.work_order_required = QCheckBox("Work Order Req'd?")
        work_order_layout.addWidget(self.work_order_required)
        
        self.work_order_charge_to = QLineEdit()
        self.work_order_charge_to.setPlaceholderText("Charge to:")
        work_order_layout.addWidget(self.work_order_charge_to)
        
        self.work_order_hours = QLineEdit()
        self.work_order_hours.setPlaceholderText("Duration (hours)")
        work_order_layout.addWidget(self.work_order_hours)
        
        self.multi_line_csv = QCheckBox("Multiple CSV Lines")
        work_order_layout.addWidget(self.multi_line_csv)
        
        work_order_layout.addStretch()
        
        self.paid_checkbox = QCheckBox("Paid")
        self.paid_checkbox.setStyleSheet("font-size: 16px; color: #333;")
        self.paid_checkbox.setChecked(False)
        work_order_layout.addWidget(self.paid_checkbox)
        
        content_layout.addWidget(work_order_group)
        
        # Add a stretch to push buttons to the bottom
        content_layout.addStretch(1)
        
        # Action buttons at the bottom
        buttons_layout = QHBoxLayout()
        
        self.delete_line_btn = QPushButton("Delete Selected Line")
        self.delete_line_btn.setObjectName("delete_line_btn")
        buttons_layout.addWidget(self.delete_line_btn)
        
        buttons_layout.addStretch(1)
        
        self.save_draft_btn = QPushButton("Save Draft")
        self.load_draft_btn = QPushButton("Load Draft")
        buttons_layout.addWidget(self.save_draft_btn)
        buttons_layout.addWidget(self.load_draft_btn)
        
        buttons_layout.addSpacing(20)
        
        self.generate_csv_btn = QPushButton("Gen. CSV & Save SP")
        self.generate_email_btn = QPushButton("Gen. Email")
        self.generate_both_btn = QPushButton("Generate All")
        buttons_layout.addWidget(self.generate_csv_btn)
        buttons_layout.addSpacing(10)
        buttons_layout.addWidget(self.generate_email_btn)
        buttons_layout.addSpacing(10)
        buttons_layout.addWidget(self.generate_both_btn)
        
        buttons_layout.addSpacing(10)
        
        self.reset_btn = QPushButton("Reset Form")
        self.reset_btn.setObjectName("reset_btn")
        buttons_layout.addWidget(self.reset_btn)
        
        content_layout.addLayout(buttons_layout)
        
        scroll_area.setWidget(content_container_widget)
        outer_layout.addWidget(scroll_area)
        
        self.logger.debug(f"{self.MODULE_DISPLAY_NAME} UI created.")

    def _update_completers(self, force_reload=False, initial_load=False):
        """Update auto-completers for input fields."""
        self.logger.debug(f"Updating completers... Force Reload: {force_reload}, Initial Load: {initial_load}")
        try:
            customer_data = self.data_loader.get_customers(force_reload or initial_load)
            salesmen_data = self.data_loader.get_salesmen(force_reload or initial_load)
            product_data = self.data_loader.get_products(force_reload or initial_load)
            parts_data = self.data_loader.get_parts(force_reload or initial_load)

            # Create models for completers
            self.customer_model = QStringListModel(customer_data)
            self.salesperson_model = QStringListModel(list(salesmen_data.keys()))
            product_keys = list(product_data.keys())
            self.product_model = QStringListModel(product_keys)
            self.part_name_model = QStringListModel(list(parts_data.keys()))
            self.part_number_model = QStringListModel(list(map(str, parts_data.values())))

            # Setup completers
            if self.customer_name and not self.customer_completer:
                self.customer_completer = QCompleter(self.customer_model, self)
                self.customer_completer.setCaseSensitivity(Qt.CaseInsensitive)
                self.customer_completer.setFilterMode(Qt.MatchContains)
                self.customer_name.setCompleter(self.customer_completer)
                self.logger.debug(f"Customer completer set with {len(customer_data)} items.")

            if self.salesperson and not self.salesperson_completer:
                self.salesperson_completer = QCompleter(self.salesperson_model, self)
                self.salesperson_completer.setCaseSensitivity(Qt.CaseInsensitive)
                self.salesperson_completer.setFilterMode(Qt.MatchContains)
                self.salesperson.setCompleter(self.salesperson_completer)
                self.logger.debug(f"Salesperson completer set with {len(salesmen_data)} items.")

            if self.equipment_product_name and not self.equipment_completer:
                self.equipment_completer = QCompleter(self.product_model, self)
                self.equipment_completer.setCaseSensitivity(Qt.CaseInsensitive)
                self.equipment_completer.setFilterMode(Qt.MatchContains)
                self.equipment_product_name.setCompleter(self.equipment_completer)
                self.connect_completer_signals(self.equipment_completer, self.on_equipment_selected)
                self.logger.debug(f"Equipment completer set with {len(product_keys)} items.")

            if self.trade_name and not self.trade_completer:
                self.trade_completer = QCompleter(self.product_model, self)
                self.trade_completer.setCaseSensitivity(Qt.CaseInsensitive)
                self.trade_completer.setFilterMode(Qt.MatchContains)
                self.trade_name.setCompleter(self.trade_completer)
                self.connect_completer_signals(self.trade_completer, self.on_trade_selected)
                self.logger.debug("Trade name completer set.")

            if self.part_name and not self.part_name_completer:
                self.part_name_completer = QCompleter(self.part_name_model, self)
                self.part_name_completer.setCaseSensitivity(Qt.CaseInsensitive)
                self.part_name_completer.setFilterMode(Qt.MatchContains)
                self.part_name.setCompleter(self.part_name_completer)
                self.connect_completer_signals(self.part_name_completer, self.on_part_selected)
                self.logger.debug(f"Part name completer set with {len(parts_data)} items.")

            if self.part_number and not self.part_number_completer:
                self.part_number_completer = QCompleter(self.part_number_model, self)
                self.part_number_completer.setCaseSensitivity(Qt.CaseInsensitive)
                self.part_number_completer.setFilterMode(Qt.MatchStartsWith)
                self.part_number.setCompleter(self.part_number_completer)
                self.connect_completer_signals(self.part_number_completer, self.on_part_number_selected)
                self.logger.debug("Part number completer set.")
        except Exception as e:
            self.logger.error(f"Error updating completers: {e}", exc_info=True)

    def connect_completer_signals(self, completer, slot_func):
        """Connect a completer's activated signal to a slot function."""
        if completer:
            try: 
                completer.activated.disconnect(slot_func)
            except TypeError: 
                pass
            completer.activated.connect(slot_func)
            self.logger.debug(f"Connected completer signal to {slot_func.__name__}")
        else:
            self.logger.warning(f"Attempted to connect signals for a non-existent completer")

    def _connect_internal_signals(self):
        """Connect all internal signals for data tracking."""
        self.logger.debug("Connecting internal signals...")
        
        # Helper to connect a widget signal to internal data update
        connect = lambda w, sig, key, func: self._check_and_connect(w, sig, 
                                                        partial(self._update_internal_data, key, func))
                                                        
        # Connect basic field changes to internal data
        connect(self.customer_name, 'textChanged', 'customer_name', lambda: self.customer_name.text())
        connect(self.salesperson, 'textChanged', 'salesperson', lambda: self.salesperson.text())
        connect(self.equipment_product_name, 'textChanged', 'equipment_product_name', 
                lambda: self.equipment_product_name.text())
        connect(self.equipment_product_code, 'textChanged', 'equipment_product_code', 
                lambda: self.equipment_product_code.text())
        connect(self.equipment_manual_stock, 'textChanged', 'equipment_manual_stock', 
                lambda: self.equipment_manual_stock.text())
        connect(self.equipment_price, 'textChanged', 'equipment_price_str', 
                lambda: self.equipment_price.text())
        connect(self.trade_name, 'textChanged', 'trade_name', lambda: self.trade_name.text())
        connect(self.trade_stock, 'textChanged', 'trade_stock', lambda: self.trade_stock.text())
        connect(self.trade_amount, 'textChanged', 'trade_amount_str', lambda: self.trade_amount.text())
        connect(self.part_quantity, 'valueChanged', 'part_quantity', lambda: self.part_quantity.value())
        connect(self.part_number, 'textChanged', 'part_number', lambda: self.part_number.text())
        connect(self.part_name, 'textChanged', 'part_name', lambda: self.part_name.text())
        connect(self.part_location, 'currentTextChanged', 'part_location', 
                lambda: self.part_location.currentText())
        connect(self.part_charge_to, 'textChanged', 'part_charge_to', lambda: self.part_charge_to.text())
        connect(self.work_order_required, 'stateChanged', 'work_order_required', 
                lambda: self.work_order_required.isChecked())
        connect(self.work_order_charge_to, 'textChanged', 'work_order_charge_to', 
                lambda: self.work_order_charge_to.text())
        connect(self.work_order_hours, 'textChanged', 'work_order_hours', 
                lambda: self.work_order_hours.text())
        connect(self.multi_line_csv, 'stateChanged', 'multi_line_csv', 
                lambda: self.multi_line_csv.isChecked())
        connect(self.paid_checkbox, 'stateChanged', 'paid', lambda: self.paid_checkbox.isChecked())
        
        # Connect notes field - handle both QPlainTextEdit and QTextEdit
        notes_sig = 'textChanged'
        notes_func = lambda: self.notes_edit.toPlainText() if isinstance(self.notes_edit, QPlainTextEdit) else \
                    self.notes_edit.toHtml() if isinstance(self.notes_edit, QTextEdit) else ""
        connect(self.notes_edit, notes_sig, 'notes', notes_func)

    def _connect_button_signals(self):
        """Connect all button click signals."""
        self.logger.debug("Connecting button signals...")
        
        # Connect all buttons to their respective handlers
        self._check_and_connect(self.save_draft_btn, 'clicked', self.save_draft)
        self._check_and_connect(self.load_draft_btn, 'clicked', self._attempt_load_draft)
        self._check_and_connect(self.generate_csv_btn, 'clicked', self.save_to_sharepoint)
        self._check_and_connect(self.generate_email_btn, 'clicked', self.generate_email)
        self._check_and_connect(self.generate_both_btn, 'clicked', self.generate_all)
        self._check_and_connect(self.reset_btn, 'clicked', self.reset_form)
        
        # Connect item add/remove buttons
        self._check_and_connect(self.add_equipment_btn, 'clicked', partial(self.add_item, "equipment"))
        self._check_and_connect(self.add_part_btn, 'clicked', partial(self.add_item, "part"))
        self._check_and_connect(self.add_trade_btn, 'clicked', partial(self.add_item, "trade"))
        self._check_and_connect(self.remove_part_btn, 'clicked', 
                               lambda: self.delete_selected_list_item(list_widget=self.part_list, item_type="Part"))
        self._check_and_connect(self.remove_trade_btn, 'clicked', 
                               lambda: self.delete_selected_list_item(list_widget=self.trade_list, item_type="Trade"))
        self._check_and_connect(self.delete_line_btn, 'clicked', self.delete_selected_line)
        
        # Connect list double-click for editing
        self._check_and_connect(self.equipment_list, 'itemDoubleClicked', self.edit_equipment_item)
        self._check_and_connect(self.trade_list, 'itemDoubleClicked', self.edit_trade_item)
        self._check_and_connect(self.part_list, 'itemDoubleClicked', self.edit_part_item)
        
        # Connect formatting signals
        self._check_and_connect(self.equipment_price, 'editingFinished', self.format_price)
        self._check_and_connect(self.trade_amount, 'editingFinished', self.format_amount)

    def _check_and_connect(self, widget, signal_name, slot_func):
        """Safely connect a widget's signal to a slot function."""
        if widget and hasattr(widget, signal_name):
            signal = getattr(widget, signal_name)
            try: 
                signal.disconnect()
            except TypeError: 
                pass
            try: 
                signal.connect(slot_func)
                return True
            except Exception as e: 
                self.logger.error(f"Connect Fail: Widget={widget.objectName() if hasattr(widget,'objectName') else widget}, sig='{signal_name}', slot={slot_func}: {e}")
        return False

    @pyqtSlot()
    def _update_internal_data(self, key, value_func):
        """Update internal data dictionary with a key/value pair."""
        try: 
            self.current_form_data[key] = value_func()
        except Exception as e: 
            self.logger.error(f"Error updating key '{key}': {e}")

    # --- Autocompletion Handlers ---
    def on_equipment_selected(self, text):
        """Handle equipment selection from autocomplete."""
        self.logger.debug(f"Equipment selected via completer: '{text}'")
        self.products_dict = self.data_loader.get_products()
        product_data = self.products_dict.get(text, (None, None))
        code, price_str = product_data
        
        if code is not None and self.equipment_product_code and self.equipment_price:
            curr_price = self.equipment_price.text().replace('$', '').replace(',', '')
            self.equipment_product_code.setText(code)
            try:
                if not curr_price or float(curr_price) == 0.0:
                    p_val = float(re.sub(r'[^\d.-]', '', price_str)) if price_str else 0.0
                    self.equipment_price.setText(f"${p_val:,.2f}")
            except ValueError:
                self.equipment_price.setText("$0.00")
            
            if self.equipment_manual_stock:
                self.equipment_manual_stock.setFocus()
        elif self.equipment_product_code:
            self.equipment_product_code.clear()
        else:
            self.logger.warning(f"Equipment '{text}' selected but required widgets not available.")

    def on_trade_selected(self, text):
        """Handle trade item selection from autocomplete."""
        self.logger.debug(f"Trade selected via completer: '{text}'")
        self.products_dict = self.data_loader.get_products()
        product_data = self.products_dict.get(text, (None, None))
        code, price_str = product_data
        
        if code is not None and self.trade_amount and self.trade_stock:
            curr_amt = self.trade_amount.text().replace('$', '').replace(',', '')
            curr_stk = self.trade_stock.text()
            try:
                if not curr_amt or float(curr_amt) == 0.0:
                    p_val = float(re.sub(r'[^\d.-]', '', price_str)) if price_str else 0.0
                    self.trade_amount.setText(f"${p_val:,.2f}")
            except ValueError:
                self.trade_amount.setText("$0.00")
                
            if not curr_stk and code:
                self.trade_stock.setText(code)
        else:
            self.logger.warning(f"Trade '{text}' selected but required widgets not available.")

    def on_part_selected(self, text):
        """Handle part name selection from autocomplete."""
        self.logger.debug(f"Part name selected via completer: '{text}'")
        self.parts_dict = self.data_loader.get_parts()
        text = text.strip()
        part_num = self.parts_dict.get(text)
        
        if part_num and self.part_number and not self.part_number.text():
            self.part_number.setText(part_num)

    def on_part_number_selected(self, text):
        """Handle part number selection from autocomplete."""
        self.logger.debug(f"Part number selected via completer: '{text}'")
        self.parts_dict = self.data_loader.get_parts()
        text = text.strip()
        found_name = None
        
        if self.parts_dict:
            for name, num in self.parts_dict.items():
                if num == text:
                    found_name = name
                    break
                    
        if found_name and self.part_name and not self.part_name.text():
            self.part_name.setText(found_name)

    # --- Add/Edit/Delete Item Logic ---
    def add_item(self, item_type):
        """Add an item to the appropriate list."""
        self.logger.debug(f"Attempting to add item of type: {item_type}")
        list_widget = None
        
        try:
            if item_type == "equipment":
                list_widget = self.equipment_list
                if not isinstance(list_widget, QListWidget):
                    raise TypeError("Equipment list widget invalid")
                    
                if not self.equipment_product_name:
                    self.logger.error("equipment_product_name widget not initialized")
                    return
                    
                name = self.equipment_product_name.text().strip()
                code = self.equipment_product_code.text().strip() if self.equipment_product_code else ""
                manual_stock = self.equipment_manual_stock.text().strip() if self.equipment_manual_stock else ""
                price_text = self.equipment_price.text().strip() if self.equipment_price else ""
                
                if not name:
                    QMessageBox.warning(self, "Missing Info", "Please enter or select a Product Name.")
                    return
                    
                if not manual_stock:
                    QMessageBox.warning(self, "Missing Info", "Please enter a manual Stock Number.")
                    return
                    
                price = self._format_currency_text(price_text)
                item_text = f'"{name}" (Code: {code}) STK#{manual_stock} {price}'
                QListWidgetItem(item_text, list_widget)
                
                self.equipment_product_name.clear()
                self.equipment_product_code.clear()
                self.equipment_manual_stock.clear()
                self.equipment_price.clear()
                self.update_charge_to_default()
                self.equipment_product_name.setFocus()

            elif item_type == "trade":
                list_widget = self.trade_list
                if not isinstance(list_widget, QListWidget):
                    raise TypeError("Trade list widget invalid")
                    
                if not self.trade_name:
                    self.logger.error("trade_name widget not initialized")
                    return
                    
                name = self.trade_name.text().strip()
                stock = self.trade_stock.text().strip() if self.trade_stock else ""
                amount_text = self.trade_amount.text().strip() if self.trade_amount else ""
                
                if name:
                    amount = self._format_currency_text(amount_text)
                    item_text = f'"{name}" STK#{stock} {amount}'
                    QListWidgetItem(item_text, list_widget)
                    
                    self.trade_name.clear()
                    self.trade_stock.clear()
                    self.trade_amount.clear()
                    self.trade_name.setFocus()

            elif item_type == "part":
                list_widget = self.part_list
                if not isinstance(list_widget, QListWidget):
                    raise TypeError("Part list widget invalid")
                    
                if not self.part_quantity:
                    self.logger.error("part_quantity widget not initialized")
                    return
                    
                qty = str(self.part_quantity.value())
                number = self.part_number.text().strip() if self.part_number else ""
                name = self.part_name.text().strip() if self.part_name else ""
                location = self.part_location.currentText().strip() if self.part_location else ""
                charge_to = self.part_charge_to.text().strip() if self.part_charge_to else ""
                
                if name or number:
                    item_text = f"{qty}x {number} {name} {location} {charge_to}"
                    QListWidgetItem(item_text, list_widget)
                    
                    if not self.last_charge_to and charge_to:
                        self.last_charge_to = charge_to
                        
                    self.part_name.clear()
                    self.part_number.clear()
                    self.part_quantity.setValue(1)
                    self.part_charge_to.setText(self.last_charge_to)
                    self.part_number.setFocus()
            else:
                self.logger.warning(f"add_item called with unknown type: {item_type}")

        except TypeError as te:
            list_name = item_type.capitalize() + " list"
            self.logger.critical(f"CRITICAL ERROR in add_item ({item_type}): Target list widget is not a QListWidget! Error: {te}")
            QMessageBox.critical(self, "Add Item Error", f"Internal error: Could not find the {list_name} to add the item.")
        except Exception as e:
            self.logger.error(f"Unexpected error in add_item ({item_type}): {e}", exc_info=True)
            QMessageBox.critical(self, "Add Item Error", f"An unexpected error occurred while adding the item: {e}")

    def edit_equipment_item(self, item):
        """Edit an equipment item in the list."""
        if not item: 
            return
            
        current_text = item.text()
        pattern = r'"(.*)"\s+\(Code:\s*(.*?)\)\s+STK#(.*?)\s+\$(.*)'
        match = re.match(pattern, current_text)
        
        if not match:
            QMessageBox.warning(self, "Edit Error", f"Cannot parse equipment: {current_text}")
            return
            
        name, code, manual_stock, price_str = match.groups()
        price = price_str.strip()
        
        new_name, ok = QInputDialog.getText(self, "Edit Equipment", "Name:", text=name)
        if not ok: 
            return
            
        self.products_dict = self.data_loader.get_products()
        new_code_lookup, new_price_str_lookup = self.products_dict.get(new_name, (code, price_str))
        new_code = new_code_lookup if new_code_lookup else code
        
        new_manual_stock, ok = QInputDialog.getText(self, "Edit Equipment", "Stock #:", text=manual_stock)
        if not ok: 
            return
            
        new_price_input, ok = QInputDialog.getText(self, "Edit Equipment", "Price:", text=price.replace(',', ''))
        if not ok: 
            return
            
        new_price_formatted = self._format_currency_text(new_price_input)
        item.setText(f'"{new_name}" (Code: {new_code}) STK#{new_manual_stock} {new_price_formatted}')
        self.update_charge_to_default()

    def edit_trade_item(self, item):
        """Edit a trade item in the list."""
        if not item: 
            return
            
        current_text = item.text()
        pattern = r'"(.*)"\s+STK#(.*?)\s+\$(.*)'
        match = re.match(pattern, current_text)
        
        if not match:
            QMessageBox.warning(self, "Edit Error", f"Cannot parse trade: {current_text}")
            return
            
        name, stock, amount_str = match.groups()
        amount = amount_str.strip()
        
        new_name, ok = QInputDialog.getText(self, "Edit Trade", "Name:", text=name)
        if not ok: 
            return
            
        new_stock, ok = QInputDialog.getText(self, "Edit Trade", "Stock #:", text=stock)
        if not ok: 
            return
            
        new_amount_input, ok = QInputDialog.getText(self, "Edit Trade", "Amount:", text=amount.replace(',', ''))
        if not ok: 
            return
            
        new_amount_formatted = self._format_currency_text(new_amount_input)
        item.setText(f'"{new_name}" STK#{new_stock} {new_amount_formatted}')

    def edit_part_item(self, item):
        """Edit a part item in the list."""
        if not item: 
            return
            
        current_text = item.text()
        parts = current_text.split(" ", 4)
        
        try:
            qty = int(parts[0].rstrip('x'))
            number = parts[1]
            name = parts[2]
            location = parts[3]
            charge_to = parts[4] if len(parts) > 4 else ""
        except (IndexError, ValueError):
            QMessageBox.warning(self, "Edit Error", f"Cannot parse part: {current_text}")
            return
            
        new_qty, ok = QInputDialog.getInt(self, "Edit Part", "Qty:", qty, 1, 999)
        if not ok: 
            return
            
        new_number, ok = QInputDialog.getText(self, "Edit Part", "Part #:", text=number)
        if not ok: 
            return
            
        new_name, ok = QInputDialog.getText(self, "Edit Part", "Name:", text=name)
        if not ok: 
            return
            
        locations = ["", "Camrose", "Killam", "Wainwright", "Provost"]
        current_loc_idx = locations.index(location) if location in locations else 0
        new_location, ok = QInputDialog.getItem(self, "Edit Part", "Location:", locations, current=current_loc_idx, editable=False)
        if not ok: 
            return
            
        new_charge_to, ok = QInputDialog.getText(self, "Edit Part", "Charge To:", text=charge_to)
        if not ok: 
            return
            
        item.setText(f"{new_qty}x {new_number} {new_name} {new_location} {new_charge_to}")

    def delete_selected_list_item(self, list_widget=None, item_type="Item"):
        """Delete a selected item from the specified list."""
        target_list = list_widget
        list_name = item_type
        
        if not isinstance(target_list, QListWidget):
            self.logger.warning(f"Delete called but target list is invalid or not provided.")
            QMessageBox.warning(self, "Delete Error", f"Could not identify the {list_name} list.")
            return
            
        if target_list.currentRow() >= 0:
            item = target_list.currentItem()
            if item:
                reply = QMessageBox.question(self, 'Confirm Delete', 
                                           f"Delete this {list_name} line?\n'{item.text()}'", 
                                           QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                if reply == QMessageBox.Yes:
                    target_list.takeItem(target_list.currentRow())
                    self.logger.info(f"Deleted item from {list_name} list.")
                    
                    if list_widget == self.equipment_list:
                        self.update_charge_to_default()
            else:
                QMessageBox.warning(self, "Delete Error", "Selected item is invalid.")
        else:
            QMessageBox.warning(self, "Delete", f"Select a line in the {list_name} list to delete.")

    def delete_selected_line(self):
        """Delete any selected item from any list."""
        # Check each list for a selection and delete from the first one found
        if self.equipment_list and self.equipment_list.currentRow() >= 0:
            self.delete_selected_list_item(self.equipment_list, "Equipment")
        elif self.trade_list and self.trade_list.currentRow() >= 0:
            self.delete_selected_list_item(self.trade_list, "Trade")
        elif self.part_list and self.part_list.currentRow() >= 0:
            self.delete_selected_list_item(self.part_list, "Part")
        else:
            QMessageBox.information(self, "Delete Line", "Please select an item in one of the lists first.")

    def format_price(self):
        """Format the price field when editing finishes."""
        sender = self.sender()
        sender.setText(self._format_currency_text(sender.text()))

    def format_amount(self):
        """Format the amount field when editing finishes."""
        sender = self.sender()
        sender.setText(self._format_currency_text(sender.text()))

    def _format_currency_text(self, text_in):
        """Format a text string as currency."""
        cleaned = re.sub(r'[^\d.-]', '', text_in)
        value = float(cleaned) if cleaned and cleaned != '-' else 0.0
        return f"${value:,.2f}"

    def update_charge_to_default(self):
        """Update the default charge-to value based on equipment stock number."""
        stock_number = ""
        eq_pattern = r'"(.*)"\s+\(Code:\s*(.*?)\)\s+STK#(.*?)\s+\$(.*)'
        eq_list = self.equipment_list
        
        if isinstance(eq_list, QListWidget) and eq_list.count() > 0:
            item = eq_list.item(0)
            if item:
                item_text = item.text()
                match = re.match(eq_pattern, item_text)
                if match: 
                    stock_number = match.group(3).strip()
            else: 
                self.logger.warning("update_charge_to_default: Item at index 0 is None.")
        else:
            self.logger.debug("update_charge_to_default: Equipment list is empty.")
            
        self.last_charge_to = stock_number
        if self.part_charge_to and not self.part_charge_to.text():
            self.part_charge_to.setText(self.last_charge_to)

    # --- Form Reset ---
    
    
    @pyqtSlot()
    def reset_form(self):
        """Reset the form by asking the main window to reload the module."""
        reply = QMessageBox.question(self, 'Confirm Reset', "Clear all fields and lists?", 
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.logger.info("Resetting Deal Form...")
            
            # Simple clearing of fields and lists
            if isinstance(self.equipment_list, QListWidget): 
                self.equipment_list.clear()
            if isinstance(self.trade_list, QListWidget): 
                self.trade_list.clear()
            if isinstance(self.part_list, QListWidget): 
                self.part_list.clear()
                
            # Clear all form fields
            for w in [self.equipment_product_name, self.equipment_product_code, 
                    self.equipment_manual_stock, self.equipment_price,
                    self.trade_name, self.trade_stock, self.trade_amount,
                    self.part_name, self.part_number, self.part_charge_to,
                    self.work_order_charge_to, self.work_order_hours, self.notes_edit,
                    self.customer_name, self.salesperson]:
                if w: 
                    if hasattr(w, 'clear'):
                        w.clear()
                    elif hasattr(w, 'setText'):
                        w.setText("")
                        
            # Reset widgets to defaults
            if self.part_quantity: 
                self.part_quantity.setValue(1)
            if self.part_location: 
                self.part_location.setCurrentIndex(0)
            if self.work_order_required: 
                self.work_order_required.setChecked(False)
            if self.multi_line_csv: 
                self.multi_line_csv.setChecked(False)
            if self.paid_checkbox: 
                self.paid_checkbox.setChecked(False)

            # Reset internal state
            self.last_charge_to = ""
            self.csv_lines = {}
            
            # Ask the main window to reload the module
            if self.main_window and hasattr(self.main_window, 'reinitialize_module'):
                QTimer.singleShot(100, lambda: self.main_window.reinitialize_module('DealFormModule'))
            elif self.main_window and hasattr(self.main_window, 'reload_module'):
                QTimer.singleShot(100, lambda: self.main_window.reload_module('DealFormModule'))
            elif self.main_window and hasattr(self.main_window, 'tabs'):
                # Try to simulate a tab switch which often reinitializes modules
                current_index = self.main_window.tabs.currentIndex()
                QTimer.singleShot(100, lambda: self.main_window.tabs.setCurrentIndex((current_index + 1) % self.main_window.tabs.count()))
                QTimer.singleShot(200, lambda: self.main_window.tabs.setCurrentIndex(current_index))
            else:
                # If we can't reload the module, show a message
                QMessageBox.information(self, "Form Reset", 
                                    "Basic form reset complete. If autocomplete doesn't work, please switch to another tab and back.")
                
            self._show_status("Form Reset.", 2000)

    def _reinitialize_completers(self):
        """Completely reinitialize all completers from scratch after a reset."""
        self.logger.debug("Reinitializing all completers after reset...")
        
        try:
            # First remove all existing completers
            if hasattr(self, 'customer_completer') and self.customer_completer and self.customer_name:
                self.customer_name.setCompleter(None)
                
            if hasattr(self, 'salesperson_completer') and self.salesperson_completer and self.salesperson:
                self.salesperson.setCompleter(None)
                
            if hasattr(self, 'equipment_completer') and self.equipment_completer and self.equipment_product_name:
                self.equipment_product_name.setCompleter(None)
                
            if hasattr(self, 'trade_completer') and self.trade_completer and self.trade_name:
                self.trade_name.setCompleter(None)
                
            if hasattr(self, 'part_name_completer') and self.part_name_completer and self.part_name:
                self.part_name.setCompleter(None)
                
            if hasattr(self, 'part_number_completer') and self.part_number_completer and self.part_number:
                self.part_number.setCompleter(None)
            
            # Reset all completer variables to None
            self.customer_completer = None
            self.salesperson_completer = None
            self.equipment_completer = None
            self.trade_completer = None
            self.part_name_completer = None
            self.part_number_completer = None
            
            # Force reload all data
            self.data_loader.ensure_all_loaded(force_reload=True)
            
            # Get fresh data
            customer_data = self.data_loader.get_customers(True)
            salesmen_data = self.data_loader.get_salesmen(True)
            product_data = self.data_loader.get_products(True)
            parts_data = self.data_loader.get_parts(True)

            # Create fresh models for completers
            self.customer_model = QStringListModel(customer_data)
            self.salesperson_model = QStringListModel(list(salesmen_data.keys()))
            product_keys = list(product_data.keys())
            self.product_model = QStringListModel(product_keys)
            self.part_name_model = QStringListModel(list(parts_data.keys()))
            self.part_number_model = QStringListModel(list(map(str, parts_data.values())))

            # Setup fresh completers
            if self.customer_name:
                self.customer_completer = QCompleter(self.customer_model, self)
                self.customer_completer.setCaseSensitivity(Qt.CaseInsensitive)
                self.customer_completer.setFilterMode(Qt.MatchContains)
                self.customer_name.setCompleter(self.customer_completer)
                self.logger.debug(f"Customer completer reset with {len(customer_data)} items.")

            if self.salesperson:
                self.salesperson_completer = QCompleter(self.salesperson_model, self)
                self.salesperson_completer.setCaseSensitivity(Qt.CaseInsensitive)
                self.salesperson_completer.setFilterMode(Qt.MatchContains)
                self.salesperson.setCompleter(self.salesperson_completer)
                self.logger.debug(f"Salesperson completer reset with {len(salesmen_data)} items.")

            if self.equipment_product_name:
                self.equipment_completer = QCompleter(self.product_model, self)
                self.equipment_completer.setCaseSensitivity(Qt.CaseInsensitive)
                self.equipment_completer.setFilterMode(Qt.MatchContains)
                self.equipment_product_name.setCompleter(self.equipment_completer)
                self.connect_completer_signals(self.equipment_completer, self.on_equipment_selected)
                self.logger.debug(f"Equipment completer reset with {len(product_keys)} items.")

            if self.trade_name:
                self.trade_completer = QCompleter(self.product_model, self)
                self.trade_completer.setCaseSensitivity(Qt.CaseInsensitive)
                self.trade_completer.setFilterMode(Qt.MatchContains)
                self.trade_name.setCompleter(self.trade_completer)
                self.connect_completer_signals(self.trade_completer, self.on_trade_selected)
                self.logger.debug("Trade name completer reset.")

            if self.part_name:
                self.part_name_completer = QCompleter(self.part_name_model, self)
                self.part_name_completer.setCaseSensitivity(Qt.CaseInsensitive)
                self.part_name_completer.setFilterMode(Qt.MatchContains)
                self.part_name.setCompleter(self.part_name_completer)
                self.connect_completer_signals(self.part_name_completer, self.on_part_selected)
                self.logger.debug(f"Part name completer reset with {len(parts_data)} items.")

            if self.part_number:
                self.part_number_completer = QCompleter(self.part_number_model, self)
                self.part_number_completer.setCaseSensitivity(Qt.CaseInsensitive)
                self.part_number_completer.setFilterMode(Qt.MatchStartsWith)
                self.part_number.setCompleter(self.part_number_completer)
                self.connect_completer_signals(self.part_number_completer, self.on_part_number_selected)
                self.logger.debug("Part number completer reset.")
                
            # Reconnect all internal signals
            self._connect_internal_signals()
            
            self.logger.info("All completers successfully reinitialized after reset.")
        except Exception as e:
            self.logger.error(f"Error reinitializing completers: {e}", exc_info=True)
            QMessageBox.warning(self, "Reset Warning", 
                            "Autocomplete may not work properly after reset. Please close and reopen the form.")

    # --- Data Collection & Draft/Recent Handling ---
    def _get_current_deal_data(self):
        """Collect data from all form fields."""
        self.logger.debug("Collecting data via _get_current_deal_data")
        
        # Get items from lists
        eq_list = self.equipment_list
        tr_list = self.trade_list
        pa_list = self.part_list
        
        equipment_items = [eq_list.item(i).text() for i in range(eq_list.count())] if isinstance(eq_list, QListWidget) else []
        trade_items = [tr_list.item(i).text() for i in range(tr_list.count())] if isinstance(tr_list, QListWidget) else []
        part_items = [pa_list.item(i).text() for i in range(pa_list.count())] if isinstance(pa_list, QListWidget) else []

        # Build data dictionary
        deal_data = {
            "timestamp": datetime.now().isoformat(),
            "customer_name": self.customer_name.text() if self.customer_name else "",
            "salesperson": self.salesperson.text() if self.salesperson else "",
            "equipment": equipment_items,
            "trades": trade_items,
            "parts": part_items,
            "work_order_required": self.work_order_required.isChecked() if self.work_order_required else False,
            "work_order_charge_to": self.work_order_charge_to.text() if self.work_order_charge_to else "",
            "work_order_hours": self.work_order_hours.text() if self.work_order_hours else "",
            "multi_line_csv": self.multi_line_csv.isChecked() if self.multi_line_csv else False,
            "paid": self.paid_checkbox.isChecked() if self.paid_checkbox else False,
            "part_location_index": self.part_location.currentIndex() if self.part_location else 0,
            "last_charge_to": self.last_charge_to,
            "notes": self.notes_edit.toPlainText() if self.notes_edit else ""
        }
        
        return deal_data

    def _collect_form_data(self):
        """Collect data for drafts or state saving."""
        self.logger.debug("Collecting form data for save_state...")
        
        # Start with current data
        data = self.current_form_data.copy()
        
        # Get list items
        eq_list = self.equipment_list
        tr_list = self.trade_list
        pa_list = self.part_list
        
        try: 
            data['equipment'] = [eq_list.item(i).text() for i in range(eq_list.count())] if isinstance(eq_list, QListWidget) else []
        except Exception as e: 
            self.logger.error(f"Error collecting equipment list state: {e}")
            data['equipment'] = ['ERROR']
            
        try: 
            data['trades'] = [tr_list.item(i).text() for i in range(tr_list.count())] if isinstance(tr_list, QListWidget) else []
        except Exception as e: 
            self.logger.error(f"Error collecting trade list state: {e}")
            data['trades'] = ['ERROR']
            
        try: 
            data['parts'] = [pa_list.item(i).text() for i in range(pa_list.count())] if isinstance(pa_list, QListWidget) else []
        except Exception as e: 
            self.logger.error(f"Error collecting part list state: {e}")
            data['parts'] = ['ERROR']
        
        # Get field values
        if self.customer_name: 
            data['customer_name'] = self.customer_name.text()
        if self.salesperson: 
            data['salesperson'] = self.salesperson.text()
        if self.work_order_required: 
            data['work_order_required'] = self.work_order_required.isChecked()
        if self.work_order_charge_to: 
            data['work_order_charge_to'] = self.work_order_charge_to.text()
        if self.work_order_hours: 
            data['work_order_hours'] = self.work_order_hours.text()
        if self.multi_line_csv: 
            data['multi_line_csv'] = self.multi_line_csv.isChecked()
        if self.paid_checkbox: 
            data['paid'] = self.paid_checkbox.isChecked()
        if self.part_location: 
            data['part_location_index'] = self.part_location.currentIndex()
        if self.notes_edit: 
            data['notes'] = self.notes_edit.toPlainText()
            
    # Add timestamp
        data['timestamp'] = datetime.now().isoformat()
        self.logger.debug(f"Form data collected: {len(data)} fields")
        return data

    def save_state(self):
        """Save the current form state for persistence."""
        try:
            form_data = self._collect_form_data()
            self.logger.debug("Form state saved.")
            return form_data
        except Exception as e:
            self.logger.error(f"Error saving form state: {e}", exc_info=True)
            return {}

    @pyqtSlot()
    def save_draft(self):
        """Save the current form state as a draft."""
        if not self.cache_path:
            QMessageBox.critical(self, "Error", "Cache directory not set!")
            return
            
        try:
            self.logger.info("Saving draft...")
            draft_path = os.path.join(self.cache_path, DRAFT_FILENAME)
            
            form_data = self._get_current_deal_data()
            if not form_data:
                self.logger.warning("No form data to save")
                QMessageBox.warning(self, "Save Draft", "No data to save.")
                return
                
            # Convert to JSON and save to file
            with open(draft_path, 'w', encoding='utf-8') as f:
                json.dump(form_data, f, indent=2, ensure_ascii=False)
                
            # Update recent deals
            self._update_recent_deals(form_data)
            
            self.logger.info(f"Draft saved to {draft_path}")
            QMessageBox.information(self, "Save Draft", "Form draft saved successfully.")
            
        except Exception as e:
            self.logger.error(f"Error saving draft: {e}", exc_info=True)
            QMessageBox.critical(self, "Save Error", f"Error saving draft: {str(e)}")

    def _update_recent_deals(self, current_deal):
        """Update the list of recent deals."""
        if not self.cache_path:
            self.logger.warning("Cache path not set, can't save recent deals")
            return
            
        try:
            # Load existing recent deals
            recent_deals_path = os.path.join(self.cache_path, RECENT_DEALS_FILENAME)
            recent_deals = []
            
            if os.path.exists(recent_deals_path):
                try:
                    with open(recent_deals_path, 'r', encoding='utf-8') as f:
                        recent_deals = json.load(f)
                except json.JSONDecodeError:
                    self.logger.warning(f"Recent deals file corrupt, creating new file")
                except Exception as e:
                    self.logger.warning(f"Error loading recent deals: {e}")
            
            # Add current deal to the list
            current_deal_summary = {
                "timestamp": current_deal.get("timestamp", datetime.now().isoformat()),
                "customer_name": current_deal.get("customer_name", ""),
                "equipment": current_deal.get("equipment", [])[:1],  # Just the first piece of equipment
                "created": datetime.now().isoformat()
            }
            
            # Insert at the beginning
            recent_deals.insert(0, current_deal_summary)
            
            # Keep only the most recent MAX_RECENT_DEALS
            recent_deals = recent_deals[:MAX_RECENT_DEALS]
            
            # Save the updated list
            with open(recent_deals_path, 'w', encoding='utf-8') as f:
                json.dump(recent_deals, f, indent=2, ensure_ascii=False)
                
            self.logger.debug(f"Recent deals updated, now have {len(recent_deals)} entries")
            
        except Exception as e:
            self.logger.error(f"Error updating recent deals: {e}", exc_info=True)

    def _attempt_load_draft(self):
        """Attempt to load a saved draft."""
        if not self.cache_path:
            self.logger.warning("Cache path not set, can't load draft")
            return
            
        try:
            draft_path = os.path.join(self.cache_path, DRAFT_FILENAME)
            
            if not os.path.exists(draft_path):
                self.logger.debug("No draft file found")
                return
                
            with open(draft_path, 'r', encoding='utf-8') as f:
                draft_data = json.load(f)
                
            self.logger.info(f"Draft loaded from {draft_path}")
            
            # Only show the dialog when manually clicked, not on startup
            is_manual = self.sender() == self.load_draft_btn
            
            if is_manual:
                msg = f"Load draft for customer: {draft_data.get('customer_name', 'Unknown')}?"
                reply = QMessageBox.question(self, 'Load Draft', msg, 
                                            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                                            
                if reply == QMessageBox.No:
                    return
                    
            self.populate_form(draft_data)
            
            if is_manual:
                QMessageBox.information(self, "Load Draft", "Draft loaded successfully.")
                
        except json.JSONDecodeError:
            self.logger.error("Draft file corrupt")
            if self.sender() == self.load_draft_btn:  # Only show error if manually clicked
                QMessageBox.critical(self, "Load Error", "Draft file is corrupt.")
        except Exception as e:
            self.logger.error(f"Error loading draft: {e}", exc_info=True)
            if self.sender() == self.load_draft_btn:  # Only show error if manually clicked
                QMessageBox.critical(self, "Load Error", f"Error loading draft: {str(e)}")

    def populate_form(self, data):
        """Populate the form with the provided data."""
        self.logger.debug(f"Populating form with data: {data.get('customer_name', 'Unknown')}")
        
        # Clear all lists first
        for list_widget in [self.equipment_list, self.trade_list, self.part_list]:
            if isinstance(list_widget, QListWidget):
                list_widget.clear()
                
        # Populate basic fields
        if self.customer_name and 'customer_name' in data:
            self.customer_name.setText(data.get('customer_name', ''))
            
        if self.salesperson and 'salesperson' in data:
            self.salesperson.setText(data.get('salesperson', ''))
            
        if self.work_order_required and 'work_order_required' in data:
            self.work_order_required.setChecked(data.get('work_order_required', False))
            
        if self.work_order_charge_to and 'work_order_charge_to' in data:
            self.work_order_charge_to.setText(data.get('work_order_charge_to', ''))
            
        if self.work_order_hours and 'work_order_hours' in data:
            self.work_order_hours.setText(data.get('work_order_hours', ''))
            
        if self.multi_line_csv and 'multi_line_csv' in data:
            self.multi_line_csv.setChecked(data.get('multi_line_csv', False))
            
        if self.paid_checkbox and 'paid' in data:
            self.paid_checkbox.setChecked(data.get('paid', False))
            
        if self.part_location and 'part_location_index' in data:
            self.part_location.setCurrentIndex(data.get('part_location_index', 0))
            
        if self.notes_edit and 'notes' in data:
            self.notes_edit.setPlainText(data.get('notes', ''))
            
        # Update last charge to
        self.last_charge_to = data.get('last_charge_to', '')
        
        # Populate list widgets
        if isinstance(self.equipment_list, QListWidget):
            for item in data.get('equipment', []):
                QListWidgetItem(item, self.equipment_list)
                
        if isinstance(self.trade_list, QListWidget):
            for item in data.get('trades', []):
                QListWidgetItem(item, self.trade_list)
                
        if isinstance(self.part_list, QListWidget):
            for item in data.get('parts', []):
                QListWidgetItem(item, self.part_list)
                
        # Update default charge to
        self.update_charge_to_default()
        
        self.logger.debug("Form populated from data")

    # --- Output Generation ---
    def generate_csv_lines(self):
        """Generate CSV lines from form data."""
        self.logger.debug("Generating CSV lines...")
        csv_lines = []
        
        # Check for customer name
        customer = self.customer_name.text().strip() if self.customer_name else ""
        salesperson = self.salesperson.text().strip() if self.salesperson else ""
        
        if not customer:
            QMessageBox.warning(self, "Missing Information", "Please enter a customer name.")
            return None
            
        if not salesperson:
            QMessageBox.warning(self, "Missing Information", "Please enter a salesperson name.")
            return None
            
        # Check for equipment
        eq_list = self.equipment_list
        if not isinstance(eq_list, QListWidget) or eq_list.count() == 0:
            QMessageBox.warning(self, "Missing Information", "Please add at least one piece of equipment.")
            return None
            
        # Get current date for the csv
        date_str = datetime.now().strftime("%m/%d/%Y")
        multi_line = self.multi_line_csv.isChecked() if self.multi_line_csv else False
        
        # Process equipment
        equipment_items = []
        total_equipment_price = 0.0
        
        for i in range(eq_list.count()):
            item = eq_list.item(i).text()
            pattern = r'"(.*)"\s+\(Code:\s*(.*?)\)\s+STK#(.*?)\s+\$(.*)'
            match = re.match(pattern, item)
            
            if match:
                name, code, stock, price_str = match.groups()
                try:
                    price = float(price_str.replace(',', '').strip())
                    total_equipment_price += price
                    equipment_items.append((name, code, stock, price))
                except ValueError:
                    self.logger.warning(f"Could not parse price from equipment: {item}")
                    equipment_items.append((name, code, stock, 0.0))
            else:
                self.logger.warning(f"Could not parse equipment item: {item}")
                
        # Process trades
        trade_items = []
        total_trade_amount = 0.0
        
        tr_list = self.trade_list
        if isinstance(tr_list, QListWidget):
            for i in range(tr_list.count()):
                item = tr_list.item(i).text()
                pattern = r'"(.*)"\s+STK#(.*?)\s+\$(.*)'
                match = re.match(pattern, item)
                
                if match:
                    name, stock, amount_str = match.groups()
                    try:
                        amount = float(amount_str.replace(',', '').strip())
                        total_trade_amount += amount
                        trade_items.append((name, stock, amount))
                    except ValueError:
                        self.logger.warning(f"Could not parse amount from trade: {item}")
                        trade_items.append((name, stock, 0.0))
                else:
                    self.logger.warning(f"Could not parse trade item: {item}")
                    
        # Process parts
        part_items = []
        
        pa_list = self.part_list
        if isinstance(pa_list, QListWidget):
            for i in range(pa_list.count()):
                item = pa_list.item(i).text()
                parts = item.split(" ", 4)
                
                try:
                    qty = int(parts[0].rstrip('x'))
                    number = parts[1]
                    name = parts[2]
                    location = parts[3]
                    charge_to = parts[4] if len(parts) > 4 else ""
                    part_items.append((qty, number, name, location, charge_to))
                except (IndexError, ValueError):
                    self.logger.warning(f"Could not parse part item: {item}")
                    
        # Process work order
        work_order_data = None
        if self.work_order_required and self.work_order_required.isChecked():
            charge_to = self.work_order_charge_to.text().strip() if self.work_order_charge_to else ""
            hours = self.work_order_hours.text().strip() if self.work_order_hours else ""
            
            if charge_to or hours:
                work_order_data = (charge_to, hours)
                
        # Format balance
        balance = total_equipment_price - total_trade_amount
        balance_str = f"${balance:,.2f}"
        paid = self.paid_checkbox.isChecked() if self.paid_checkbox else False
        
        # Create CSV header
        notes = self.notes_edit.toPlainText() if self.notes_edit else ""
        
        # Generate CSV lines
        if multi_line:
            # Create one line per equipment item
            for eq_idx, (eq_name, eq_code, eq_stock, eq_price) in enumerate(equipment_items):
                # Only include trades on the first line
                line_trade_str = ""
                if eq_idx == 0 and trade_items:
                    trade_names = [f"{t[0]} (${t[2]:,.2f})" for t in trade_items]
                    line_trade_str = " | ".join(trade_names)
                    
                # Format the line
                line = [
                    date_str,
                    customer,
                    salesperson,
                    eq_name,
                    eq_stock,
                    f"${eq_price:,.2f}",
                    line_trade_str,
                    balance_str if eq_idx == 0 else "",
                    "PAID" if paid and eq_idx == 0 else "",
                    notes if eq_idx == 0 else ""
                ]
                csv_lines.append(line)
        else:
            # Basic single line
            equipment_names = [f"{eq[0]} (${eq[3]:,.2f})" for eq in equipment_items]
            equipment_str = " | ".join(equipment_names)
            
            trade_str = ""
            if trade_items:
                trade_names = [f"{t[0]} (${t[2]:,.2f})" for t in trade_items]
                trade_str = " | ".join(trade_names)
                
            line = [
                date_str,
                customer,
                salesperson,
                equipment_str,
                equipment_items[0][2] if equipment_items else "",  # Stock number from first equipment
                f"${total_equipment_price:,.2f}",
                trade_str,
                balance_str,
                "PAID" if paid else "",
                notes
            ]
            csv_lines.append(line)
            
        # Add parts info
        for part in part_items:
            qty, number, name, location, charge_to = part
            part_line = [
                date_str,
                customer,
                salesperson,
                f"PART: {qty}x {name}",
                number,
                "",
                location,
                charge_to,
                "",
                ""
            ]
            csv_lines.append(part_line)
            
        # Add work order info if needed
        if work_order_data:
            charge_to, hours = work_order_data
            wo_line = [
                date_str,
                customer,
                salesperson,
                "WORK ORDER",
                "",
                "",
                "",
                charge_to,
                hours,
                ""
            ]
            csv_lines.append(wo_line)
            
        self.logger.info(f"Generated {len(csv_lines)} CSV lines")
        return csv_lines

    def save_to_sharepoint(self):
        """Save CSV data to SharePoint Excel spreadsheet with the correct column mapping and timeout handling."""
        self.logger.debug("Attempting to save to SharePoint...")
        
        if not self.sharepoint_manager:
            QMessageBox.critical(self, "Error", "SharePoint manager not initialized.")
            return
            
        # Generate CSV lines
        csv_lines = self.generate_csv_lines()
        if not csv_lines:
            return  # Error messages already shown
            
        # Store for use in generate_email
        self.csv_lines = csv_lines
        
        # Get customer name for the spreadsheet
        customer_name = self.customer_name.text().strip() if self.customer_name else "Unknown"
        salesperson_name = self.salesperson.text().strip() if self.salesperson else ""
        
        # Get payment status
        paid = self.paid_checkbox.isChecked() if self.paid_checkbox else False
        payment_status = "🟩" if paid else "🟥"
        
        # Get today's date for the timestamp
        current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        email_date = datetime.now().strftime("%Y-%m-%d")
        
        # Show progress message
        self.logger.info(f"Starting SharePoint update for customer: {customer_name}")
        msgbox = QMessageBox(self)
        msgbox.setIcon(QMessageBox.Information)
        msgbox.setText("Updating SharePoint spreadsheet...")
        msgbox.setStandardButtons(QMessageBox.NoButton)
        msgbox.show()
        QApplication.processEvents()
        
        try:
            # Use the update_excel_data method which is available in your SharePoint manager
            if hasattr(self.sharepoint_manager, 'update_excel_data'):
                # Define the target column headers
                target_headers = [
                    "Payment", "CustomerName", "Equipment", "Stock Number", "Amount", 
                    "Trade", "Attached to stk#", "Trade STK#", "Amount2", 
                    "Salesperson", "Email Date", "Status", "Timestamp"
                ]
                
                # Create spreadsheet data in the correct format
                spreadsheet_data = []
                
                # Process each CSV line to extract the relevant information
                for line in csv_lines:
                    # Extract data from CSV line
                    if len(line) >= 8:  # Make sure we have enough elements
                        # Basic fields from CSV
                        date = line[0] if len(line) > 0 else ""
                        customer = line[1] if len(line) > 1 else customer_name
                        salesperson = line[2] if len(line) > 2 else salesperson_name
                        equipment = line[3] if len(line) > 3 else ""
                        stock_number = line[4] if len(line) > 4 else ""
                        amount = line[5] if len(line) > 5 else ""
                        trade = line[6] if len(line) > 6 else ""
                        balance = line[7] if len(line) > 7 else ""
                        
                        # Create a dictionary with the correct target columns
                        row_dict = {
                            "Payment": payment_status,
                            "CustomerName": customer,
                            "Equipment": equipment,
                            "Stock Number": stock_number,
                            "Amount": amount,
                            "Trade": trade,
                            "Attached to stk#": stock_number,  # Same as stock number
                            "Trade STK#": "",  # Extract from trade info if available
                            "Amount2": "",  # Extract from trade info if available
                            "Salesperson": salesperson,
                            "Email Date": email_date,
                            "Status": "New",
                            "Timestamp": current_date
                        }
                        
                        # Try to extract trade stock number and amount if available
                        if trade:
                            trade_match = re.search(r'STK#(.*?)\s+\$([\d,.]+)', trade)
                            if trade_match:
                                row_dict["Trade STK#"] = trade_match.group(1).strip()
                                row_dict["Amount2"] = trade_match.group(2).strip()
                        
                        spreadsheet_data.append(row_dict)
                
                # Set a timeout timer to prevent freezes
                timeout_timer = QTimer()
                timeout_timer.setSingleShot(True)
                timeout_timer.timeout.connect(lambda: self._on_sharepoint_upload_timeout(msgbox))
                timeout_timer.start(30000)  # 30 second timeout
                
                # Call the update method directly without Worker for simplicity
                try:
                    result = self.sharepoint_manager.update_excel_data(spreadsheet_data, "App")
                    # If we get here, the call succeeded
                    timeout_timer.stop()
                    msgbox.hide()
                    self._on_sharepoint_upload_success(result)
                except Exception as inner_e:
                    # Handle exceptions from the direct call
                    timeout_timer.stop()
                    msgbox.hide()
                    self.logger.error(f"Error during SharePoint update: {inner_e}", exc_info=True)
                    QMessageBox.critical(self, "Upload Error", f"Error: {str(inner_e)}")
                    
            else:
                raise AttributeError(f"SharePoint manager does not have the 'update_excel_data' method.")
            
        except Exception as e:
            msgbox.hide()
            self.logger.error(f"Error updating SharePoint: {e}", exc_info=True)
            QMessageBox.critical(self, "Upload Error", f"Error: {str(e)}")

    def _on_sharepoint_upload_timeout(self, msgbox):
        """Handle timeout during SharePoint upload."""
        if msgbox and msgbox.isVisible():
            msgbox.hide()
        self.logger.error("SharePoint upload timed out after 30 seconds")
        QMessageBox.warning(self, "Upload Timeout", 
                        "The SharePoint upload is taking too long and may have stalled.\n\n"
                        "The application will continue to function. You can try again later.")

    def _on_sharepoint_upload_success(self, result):
        """Handle successful SharePoint upload."""
        self.logger.info(f"SharePoint upload successful: {result}")
        QMessageBox.information(self, "Upload Success", 
                            f"Deal information uploaded to SharePoint successfully.")
        
        # Signal that a deal was saved
        deal_data = self._get_current_deal_data()
        self.deal_saved.emit(deal_data)

    def _on_sharepoint_upload_error(self, error_tuple):
        """Handle SharePoint upload error."""
        error, traceback_str = error_tuple
        self.logger.error(f"SharePoint upload error: {error}\n{traceback_str}")
        
        err_msg = str(error)
        if "credential" in err_msg.lower() or "authenticat" in err_msg.lower():
            err_msg = "Authentication error: Your SharePoint credentials may have expired." \
                    "\n\nPlease try restarting the application."
                    
        QMessageBox.critical(self, "Upload Error", f"Error uploading to SharePoint:\n\n{err_msg}")

    def generate_email(self):
        """Generate an email for the current deal with improved formatting."""
        self.logger.debug("Generating email...")
        
        # Generate CSV lines if needed
        if not self.csv_lines:
            self.csv_lines = self.generate_csv_lines()
            if not self.csv_lines:
                return  # Error messages already shown
                
        # Get salespeople data for email addresses
        salespeople = self.data_loader.get_salesmen()
        salesperson_name = self.salesperson.text().strip() if self.salesperson else ""
        
        # Get basic deal info
        customer_name = self.customer_name.text().strip() if self.customer_name else "Unknown"
        date_str = datetime.now().strftime("%B %d, %Y")
        
        # Get first product name for subject line
        first_product_name = ""
        if isinstance(self.equipment_list, QListWidget) and self.equipment_list.count() > 0:
            item_text = self.equipment_list.item(0).text()
            eq_pattern = r'"(.*)"\s+\(Code:\s*(.*?)\)\s+STK#(.*?)\s+\$(.*)'
            match = re.match(eq_pattern, item_text)
            if match:
                first_product_name = match.group(1)
        
        # Build email subject
        subject = f"AMS Deal ({customer_name} - {first_product_name})"
        
        # Build email body with improved formatting
        body = f"CUSTOMER: {customer_name}\n"
        body += f"SALESPERSON: {salesperson_name}\n"
        body += f"DATE: {date_str}\n\n"
        
        # Add equipment details with formatting
        body += "EQUIPMENT DETAILS"
        body += "\n" + "-" * 65 + "\n"
        
        # Get equipment items from the actual equipment list for more accurate data
        eq_items = []
        if isinstance(self.equipment_list, QListWidget):
            eq_pattern = r'"(.*)"\s+\(Code:\s*(.*?)\)\s+STK#(.*?)\s+\$(.*)'
            for i in range(self.equipment_list.count()):
                item_text = self.equipment_list.item(i).text()
                match = re.match(eq_pattern, item_text)
                if match:
                    name, code, stock, price = match.groups()
                    # Format each item on its own line with simple format
                    eq_entry = f"{name} STK#{stock} ${price.strip()}"
                    eq_items.append(eq_entry)
        
        # If no items in equipment list, try extracting from CSV lines as fallback
        if not eq_items:
            for line in self.csv_lines:
                if len(line) >= 6 and not line[3].startswith("PART:") and not line[3].startswith("WORK ORDER"):
                    eq_name = line[3]
                    eq_stock = line[4] if len(line) > 4 else ""
                    eq_price = line[5]
                    if eq_name:
                        eq_entry = f"{eq_name} STK#{eq_stock} {eq_price}"
                        eq_items.append(eq_entry)
        
        if eq_items:
            body += "\n".join(eq_items)
        else:
            body += "No equipment items found."
        
        body += "\n" + "-" * 65 + "\n"
        
        # Add trade details
        trade_items = []
        for line in self.csv_lines:
            if len(line) >= 7 and line[6] and not line[3].startswith("PART:") and not line[3].startswith("WORK ORDER"):
                trade_str = line[6]
                if trade_str:
                    trade_items = [t.strip() for t in trade_str.split("|")]
                    break
        
        if trade_items:
            body += "\nTRADE DETAILS"
            body += "\n" + "-" * 65 + "\n"
            for item in trade_items:
                body += f"{item}\n"
            body += "-" * 65 + "\n"
        
        # Add parts section if needed
        part_items = []
        for line in self.csv_lines:
            if len(line) >= 4 and line[3].startswith("PART:"):
                part_details = f"{line[3].replace('PART: ', '')}"
                if len(line) >= 5 and line[4]:  # Part number
                    part_details += f" • Part #: {line[4]}"
                if len(line) >= 7 and line[6]:  # Location
                    part_details += f" • Location: {line[6]}"
                if len(line) >= 8 and line[7]:  # Charge to
                    part_details += f" • Charge to: {line[7]}"
                part_items.append(part_details)
        
        if part_items:
            body += "\nPARTS DETAILS"
            body += "\n" + "-" * 65 + "\n"
            body += "\n".join(part_items)
            body += "\n" + "-" * 65 + "\n"
        
        # Add work order section if needed
        work_order_found = False
        work_order_details = []
        
        for line in self.csv_lines:
            if len(line) >= 4 and line[3] == "WORK ORDER":
                work_order_found = True
                wo_entry = "WORK ORDER"
                if len(line) >= 8 and line[7]:  # Charge to
                    wo_entry += f" • Charge to: {line[7]}"
                if len(line) >= 9 and line[8]:  # Hours
                    wo_entry += f" • Hours: {line[8]}"
                work_order_details.append(wo_entry)
        
        if work_order_found:
            body += "\nWORK ORDER DETAILS"
            body += "\n" + "-" * 65 + "\n"
            body += "\n".join(work_order_details)
            body += "\n" + "-" * 65 + "\n"
        
        # Add balance and paid status
        body += "\nDEAL STATUS"
        body += "\n" + "-" * 65 + "\n"
        
        # Add status checklist
        body += "✓ PFW updated\n"
        body += "✓ Spreadsheet updated\n"
        
        # Add balance
        balance = ""
        for line in self.csv_lines:
            if len(line) >= 8 and line[7] and not line[7].startswith("PAID"):
                balance = line[7]
                break
        
        # Add paid status
        paid = False
        for line in self.csv_lines:
            if len(line) >= 9 and line[8] == "PAID":
                paid = True
                break
        
        if paid:
            body += "✓ Payment received\n"
        else:
            body += f"✓ {salesperson_name} to collect\n"
        
        body += "=" * 65 + "\n"
        
        # Add notes if present
        notes = ""
        for line in self.csv_lines:
            if len(line) >= 10 and line[9]:
                notes = line[9]
                break
        
        if notes:
            body += "\nNOTES\n"
            body += "-" * 65 + "\n"
            body += f"{notes}\n"
            body += "=" * 65 + "\n"
        
        # --- Recipients ---
        # --- Recipients ---
        fixed_recipients = [
            'bstdenis@briltd.com', 'cgoodrich@briltd.com', 'dvriend@briltd.com',
            'rbendfeld@briltd.com', 'bfreadrich@briltd.com', 'vedwards@briltd.com'
        ]
        
        # Use .get on the dictionary which handles None if loading failed
        salesman_email = salespeople.get(salesperson_name)
        if salesman_email:
            fixed_recipients.append(salesman_email)
        else:
            print(f"Warning: Email for salesperson '{salesperson_name}' not found.")
            QMessageBox.warning(self, "Salesperson Email", f"Email for {salesperson_name} not found.")
        
        # Add parts department if parts are included
        if part_items:
            fixed_recipients.append('rkrys@briltd.com')
        
        # Filter out invalid emails
        recipient_list = [r for r in fixed_recipients if r and '@' in r]
        recipients_string = ";".join(recipient_list)
        
        # Create mailto URL with all recipients
        mail_to = f"mailto:{recipients_string}"
        mail_to += f"?subject={quote(subject)}&body={quote(body)}"
        
        # Open default mail client
        self.logger.info(f"Opening email client with mailto URL")
        webbrowser.open(mail_to)
        
        QMessageBox.information(self, "Email Generated", 
                            "An email has been generated and opened in your default email client.")

    def generate_all(self):
        """Generate both CSV for SharePoint and email."""
        self.logger.debug("Generating all outputs...")
        
        # First save to SharePoint
        self.save_to_sharepoint()
        
        # Then generate email
        self.generate_email()

    # --- UI Style ---
    def apply_styles(self):
        """Apply CSS styles to form elements."""
        self.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                font-size: 14px;
                border: 1px solid #cccccc;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 12px;
            }
            
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 10px;
                padding: 0 5px;
            }
            
            QLabel {
                font-size: 12px;
            }
            
            QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {
                padding: 5px;
                border: 1px solid #aaa;
                border-radius: 3px;
                background-color: white;
                height: 25px;
            }
            
            QTextEdit, QPlainTextEdit, QListWidget {
                border: 1px solid #aaa;
                border-radius: 3px;
                background-color: white;
            }
            
            QPushButton {
                background-color: #4a9c3e;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 15px;
                font-weight: bold;
            }
            
            QPushButton:hover {
                background-color: #367C2B;
            }
            
            QPushButton:pressed {
                background-color: #2a5d24;
            }
            
            #reset_btn {
                background-color: #d9534f;
            }
            
            #reset_btn:hover {
                background-color: #c9302c;
            }
            
            #delete_line_btn, #remove_part_btn, #remove_trade_btn {
                background-color: #f0ad4e;
                color: #333;
            }
            
            #delete_line_btn:hover, #remove_part_btn:hover, #remove_trade_btn:hover {
                background-color: #ec971f;
            }
        """)
        
        # Apply specific font to header
        if hasattr(self, 'header_label'):
            font = QFont("Arial", 16, QFont.Bold)
            self.header_label.setFont(font)

    def _show_status(self, message, timeout=3000):
        """Show a status message in the main window status bar."""
        if self.main_window:
            try:
                # Try as a method first
                if callable(self.main_window.statusBar):
                    self.main_window.statusBar().showMessage(message, timeout)
                # Then try as a property
                elif hasattr(self.main_window, 'statusBar'):
                    self.main_window.statusBar.showMessage(message, timeout)
                # Fallback to logging
                else:
                    self.logger.info(f"Status message: {message}")
            except Exception as e:
                self.logger.warning(f"Could not show status message: {e}")

    def get_title(self):
        """Get the module display title."""
        return self.MODULE_DISPLAY_NAME

    def get_icon_name(self):
        """Get the module icon filename."""
        return self.MODULE_ICON_NAME

    def refresh(self):
        """Refresh completers and data."""
        self.logger.debug("Refreshing module data...")
        self._update_completers(force_reload=True)

    def close(self):
        """Perform cleanup when module is closed."""
        self.logger.debug("Closing module...")
        self.save_draft()
        super().close()
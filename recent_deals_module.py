from PyQt5.QtWidgets import (QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
                           QTableWidget, QTableWidgetItem, QHeaderView)
from PyQt5.QtCore import Qt
from ui.base_module import BaseModule
import os
import json

class RecentDealsModule(BaseModule):
    def __init__(self, main_window, data_path=None):
        self.data_path = data_path
        super().__init__(main_window)
    
    def init_ui(self):
        """Initialize the user interface."""
        # Main layout
        layout = QVBoxLayout()
        
        # Title
        title = QLabel("Recent Deals")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 24pt; font-weight: bold;")
        layout.addWidget(title)
        
        # Table for recent deals
        self.deals_table = QTableWidget()
        self.deals_table.setColumnCount(6)
        self.deals_table.setHorizontalHeaderLabels([
            "Deal ID", "Customer", "Equipment", "Date", "Status", "Total"
        ])
        self.deals_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.deals_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.deals_table.doubleClicked.connect(self.open_deal)
        
        layout.addWidget(self.deals_table)
        
        # Load data
        self.load_deals()
        
        # Set layout
        self.setLayout(layout)
    
    def load_deals(self):
        """Load recent deals from data file."""
        if not self.data_path:
            return
            
        json_file = os.path.join(self.data_path, "recent_deals.json")
        csv_file = os.path.join(self.data_path, "recent_deals.csv")
        
        # Try loading from JSON first
        if os.path.exists(json_file):
            try:
                with open(json_file, 'r') as f:
                    deals = json.load(f)
                self.populate_table(deals)
                return
            except Exception as e:
                if hasattr(self, 'logger'):
                    self.logger.error(f"Error loading deals from JSON: {str(e)}")
        
        # Fall back to CSV if available
        if os.path.exists(csv_file):
            try:
                import csv
                deals = []
                with open(csv_file, 'r', newline='') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        deals.append(row)
                self.populate_table(deals)
            except Exception as e:
                if hasattr(self, 'logger'):
                    self.logger.error(f"Error loading deals from CSV: {str(e)}")
    
    def populate_table(self, deals):
        """Populate the table with deal data.
        
        Args:
            deals: List of deal dictionaries
        """
        self.deals_table.setRowCount(len(deals))
        
        for row, deal in enumerate(deals):
            # Deal ID
            self.deals_table.setItem(row, 0, QTableWidgetItem(str(deal.get("id", ""))))
            
            # Customer
            self.deals_table.setItem(row, 1, QTableWidgetItem(deal.get("customer", "")))
            
            # Equipment
            self.deals_table.setItem(row, 2, QTableWidgetItem(deal.get("equipment", "")))
            
            # Date
            self.deals_table.setItem(row, 3, QTableWidgetItem(deal.get("date", "")))
            
            # Status
            self.deals_table.setItem(row, 4, QTableWidgetItem(deal.get("status", "")))
            
            # Total
            self.deals_table.setItem(row, 5, QTableWidgetItem(deal.get("total", "")))
    
    def open_deal(self, index):
        """Open the selected deal."""
        row = index.row()
        deal_id = self.deals_table.item(row, 0).text()
        
        # TODO: Implement deal opening functionality
        if hasattr(self.main_window, 'statusBar'):
            self.main_window.statusBar().showMessage(f"Opening deal {deal_id}...")
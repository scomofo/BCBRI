from PyQt5.QtWidgets import (QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
                            QTableWidget, QTableWidgetItem, QComboBox,
                            QDateEdit, QMessageBox, QDialog, QFormLayout,
                            QLineEdit, QTextEdit, QDialogButtonBox, QHeaderView)
from PyQt5.QtCore import Qt, QDate
import logging
import traceback
import requests
import time
import json
from datetime import datetime, timedelta

# Try to import as much as possible, with fallbacks
try:
    from ui.base_module import BaseModule
except ImportError:
    # Create a fallback BaseModule if the original can't be imported
    from PyQt5.QtWidgets import QWidget
    class BaseModule(QWidget):
        def __init__(self, main_window):
            super().__init__()
            self.main_window = main_window
            self.init_ui()
        
        def init_ui(self):
            pass
        
        def refresh(self):
            pass

try:
    from api.QuoteBuilder import QuoteBuilder
except ImportError:
    # Create a minimal QuoteBuilder if the original can't be imported
    class QuoteBuilder:
        @staticmethod
        def create_customer_data(first_name, last_name, email=None, phone=None, **kwargs):
            customer = {
                "customerFirstName": first_name,
                "customerLastName": last_name,
            }
            if email:
                customer["customerEmail"] = email
            if phone:
                customer["customerPhone"] = phone
            return customer
            
        @staticmethod
        def create_basic_quote(dealer_racf_id, customer_data, quote_name, expiration_date=None):
            from datetime import datetime
            quote = {
                "dealerRacfId": dealer_racf_id,
                "quoteName": quote_name,
                "customerData": customer_data,
                "quoteStatusId": 1,
                "quoteType": 2,
                "equipmentData": [],
                "creationDate": datetime.now().strftime("%m/%d/%Y")
            }
            if expiration_date:
                quote["expirationDate"] = expiration_date
            return quote

class JDQuotesModule(BaseModule):
    """Module for interacting with John Deere Quotes."""
    
    def __init__(self, main_window, logger=None, quote_integration=None):
        """Initialize the JD Quotes module.
        
        Args:
            main_window: Reference to the main application window.
            logger: Logger instance.
            quote_integration: Instance of QuoteIntegration.
        """
        # Initialize variables before calling super().__init__
        self.main_window = main_window
        self.logger = logger or logging.getLogger(__name__)
        self.quote_integration = quote_integration
        
        # Check if quote_integration was actually passed
        if self.quote_integration is None:
            self.logger.error("QuoteIntegration instance was not provided to JDQuotesModule!")
            # Try to get it from main_window as fallback
            self.quote_integration = getattr(self.main_window, 'quote_integration', None)
            if self.quote_integration is None:
                self.logger.error("QuoteIntegration not available on main_window either!")
        
        # Use fixed dealer ID instead of a dropdown
        self.dealer_racf_id = "731804"  # Fixed dealer ID
        self.dealer_account_no = "010102"  # Example account
        
        # Current quotes data
        self.quotes_data = []
        
        # Now that all attributes are initialized, call parent constructor
        # This will indirectly call init_ui through the BaseModule class
        super().__init__(main_window)
    
    def init_ui(self):
        """Initialize the user interface."""
        # Create layout
        layout = QVBoxLayout()
        
        # Add header
        header_layout = QHBoxLayout()
        
        # Title
        title_label = QLabel("John Deere Quotes")
        title_label.setStyleSheet("font-size: 18pt; font-weight: bold;")
        header_layout.addWidget(title_label)
        
        # Add spacer
        header_layout.addStretch()
        
        # Show dealer ID in a label instead of dropdown
        dealer_label = QLabel(f"Dealer: {self.dealer_racf_id}")
        header_layout.addWidget(dealer_label)
        
        # Add filter controls
        filter_label = QLabel("Date Range:")
        header_layout.addWidget(filter_label)
        
        # Start date
        self.start_date = QDateEdit()
        self.start_date.setDate(QDate.currentDate().addDays(-30))
        self.start_date.setCalendarPopup(True)
        header_layout.addWidget(self.start_date)
        
        # End date
        self.end_date = QDateEdit()
        self.end_date.setDate(QDate.currentDate())
        self.end_date.setCalendarPopup(True)
        header_layout.addWidget(self.end_date)
        
        # Refresh button
        refresh_button = QPushButton("Refresh")
        refresh_button.clicked.connect(self.load_quotes)
        header_layout.addWidget(refresh_button)
        
        # Create quote button
        create_button = QPushButton("Create Quote")
        create_button.clicked.connect(self.show_create_quote_dialog)
        header_layout.addWidget(create_button)
        
        # Add header layout to main layout
        layout.addLayout(header_layout)
        
        # Add status label
        self.status_label = QLabel("Ready")
        layout.addWidget(self.status_label)
        
        # Add table
        self.quotes_table = QTableWidget()
        self.quotes_table.setColumnCount(7)
        self.quotes_table.setHorizontalHeaderLabels([
            "Quote ID", "Quote Name", "Customer", "Creation Date", 
            "Expiration Date", "Status", "Actions"
        ])
        
        # Set table properties
        self.quotes_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.quotes_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.quotes_table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)
        
        layout.addWidget(self.quotes_table)
        
        # Set layout
        self.setLayout(layout)
        
        # Load quotes when module is shown
        self.load_quotes()
    
    def refresh(self):
        """Refresh the quotes list."""
        self.load_quotes()
    
    def load_quotes(self):
        """Load quotes from JD Quotes API with improved error handling and retries."""
        # Try to refresh the token if available
        if hasattr(self.main_window, 'refresh_jd_token'):
            self.main_window.refresh_jd_token()
        
        if not self.quote_integration:
            self.logger.error("QuoteIntegration is not available.")
            self.show_error_message("Quote Integration Error", "Quote integration service is not available.")
            return

        if not self.dealer_racf_id:
            self.logger.error("No dealer ID specified.")
            self.show_error_message("Dealer Error", "No dealer ID specified.")
            return
            
        self.status_label.setText(f"Loading quotes for dealer {self.dealer_racf_id}...")
        self.quotes_table.setRowCount(0)  # Clear table
        
        # Number of retries for API requests
        max_retries = 3
        retry_count = 0
        
        while retry_count <= max_retries:
            try:
                # Format dates for API
                start_date = self.start_date.date().toString("MM/dd/yyyy")
                end_date = self.end_date.date().toString("MM/dd/yyyy")
                
                # Try to get quotes with default parameters
                quotes = self.quote_integration.get_dealer_quotes(
                    dealer_racf_id=self.dealer_racf_id,
                    start_date=start_date,
                    end_date=end_date
                )
                
                if quotes and isinstance(quotes, list):
                    self.quotes_data = quotes
                    self.update_quotes_table()
                    self.status_label.setText(f"Loaded {len(quotes)} quotes")
                    return
                elif retry_count < max_retries:
                    # If failed but we have retries left, try to refresh the token and retry
                    self.logger.warning(f"Failed to get quotes, retrying ({retry_count + 1}/{max_retries})...")
                    self.status_label.setText(f"Retrying quote load... ({retry_count + 1}/{max_retries})")
                    
                    # Force token refresh on retry
                    if hasattr(self.main_window, 'refresh_jd_token'):
                        self.main_window.refresh_jd_token()
                        
                    retry_count += 1
                    time.sleep(1)  # Brief delay before retry
                    continue
                else:
                    # All retries failed
                    self.quotes_data = []
                    self.update_quotes_table()
                    self.status_label.setText("No quotes found or error occurred")
                    self.show_error_message("Load Error", "Failed to retrieve quotes after multiple attempts.")
                    return
                    
            except Exception as e:
                self.logger.error(f"Error loading quotes: {str(e)}")
                self.logger.error(traceback.format_exc())  # Log the full traceback
                
                if retry_count < max_retries:
                    # If failed but we have retries left, retry
                    retry_count += 1
                    self.logger.warning(f"Exception during quote load, retrying ({retry_count}/{max_retries})...")
                    self.status_label.setText(f"Retrying after error... ({retry_count}/{max_retries})")
                    time.sleep(1)  # Brief delay before retry
                    continue
                else:
                    # All retries failed
                    self.show_error_message("Load Error", f"An error occurred while loading quotes: {str(e)}")
                    self.status_label.setText("Error loading quotes.")
                    return
    
    def show_error_message(self, title, message):
        """Show an error message dialog.
        
        Args:
            title: Dialog title
            message: Error message
        """
        try:
            QMessageBox.critical(self, title, message)
        except Exception as e:
            # Fallback if UI fails
            print(f"UI ERROR [{title}]: {message} (Original error: {e})")
    
    def update_quotes_table(self):
        """Update the quotes table with current data."""
        # Clear table
        self.quotes_table.setRowCount(0)
        
        # Add quotes to table
        for row, quote in enumerate(self.quotes_data):
            self.quotes_table.insertRow(row)
            
            # Quote ID
            self.quotes_table.setItem(row, 0, QTableWidgetItem(str(quote.get("quoteID", ""))))
            
            # Quote Name
            self.quotes_table.setItem(row, 1, QTableWidgetItem(quote.get("quoteName", "")))
            
            # Customer
            customer_data = quote.get("customerData", {})
            customer_name = f"{customer_data.get('customerFirstName', '')} {customer_data.get('customerLastName', '')}"
            self.quotes_table.setItem(row, 2, QTableWidgetItem(customer_name))
            
            # Creation Date
            creation_date = quote.get("creationDate", "")
            self.quotes_table.setItem(row, 3, QTableWidgetItem(creation_date))
            
            # Expiration Date
            expiration_date = quote.get("expirationDate", "")
            self.quotes_table.setItem(row, 4, QTableWidgetItem(expiration_date))
            
            # Status
            status_id = quote.get("quoteStatusId", 0)
            status_text = self.get_status_text(status_id)
            self.quotes_table.setItem(row, 5, QTableWidgetItem(status_text))
            
            # Actions
            actions_cell = QTableWidgetItem("Actions")
            actions_cell.setData(Qt.UserRole, quote.get("quoteID"))
            self.quotes_table.setItem(row, 6, actions_cell)
            
            # Add action buttons
            actions_widget = QWidget()
            actions_layout = QHBoxLayout(actions_widget)
            actions_layout.setContentsMargins(0, 0, 0, 0)
            
            # View button
            view_button = QPushButton("View")
            view_button.setProperty("quote_id", quote.get("quoteID"))
            view_button.clicked.connect(self.view_quote)
            actions_layout.addWidget(view_button)
            
            # Edit button
            edit_button = QPushButton("Edit")
            edit_button.setProperty("quote_id", quote.get("quoteID"))
            edit_button.clicked.connect(self.edit_quote)
            actions_layout.addWidget(edit_button)
            
            # Delete button
            delete_button = QPushButton("Delete")
            delete_button.setProperty("quote_id", quote.get("quoteID"))
            delete_button.clicked.connect(self.delete_quote)
            actions_layout.addWidget(delete_button)
            
            self.quotes_table.setCellWidget(row, 6, actions_widget)
    
    def get_status_text(self, status_id):
        """Get text representation of quote status.
        
        Args:
            status_id: Quote status ID
            
        Returns:
            Status text
        """
        status_map = {
            0: "Draft",
            1: "Active",
            2: "Expired",
            3: "Closed-Won",
            4: "Closed-Lost",
            5: "Archived"
        }
        
        return status_map.get(status_id, f"Unknown ({status_id})")
    
    def view_quote(self):
        """View a quote's details."""
        button = self.sender()
        quote_id = button.property("quote_id")
        
        if not quote_id:
            return
        
        self.main_window.show_loading(f"Loading Quote {quote_id}...")
        
        try:
            # Get quote details
            if hasattr(self.quote_integration, 'get_quote_details'):
                quote_details = self.quote_integration.get_quote_details(
                    quote_id=quote_id,
                    dealer_account_no=self.dealer_account_no
                )
                
                if quote_details:
                    self.show_quote_details_dialog(quote_details)
                else:
                    self.show_error_message(
                        "Error", 
                        f"Failed to load quote details for Quote {quote_id}"
                    )
            else:
                self.logger.error("QuoteIntegration doesn't have get_quote_details method")
                self.show_error_message(
                    "API Error", 
                    "Quote details feature is not available."
                )
                
        except Exception as e:
            self.logger.error(f"Error viewing quote: {str(e)}")
            self.logger.error(traceback.format_exc())
            self.show_error_message(
                "Error", 
                f"Failed to view quote: {str(e)}"
            )
        finally:
            self.main_window.hide_loading()
    
    def show_quote_details_dialog(self, quote_details):
        """Show dialog with quote details.
        
        Args:
            quote_details: Quote details data
        """
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Quote Details - {quote_details.get('quoteName', '')}")
        dialog.resize(800, 600)
        
        layout = QVBoxLayout(dialog)
        
        # Quote information
        info_layout = QFormLayout()
        
        # Quote ID
        info_layout.addRow("Quote ID:", QLabel(str(quote_details.get("quoteID", ""))))
        
        # Quote Name
        info_layout.addRow("Quote Name:", QLabel(quote_details.get("quoteName", "")))
        
        # Customer information
        customer_data = quote_details.get("customerData", {})
        customer_name = f"{customer_data.get('customerFirstName', '')} {customer_data.get('customerLastName', '')}"
        info_layout.addRow("Customer:", QLabel(customer_name))
        
        # Customer Email
        if customer_data.get("customerEmail"):
            info_layout.addRow("Email:", QLabel(customer_data.get("customerEmail", "")))
        
        # Creation Date
        info_layout.addRow("Created:", QLabel(quote_details.get("creationDate", "")))
        
        # Expiration Date
        info_layout.addRow("Expires:", QLabel(quote_details.get("expirationDate", "")))
        
        # Add quote info to main layout
        layout.addLayout(info_layout)
        
        # Equipment table
        equipment_label = QLabel("Equipment")
        equipment_label.setStyleSheet("font-weight: bold; font-size: 14pt;")
        layout.addWidget(equipment_label)
        
        equipment_table = QTableWidget()
        equipment_table.setColumnCount(5)
        equipment_table.setHorizontalHeaderLabels([
            "Model", "Description", "Serial #", "Year", "Price"
        ])
        
        # Add equipment to table
        equipment_data = quote_details.get("equipmentData", [])
        equipment_table.setRowCount(len(equipment_data))
        
        for row, equipment in enumerate(equipment_data):
            # Model
            equipment_table.setItem(row, 0, QTableWidgetItem(str(equipment.get("modelID", ""))))
            
            # Description
            equipment_table.setItem(row, 1, QTableWidgetItem(equipment.get("dealerSpecifiedModel", "")))
            
            # Serial #
            equipment_table.setItem(row, 2, QTableWidgetItem(equipment.get("serialNo", "")))
            
            # Year
            equipment_table.setItem(row, 3, QTableWidgetItem(str(equipment.get("yearOfManufacture", ""))))
            
            # Price
            equipment_table.setItem(row, 4, QTableWidgetItem(f"${equipment.get('listPrice', 0):,.2f}"))
        
        equipment_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(equipment_table)
        
        # Trade-in table (if any)
        trade_in_data = quote_details.get("tradeInEquipmentData", [])
        if trade_in_data:
            trade_in_label = QLabel("Trade-Ins")
            trade_in_label.setStyleSheet("font-weight: bold; font-size: 14pt;")
            layout.addWidget(trade_in_label)
            
            trade_in_table = QTableWidget()
            trade_in_table.setColumnCount(5)
            trade_in_table.setHorizontalHeaderLabels([
                "Description", "Serial #", "Year", "Hours", "Value"
            ])
            
            trade_in_table.setRowCount(len(trade_in_data))
            
            for row, trade_in in enumerate(trade_in_data):
                # Description
                trade_in_table.setItem(row, 0, QTableWidgetItem(trade_in.get("description", "")))
                
                # Serial #
                trade_in_table.setItem(row, 1, QTableWidgetItem(trade_in.get("serialNo", "")))
                
                # Year
                trade_in_table.setItem(row, 2, QTableWidgetItem(str(trade_in.get("yearOfManufacture", ""))))
                
                # Hours
                trade_in_table.setItem(row, 3, QTableWidgetItem(trade_in.get("hourMeterReading", "")))
                
                # Value
                trade_in_table.setItem(row, 4, QTableWidgetItem(f"${trade_in.get('netTradeValue', 0):,.2f}"))
            
            trade_in_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            layout.addWidget(trade_in_table)
        
        # Total values
        totals_layout = QFormLayout()
        
        # Total equipment
        total_equipment = 0
        for equipment in equipment_data:
            total_equipment += float(equipment.get("listPrice", 0))
        
        totals_layout.addRow("Total Equipment:", QLabel(f"${total_equipment:,.2f}"))
        
        # Total trade-in
        total_trade_in = quote_details.get("totalNetTradeValue", 0)
        totals_layout.addRow("Total Trade-In:", QLabel(f"${total_trade_in:,.2f}"))
        
        # Final total
        final_total = total_equipment - total_trade_in
        totals_layout.addRow("Final Total:", QLabel(f"${final_total:,.2f}"))
        
        layout.addLayout(totals_layout)
        
        # Close button
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        
        # Show dialog
        dialog.exec_()
    
    def edit_quote(self):
        """Edit a quote."""
        button = self.sender()
        quote_id = button.property("quote_id")
        
        if not quote_id:
            return
        
        # Show message indicating this feature is in development
        QMessageBox.information(
            self,
            "Feature in Development",
            f"Editing Quote {quote_id} will be available in a future update."
        )
    
    def delete_quote(self):
        """Delete a quote."""
        button = self.sender()
        quote_id = button.property("quote_id")
        
        if not quote_id:
            return
        
        # Confirm deletion
        reply = QMessageBox.question(
            self,
            "Confirm Deletion",
            f"Are you sure you want to delete Quote {quote_id}?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        self.main_window.show_loading(f"Deleting Quote {quote_id}...")
        
        try:
            # Delete quote
            if hasattr(self.quote_integration, 'delete_quote'):
                success = self.quote_integration.delete_quote(quote_id)
                
                if success:
                    self.main_window.show_notification(
                        "Success", 
                        f"Quote {quote_id} deleted successfully",
                        notification_type="success"
                    )
                    
                    # Refresh quotes list
                    self.load_quotes()
                else:
                    self.show_error_message(
                        "Error", 
                        f"Failed to delete Quote {quote_id}"
                    )
            else:
                self.logger.error("QuoteIntegration doesn't have delete_quote method")
                self.show_error_message(
                    "API Error", 
                    "Quote deletion feature is not available."
                )
                
        except Exception as e:
            self.logger.error(f"Error deleting quote: {str(e)}")
            self.logger.error(traceback.format_exc())
            self.show_error_message(
                "Error", 
                f"Failed to delete quote: {str(e)}"
            )
        finally:
            self.main_window.hide_loading()
    
    def show_create_quote_dialog(self):
        """Show dialog to create a new quote."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Create New Quote")
        dialog.resize(500, 400)
        
        layout = QVBoxLayout(dialog)
        
        # Form layout
        form_layout = QFormLayout()
        
        # Quote Name
        quote_name_edit = QLineEdit()
        quote_name_edit.setText(f"New Quote - {datetime.now().strftime('%Y-%m-%d')}")
        form_layout.addRow("Quote Name:", quote_name_edit)
        
        # Customer First Name
        customer_first_name_edit = QLineEdit()
        form_layout.addRow("First Name:", customer_first_name_edit)
        
        # Customer Last Name
        customer_last_name_edit = QLineEdit()
        form_layout.addRow("Last Name:", customer_last_name_edit)
        
        # Customer Email
        customer_email_edit = QLineEdit()
        form_layout.addRow("Email:", customer_email_edit)
        
        # Customer Phone
        customer_phone_edit = QLineEdit()
        form_layout.addRow("Phone:", customer_phone_edit)
        
        # Expiration Date
        expiration_date_edit = QDateEdit()
        expiration_date_edit.setDate(QDate.currentDate().addDays(30))
        expiration_date_edit.setCalendarPopup(True)
        form_layout.addRow("Expiration Date:", expiration_date_edit)
        
        # Notes
        notes_edit = QTextEdit()
        form_layout.addRow("Notes:", notes_edit)
        
        layout.addLayout(form_layout)
        
        # Buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        
        # Execute dialog
        if dialog.exec_() == QDialog.Accepted:
            self.create_quote(
                quote_name=quote_name_edit.text(),
                first_name=customer_first_name_edit.text(),
                last_name=customer_last_name_edit.text(),
                email=customer_email_edit.text(),
                phone=customer_phone_edit.text(),
                expiration_date=expiration_date_edit.date().toString("MM/dd/yyyy"),
                notes=notes_edit.toPlainText()
            )
    
    def create_quote(self, quote_name, first_name, last_name, email, phone, expiration_date, notes):
        """Create a new quote.
        
        Args:
            quote_name: Quote name
            first_name: Customer first name
            last_name: Customer last name
            email: Customer email
            phone: Customer phone
            expiration_date: Expiration date
            notes: Quote notes
        """
        self.main_window.show_loading("Creating Quote...")
        
        try:
            # Check if QuoteBuilder is available
            if not hasattr(QuoteBuilder, 'create_customer_data'):
                self.logger.error("QuoteBuilder missing required methods")
                self.show_error_message(
                    "Error", 
                    "Quote creation feature is not available (missing QuoteBuilder methods)."
                )
                return
            
            # Create customer data
            customer_data = QuoteBuilder.create_customer_data(
                first_name=first_name,
                last_name=last_name,
                email=email,
                phone=phone
            )
            
            # Create quote data
            quote_data = QuoteBuilder.create_basic_quote(
                dealer_racf_id=self.dealer_racf_id,
                customer_data=customer_data,
                quote_name=quote_name,
                expiration_date=expiration_date
            )
            
            # Add notes if provided
            if notes:
                quote_data["custNotes"] = notes
            
            # Create quote
            if hasattr(self.quote_integration, 'create_quote'):
                result = self.quote_integration.create_quote(quote_data)
                
                if result and isinstance(result, dict) and 'body' in result:
                    new_quote_id = result['body'].get('quoteID')
                    self.main_window.show_notification(
                        "Success", 
                        f"Quote {new_quote_id} created successfully",
                        notification_type="success"
                    )
                    
                    # Refresh quotes list
                    self.load_quotes()
                else:
                    self.show_error_message(
                        "Error", 
                        "Failed to create quote"
                    )
            else:
                self.logger.error("QuoteIntegration doesn't have create_quote method")
                self.show_error_message(
                    "API Error", 
                    "Quote creation feature is not available."
                )
                
        except Exception as e:
            self.logger.error(f"Error creating quote: {str(e)}")
            self.logger.error(traceback.format_exc())
            self.show_error_message(
                "Error", 
                f"Failed to create quote: {str(e)}"
            )
        finally:
            self.main_window.hide_loading()
    
    def search(self, search_text):
        """Search for quotes matching text.
        
        Args:
            search_text: Text to search for
            
        Returns:
            List of search results
        """
        results = []
        
        for quote in self.quotes_data:
            # Search in quote name
            if search_text.lower() in quote.get("quoteName", "").lower():
                results.append({
                    'type': 'quote',
                    'title': quote.get("quoteName", ""),
                    'id': quote.get("quoteID"),
                    'details': f"Quote ID: {quote.get('quoteID')}"
                })
                continue
            
            # Search in customer name
            customer_data = quote.get("customerData", {})
            customer_name = f"{customer_data.get('customerFirstName', '')} {customer_data.get('customerLastName', '')}".lower()
            
            if search_text.lower() in customer_name:
                results.append({
                    'type': 'quote',
                    'title': f"Quote for {customer_name}",
                    'id': quote.get("quoteID"),
                    'details': f"Quote ID: {quote.get('quoteID')}, Name: {quote.get('quoteName', '')}"
                })
                continue
        
        return results
    
    def navigate_to(self, result):
        """Navigate to a search result.
        
        Args:
            result: Search result data
        """
        if result.get('type') == 'quote' and result.get('id'):
            # Show quote details
            self.main_window.show_loading(f"Loading Quote {result.get('id')}...")
            
            try:
                # Get quote details
                if hasattr(self.quote_integration, 'get_quote_details'):
                    quote_details = self.quote_integration.get_quote_details(
                        quote_id=result.get('id'),
                        dealer_account_no=self.dealer_account_no
                    )
                    
                    if quote_details:
                        self.show_quote_details_dialog(quote_details)
                    else:
                        self.show_error_message(
                            "Error", 
                            f"Failed to load quote details for Quote {result.get('id')}"
                        )
                else:
                    self.logger.error("QuoteIntegration doesn't have get_quote_details method")
                    self.show_error_message(
                        "API Error", 
                        "Quote details feature is not available."
                    )
                    
            except Exception as e:
                self.logger.error(f"Error loading search result: {str(e)}")
                self.logger.error(traceback.format_exc())
                self.show_error_message(
                    "Error", 
                    f"Failed to load search result: {str(e)}"
                )
            finally:
                self.main_window.hide_loading()
# PriceBookModule.py - CORRECTED & UPDATED (v5)
import os
import json
import logging # Added logging import
import traceback # Added for error handling in worker
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                             QPushButton, QTableWidget, QTableWidgetItem,
                             QHeaderView, QMessageBox, QAbstractItemView)
from PyQt5.QtCore import Qt, QTimer

# Try importing SharePoint manager from the correct location
try:
    # Assuming sharepoint_manager.py is in the modules directory
    from modules.sharepoint_manager import SharePointManager
except ImportError:
    SharePointManager = None # Define as None if import fails
    # Logging might not be set up yet, use print for critical import failure
    print("CRITICAL WARNING: SharePointManager could not be imported in PriceBookModule.")

# Try importing BaseModule
try:
    from ui.base_module import BaseModule
except ImportError:
    # Dummy BaseModule if import fails - replace with actual BaseModule structure
    print("WARNING: BaseModule could not be imported in PriceBookModule. Using dummy.")
    class BaseModule(QWidget):
        # Match signature expected by main.py's call
        def __init__(self, main_window=None, parent=None):
            super().__init__(parent)
            self.main_window = main_window
        def get_title(self): return "Price Book (Dummy)"
        def get_icon_name(self): return "price_book_icon.png" # Provide a default icon name

class PriceBookModule(BaseModule): # Inherit from BaseModule
    # Define constants for display name and icon
    MODULE_DISPLAY_NAME = "Price Book"
    MODULE_ICON_NAME = "price_book_icon.png" # Example icon filename

    # Corrected __init__ signature to accept sharepoint_manager
    # Removed parent=None as BaseModule doesn't seem to take it based on other modules
    def __init__(self, main_window=None, sharepoint_manager=None):
        # Pass main_window to BaseModule constructor
        super().__init__(main_window=main_window)
        # self.main_window is set by super()
        self.setObjectName("PriceBookModule")
        # Store the passed sharepoint_manager instance
        self.sharepoint_manager = sharepoint_manager
        # Get logger from main_window if available, otherwise create one
        # Use getChild for specific module logging
        # Make logger creation safer in case main_window or its logger is None
        parent_logger = getattr(main_window, 'logger', None)
        self.logger = parent_logger.getChild("PriceBook") if parent_logger else logging.getLogger(__name__).getChild("PriceBook")

        self.price_data = []
        self.filtered_data = []
        self.headers = []
        self.ui_initialized = False # Flag to track UI setup

        # Safely get data_dir from config via main_window
        self.cache_path = None
        if hasattr(self.main_window, 'config') and self.main_window.config:
            data_dir = getattr(self.main_window.config, 'data_dir', None)
            if data_dir:
                # Ensure cache directory exists before creating file path
                cache_base_dir = getattr(self.main_window.config, 'cache_dir', None)
                if cache_base_dir:
                     os.makedirs(cache_base_dir, exist_ok=True) # Ensure base cache dir exists
                     self.cache_path = os.path.join(cache_base_dir, "pricebook_cache.json") # Place cache file in base cache dir
                else:
                     self.logger.warning("Config object found but 'cache_dir' attribute missing for cache path.")
                     # Fallback: use data_dir if cache_dir is missing
                     if data_dir:
                          self.cache_path = os.path.join(data_dir, "pricebook_cache.json")
            else:
                self.logger.warning("Config object found but 'data_dir' attribute missing for cache path.")
        else:
            self.logger.warning("Main window or config object not available for cache path.")

        if not self.cache_path:
            self.logger.error("Could not determine cache path for Price Book. Caching disabled.")

        # --- REVISED Initialization Logic v2 ---
        # Determine if UI setup should proceed
        can_initialize_ui = True
        if SharePointManager is None:
            self.logger.critical("SharePointManager class not available. Price Book disabled.")
            self._setup_error_ui("SharePoint Manager module could not be imported.")
            can_initialize_ui = False

        if can_initialize_ui and self.sharepoint_manager is None:
             self.logger.warning("SharePointManager instance was not provided. Price Book disabled.")
             self._setup_error_ui("SharePoint connection not initialized.")
             can_initialize_ui = False

        # Only setup full UI if dependencies are met
        if can_initialize_ui:
            # Check if a layout was already set (e.g., by BaseModule - unlikely but possible)
            if self.layout() is None:
                self.setup_ui() # Sets self.ui_initialized = True
                if self.ui_initialized: # Double check setup_ui didn't fail
                    self.load_data()
                else:
                    self.logger.error("UI initialization failed unexpectedly.")
                    # Attempt to set error UI if setup_ui failed but didn't set a layout
                    if self.layout() is None:
                         self._setup_error_ui("Failed to initialize UI components.")
            else:
                 self.logger.warning("__init__: Layout already exists before calling setup_ui. Skipping setup.")
                 self.ui_initialized = False # Mark as not initialized if we skipped setup
        # --- END REVISED Initialization Logic v2 ---


    # --- BaseModule Implementation ---
    def get_title(self):
        """Return the module title for display purposes."""
        return self.MODULE_DISPLAY_NAME

    def get_icon_name(self):
        """Return the icon filename for the module."""
        # NOTE: Log analysis indicated an error locating 'used_inventory_icon.png'.
        # Ensure this file exists at the expected path relative to the main application
        # or that the resource loading mechanism can find it.
        return self.MODULE_ICON_NAME

    # --- UI Setup ---
    def _setup_error_ui(self, error_message):
        """Setup UI to show an error if initialization fails."""
        # Ensure we only set the layout once
        if self.layout() is None:
            layout = QVBoxLayout() # Don't parent layout immediately
            layout.setAlignment(Qt.AlignCenter)
            error_label = QLabel(f"❌ Error: Price Book Unavailable\n\n{error_message}")
            error_label.setAlignment(Qt.AlignCenter)
            error_label.setStyleSheet("font-size: 16px; color: red;")
            layout.addWidget(error_label)
            self.setLayout(layout) # Explicitly set layout on self
            self.ui_initialized = False # Explicitly mark UI as not properly initialized
            self.logger.debug(f"_setup_error_ui: Set error layout with message: {error_message}")
        else:
            self.logger.warning("_setup_error_ui: Layout already exists, cannot add error UI.")

    def setup_ui(self):
        """Set up the UI elements for the price book module."""
        # Prevent running if UI is already initialized or has an error layout
        if self.ui_initialized or self.layout() is not None:
            self.logger.debug(f"setup_ui: Skipping as UI is already initialized (ui_initialized={self.ui_initialized}) or has layout (layout={self.layout()}).")
            return

        self.logger.debug("setup_ui: Setting up main UI components...")
        try:
            layout = QVBoxLayout() # Don't parent layout immediately
            layout.setContentsMargins(15, 15, 15, 15)
            layout.setSpacing(10)

            # --- Title ---
            title = QLabel(f"📘 {self.get_title()} (from 'App Source')") # Use get_title()
            title.setAlignment(Qt.AlignCenter)
            title.setStyleSheet("font-size: 20px; font-weight: bold; color: #2a5d24; margin-bottom: 10px;")
            layout.addWidget(title)

            # --- Search Bar ---
            search_layout = QHBoxLayout()
            search_layout.addWidget(QLabel("Search:"))
            self.search_input = QLineEdit()
            self.search_input.setPlaceholderText("Filter table by any column...")
            self.search_input.textChanged.connect(self.filter_data)
            search_layout.addWidget(self.search_input)

            self.reload_btn = QPushButton("🔄 Reload from SharePoint")
            self.reload_btn.setToolTip("Reload data from SharePoint")
            self.reload_btn.clicked.connect(self.load_data)
            search_layout.addWidget(self.reload_btn)
            layout.addLayout(search_layout)

            # --- Table Widget ---
            self.table = QTableWidget()
            self.table.setEditTriggers(QAbstractItemView.NoEditTriggers) # Read-only
            self.table.setAlternatingRowColors(True)
            self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
            self.table.setSelectionMode(QAbstractItemView.SingleSelection)
            self.table.verticalHeader().setVisible(False) # Hide row numbers
            self.table.setStyleSheet("font-size: 10pt;")
            self.table.setSortingEnabled(True) # Enable sorting
            layout.addWidget(self.table)

            # Set initial loading message
            self.table.setRowCount(1)
            self.table.setColumnCount(1)
            self.table.setHorizontalHeaderLabels(["Status"])
            self.table.setItem(0, 0, QTableWidgetItem("Initializing..."))
            self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)

            self.setLayout(layout) # Explicitly set the layout on self
            self.ui_initialized = True # Mark UI as successfully initialized
            self.logger.debug("setup_ui: UI components created successfully.")

        except Exception as e:
             self.logger.error(f"Failed to setup PriceBook UI: {e}", exc_info=True)
             self.ui_initialized = False
             # Attempt to show error UI if setup fails mid-way
             if self.layout() is None: # Check if layout was partially set
                 self._setup_error_ui(f"UI Creation Error: {e}")


    def _show_status(self, msg, timeout=3000):
        """Helper to show messages on main window status bar or print."""
        # Use main_window's update_status method if available
        if hasattr(self.main_window, 'update_status') and callable(self.main_window.update_status):
            self.main_window.update_status(msg, timeout)
        else:
            # Fallback print if method not found
            self.logger.debug(f"Status (PriceBook): {msg} (main_window or update_status not available)")

    def load_data(self):
        """Loads data from the 'App Source' sheet via SharePointManager."""
        # Check if UI is ready and SP manager exists
        if not self.ui_initialized or not self.sharepoint_manager:
             self.logger.warning("load_data: Skipping as UI not initialized or SP Manager missing.")
             if hasattr(self, 'table') and self.table: # Update table status if possible
                  self.table.setRowCount(1); self.table.setColumnCount(1)
                  self.table.setHorizontalHeaderLabels(["Status"])
                  self.table.setItem(0, 0, QTableWidgetItem("Disabled (Connection Error)"))
                  self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
             self._show_status("Price Book disabled or not ready.", 5000)
             return

        self._show_status("Loading Price Book data from SharePoint...", 0) # Persistent message
        if hasattr(self, 'table') and self.table:
            # Ensure row/column count is correct before setting item
            if self.table.rowCount() == 0: self.table.setRowCount(1)
            if self.table.columnCount() == 0: self.table.setColumnCount(1)
            if self.table.columnCount() > 0: # Check column count before setting header
                 self.table.setHorizontalHeaderLabels(["Status"])
                 self.table.setItem(0, 0, QTableWidgetItem("Loading data...")) # Update status in table
                 self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)


        # Run SP read in background thread if main_window provides it
        if hasattr(self.main_window, 'run_background_task') and callable(self.main_window.run_background_task):
            if hasattr(self, 'reload_btn') and self.reload_btn: self.reload_btn.setEnabled(False)
            # Pass the worker function itself
            self.main_window.run_background_task(
                self._load_data_worker,
                on_result=self._handle_load_result,
                on_error=self._handle_load_error,
                on_finished=self._load_finished,
                loading_message="Loading Price Book..."
            )
        else:
            # Run synchronously if no background task runner
            self.logger.warning("Running load_data synchronously (no thread pool access).")
            if hasattr(self, 'reload_btn') and self.reload_btn: self.reload_btn.setEnabled(False) # Disable while loading
            try:
                # Call worker directly (no extra args expected here)
                result_data = self._load_data_worker()
                self._handle_load_result(result_data)
            except Exception as e:
                # Manually handle error if sync
                self._handle_load_error((type(e), e, traceback.format_exc()))
            finally:
                 self._load_finished()


    # --- Definition accepts kwargs to handle unexpected arguments ---
    def _load_data_worker(self, progress_callback=None, **kwargs):
        """
        Worker function to load data from SharePoint.
        Accepts optional progress_callback and absorbs other kwargs.
        """
        # This runs in a background thread
        sheet_name = "App Source" # Sheet name specified
        self.logger.debug(f"Worker: Attempting to read sheet '{sheet_name}'")
        try:
            # Ensure manager exists before calling method
            if not self.sharepoint_manager:
                raise RuntimeError("SharePointManager instance is not available in worker.")

            # Use use_cache=False to force reload from SharePoint
            # Note: SharePointManager.read_excel_sheet needs to handle progress_callback if it uses it
            sheet_data = self.sharepoint_manager.read_excel_sheet(sheet_name, use_cache=False)

            # Example of using progress_callback if needed within this worker
            # if progress_callback:
            #     progress_callback.emit(50) # Halfway done after reading

            self.logger.debug(f"Worker: Read {len(sheet_data) if sheet_data else 'None'} rows from sheet '{sheet_name}'")

            # if progress_callback:
            #     progress_callback.emit(100) # Done

            return sheet_data # Return data or None
        except Exception as e:
            self.logger.error(f"Exception in _load_data_worker for sheet '{sheet_name}': {e}", exc_info=True)
            # Re-raise or return an error indicator if needed by error handler
            raise # Re-raise to be caught by Worker's error handling


    def _handle_load_result(self, sheet_data):
        """Handle successful data loading result from worker."""
        # Ensure table exists before modifying
        if not self.ui_initialized or not hasattr(self, 'table') or not self.table:
            self.logger.error("_handle_load_result: UI/Table widget not initialized.")
            return

        self.table.setSortingEnabled(False) # Disable sorting during population
        self.table.clearContents()
        self.table.setRowCount(0)

        if sheet_data is None:
            self.table.setRowCount(1); self.table.setColumnCount(1)
            self.table.setHorizontalHeaderLabels(["Error"])
            self.table.setItem(0, 0, QTableWidgetItem("Failed to load data from sheet 'App Source'."))
            self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
            self._show_status("Error loading Price Book.", 5000)
            # No message box here, handled by _handle_load_error if exception occurred
            return

        if not sheet_data or len(sheet_data) < 1:
            self.table.setRowCount(1); self.table.setColumnCount(1)
            self.table.setHorizontalHeaderLabels(["Info"])
            self.table.setItem(0, 0, QTableWidgetItem("No data found in sheet 'App Source'."))
            self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
            self._show_status("Price Book sheet 'App Source' is empty.", 3000)
            self.price_data = [] # Clear local data
            self.headers = []
            return

        # Assume first row is header
        self.headers = [str(h).strip() for h in sheet_data[0]]
        self.price_data = sheet_data[1:] # Store data rows only

        # Save to cache
        if self.cache_path:
            try:
                with open(self.cache_path, 'w', encoding='utf-8') as f:
                    json.dump({"headers": self.headers, "data": self.price_data}, f, indent=2) # Added indent
                self.logger.info(f"Saved price book data to cache: {self.cache_path}")
            except Exception as e:
                self.logger.error(f"Failed to save price book cache: {e}")
        else:
            self.logger.warning("Cache path not set, cannot save price book cache.")

        # Populate table (filter_data will do this)
        self.filter_data() # Apply current filter to newly loaded data
        self._show_status(f"Price Book data loaded ({len(self.price_data)} items).", 4000)


    def _handle_load_error(self, error_info):
        """Handle error result from worker."""
        # Check if error_info is a tuple/list and has the expected structure
        if isinstance(error_info, (tuple, list)) and len(error_info) == 3:
            exc_type, exc_value, tb_str = error_info # Unpack if it's safe
            # Check for the specific TypeError we saw in the logs
            if isinstance(exc_value, TypeError) and "unexpected keyword argument 'progress_callback'" in str(exc_value):
                 log_message = f"Error in Price Book worker: {exc_value}. The worker function signature might need adjustment if 'progress_callback' is expected."
                 error_display = "Worker function argument mismatch." # Simpler display message
                 self.logger.error(log_message) # Log detailed info
                 # Optionally log traceback string if helpful: self.logger.debug(f"Traceback:\n{tb_str}")
            else:
                 # Handle other 3-element errors as before
                 log_message = f"Error loading Price Book data in background: {exc_value}\nTraceback:\n{tb_str}"
                 error_display = str(exc_value)
                 self.logger.error(log_message)
        else:
            # If error_info is not as expected, log its representation directly
            log_message = f"Error loading Price Book data in background. Received unexpected error info: {error_info!r}"
            error_display = f"Unexpected error format: {error_info!r}"
            self.logger.warning("Received error_info in unexpected format during Price Book load.")
            self.logger.error(log_message) # Log the detailed or generic message

        # Update UI only if it was initialized correctly
        if self.ui_initialized and hasattr(self, 'table') and self.table:
            self.table.setSortingEnabled(False)
            self.table.clearContents()
            self.table.setRowCount(1); self.table.setColumnCount(1)
            self.table.setHorizontalHeaderLabels(["Error"])
            # Display the extracted or generic error message
            self.table.setItem(0, 0, QTableWidgetItem(f"Failed to load data: {error_display}"))
            self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        else:
             self.logger.error("Cannot display load error in table, UI not initialized.")

        self._show_status(f"Error loading Price Book: {error_display}", 5000)
        QMessageBox.warning(self, "Load Error", f"Failed to load Price Book data:\n{error_display}")


    def _load_finished(self):
        """Re-enable button when loading finishes."""
        # Ensure reload_btn exists before enabling
        if hasattr(self, 'reload_btn') and self.reload_btn:
            self.reload_btn.setEnabled(True)
        # Ensure table exists before enabling sorting
        if self.ui_initialized and hasattr(self, 'table') and self.table:
            self.table.setSortingEnabled(True)


    def filter_data(self):
        """Hides rows that do not contain the search text in any column."""
        # Ensure UI elements exist before proceeding
        if not self.ui_initialized or not hasattr(self, 'search_input') or not hasattr(self, 'table'):
             self.logger.warning("Filter called before UI is fully initialized or table/search missing.")
             return

        search_text = self.search_input.text().strip().lower()

        # If headers are not loaded yet, don't filter/populate
        if not self.headers:
            self.logger.debug("Headers not loaded yet, skipping filter.")
            # Ensure table is cleared if headers are missing
            self.table.clearContents()
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            return

        self.table.setSortingEnabled(False) # Disable sorting during filtering/population
        self.table.clearContents()
        self.table.setRowCount(0) # Clear existing rows before filtering

        # Filter data based on search term
        if search_text:
            self.filtered_data = [
                row for row in self.price_data
                # Check if row has enough elements and search term is in any cell
                if len(row) == len(self.headers) and any(search_text in str(cell).lower() for cell in row)
            ]
        else:
            self.filtered_data = self.price_data # Show all data if search is empty

        # Populate table with filtered data
        self.table.setColumnCount(len(self.headers))
        self.table.setHorizontalHeaderLabels(self.headers)
        self.table.setRowCount(len(self.filtered_data))

        for r, row_data in enumerate(self.filtered_data):
            # Check row length again for safety
            if len(row_data) == len(self.headers):
                for c, cell_value in enumerate(row_data):
                    item_text = str(cell_value) if cell_value is not None else ""
                    self.table.setItem(r, c, QTableWidgetItem(item_text))
            else:
                self.logger.warning(f"Skipping row {r+1} due to column count mismatch ({len(row_data)} vs {len(self.headers)})")


        self.table.resizeColumnsToContents()
        # Optionally stretch a specific column like Description if it exists
        try:
             # Adjust if column name is different
             desc_col_name = "Description" # Example name
             if desc_col_name in self.headers:
                 desc_col_index = self.headers.index(desc_col_name)
                 self.table.horizontalHeader().setSectionResizeMode(desc_col_index, QHeaderView.Stretch)
             else: # Fallback: stretch last column if Description not found
                 self.table.horizontalHeader().setStretchLastSection(True)
        except Exception as e:
             self.logger.warning(f"Could not stretch Description column: {e}")
             self.table.horizontalHeader().setStretchLastSection(True) # Fallback

        self.table.setSortingEnabled(True) # Re-enable sorting
        # Update status only if not currently loading
        if not hasattr(self, 'reload_btn') or self.reload_btn.isEnabled():
            self._show_status(f"Showing {len(self.filtered_data)} item(s).", 3000)


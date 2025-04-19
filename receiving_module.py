# File: modules/receiving_module.py - CORRECTED
# Desc: Module to handle receiving/audit automation, integrating TrafficAuto logic.

import os
import csv
import time # Import time module
import sys
import logging
try:
    import pyautogui # Dependency needed for this module
except ImportError:
    pyautogui = None # Define as None if import fails
    # Logging might not be set up yet, use print
    print("ERROR: PyAutoGUI not installed. Receiving Automation module will not function.")


from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                             QPushButton, QTextEdit, QMessageBox, QProgressBar)
from PyQt5.QtCore import Qt, pyqtSlot

# Assuming BaseModule is in ui.base_module
try:
    from ui.base_module import BaseModule
except ImportError:
    # Dummy BaseModule if import fails - replace with actual BaseModule structure
    print("WARNING: BaseModule could not be imported in ReceivingModule. Using dummy.")
    class BaseModule(QWidget):
        # Match signature expected by main.py's call
        def __init__(self, main_window=None, parent=None): # BaseModule init only takes main_window
            super().__init__(parent)
            self.main_window = main_window
        def get_title(self): return "Receiving (Dummy)"
        def get_icon_name(self): return "receiving_icon.png" # Provide a default icon name

# Assuming Worker is in utils.worker
try:
    # Import WorkerSignals too if your Worker uses it
    from utils.worker import Worker, WorkerSignals
except ImportError:
    logging.getLogger(__name__).critical("Failed to import Worker class from utils.worker!")
    # Define a dummy Worker if import fails to prevent immediate crash
    # Need a dummy pyqtSignal if PyQt5 isn't fully available here either
    try: from PyQt5.QtCore import pyqtSignal
    except ImportError: pyqtSignal = object

    class WorkerSignals(object): # Inherit from object for attribute assignment
        # Define dummy signals
        started = pyqtSignal() if pyqtSignal is not object else object()
        finished = pyqtSignal() if pyqtSignal is not object else object()
        error = pyqtSignal(tuple) if pyqtSignal is not object else object()
        result = pyqtSignal(object) if pyqtSignal is not object else object()
        progress = pyqtSignal(int) if pyqtSignal is not object else object()
        status = pyqtSignal(str) if pyqtSignal is not object else object()

    class Worker: # Define dummy Worker class
        signals = WorkerSignals()
        def __init__(self, fn, *args, **kwargs): pass
        def run(self): pass


# Constants for display name and icon
MODULE_DISPLAY_NAME = "Receiving Automation"
MODULE_ICON_NAME = "receiving_icon.png" # Example icon filename

class ReceivingModule(BaseModule):
    """
    Module for running the Traffic Auto GUI automation task.
    """
    # Corrected __init__ signature based on response #54 design
    # Accepts args passed from main.py's corrected _load_modules
    def __init__(self, config=None, logger=None, thread_pool=None, notification_manager=None, main_window=None, parent=None):
        # Call BaseModule's __init__ - Pass ONLY arguments it accepts (main_window)
        super().__init__(main_window=main_window) # Corrected super call

        # Store other dependencies directly
        self.config = config
        # Ensure logger exists, create default if None, get child
        self.logger = logger.getChild("ReceivingModule") if logger else logging.getLogger(__name__).getChild("ReceivingModule")
        self.thread_pool = thread_pool
        self.notification_manager = notification_manager
        # self.main_window is already set by super().__init__

        self.setObjectName("ReceivingModule")
        self.logger.info(f"Initializing {MODULE_DISPLAY_NAME}...")

        # Check dependency
        if pyautogui is None:
            self._setup_error_ui("PyAutoGUI library not installed.\nPlease run 'pip install pyautogui'.")
            return # Stop initialization

        # Get specific config values needed for automation
        # Use getattr for safety in case config is None or missing attributes
        base_path = getattr(self.config, 'base_path', os.getcwd()) if self.config else os.getcwd()
        resources_dir = getattr(self.config, 'resources_dir', os.path.join(base_path, "resources")) if self.config else os.path.join(base_path, "resources")
        data_dir = getattr(self.config, 'data_dir', os.path.join(base_path, "data")) if self.config else os.path.join(base_path, "data")

        # Use config.get() which handles defaults
        self.images_dir = self.config.get('traffic_images_dir', os.path.join(resources_dir, "traffic_images")) if self.config else os.path.join(resources_dir, "traffic_images")
        self.tasks_csv_file = self.config.get('traffic_csv_path', os.path.join(data_dir, 'traffic_tasks.csv')) if self.config else os.path.join(data_dir, 'traffic_tasks.csv')
        self.pyautogui_pause = self.config.get('pyautogui_pause', 0.2) if self.config else 0.2
        self.pyautogui_timeout = self.config.get('pyautogui_timeout', 10) if self.config else 10
        self.pyautogui_confidence = self.config.get('pyautogui_confidence', 0.8) if self.config else 0.8

        # Check if image directory exists
        if not os.path.isdir(self.images_dir):
            self.logger.error(f"Traffic Auto image directory not found: {self.images_dir}. Automation will likely fail.")
            # Optionally show a warning immediately via notification manager if available
            if self.notification_manager and hasattr(self.notification_manager, 'show_message'):
                self.notification_manager.show_message("Config Error", f"Image directory not found:\n{self.images_dir}", duration=10000)
            # Or use QMessageBox as fallback
            # QMessageBox.warning(self, "Config Error", f"Traffic Auto image directory not found:\n{self.images_dir}\nAutomation may fail.")

        # Setup UI
        self._setup_ui()

        self.logger.info(f"{MODULE_DISPLAY_NAME} initialized.")

    # --- BaseModule Implementation ---
    def get_title(self):
        """Return the display name for the module."""
        return MODULE_DISPLAY_NAME

    def get_icon_name(self):
        """Return the icon filename for the module."""
        return MODULE_ICON_NAME

    # --- UI Setup ---
    def _setup_error_ui(self, error_message):
        """Setup UI to show an error if initialization fails."""
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        error_label = QLabel(f"❌ Error: Receiving Automation Unavailable\n\n{error_message}")
        error_label.setAlignment(Qt.AlignCenter)
        error_label.setStyleSheet("font-size: 16px; color: red;")
        layout.addWidget(error_label)

    def _setup_ui(self):
        """Create the user interface for this module."""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Title
        title = QLabel(f"📦 {self.get_title()}") # Use get_title()
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(title)

        # --- Controls ---
        control_layout = QHBoxLayout()
        self.run_button = QPushButton("▶ Run Automation Task")
        self.run_button.setStyleSheet("padding: 5px 10px; font-size: 14px;")
        self.run_button.clicked.connect(self.start_automation_task)
        # Disable button if pyautogui is missing
        if pyautogui is None:
            self.run_button.setEnabled(False)
            self.run_button.setToolTip("PyAutoGUI library not installed.")
        control_layout.addWidget(self.run_button)
        control_layout.addStretch()
        layout.addLayout(control_layout)

        # --- Status Display ---
        status_layout = QHBoxLayout()
        status_label = QLabel("Status:")
        self.status_display = QLabel("Idle")
        self.status_display.setStyleSheet("font-weight: bold;")
        status_layout.addWidget(status_label)
        status_layout.addWidget(self.status_display, 1) # Allow label to expand
        layout.addLayout(status_layout)

        # --- Progress Bar (Optional) ---
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False) # Hide initially
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setRange(0, 100)
        layout.addWidget(self.progress_bar)

        # --- Output Log ---
        log_label = QLabel("Automation Log / Results:")
        layout.addWidget(log_label)
        self.output_log = QTextEdit()
        self.output_log.setReadOnly(True)
        self.output_log.setFontFamily("Courier New") # Monospaced font for logs
        self.output_log.setLineWrapMode(QTextEdit.NoWrap) # Prevent wrapping
        layout.addWidget(self.output_log)

        # self.setLayout(layout) # Already set by QVBoxLayout(self)

    # --- Automation Logic (Adapted from TrafficAuto.py) ---

    def _get_image_path(self, image_name):
        """Constructs the full path to an image file."""
        if not hasattr(self, 'images_dir') or not self.images_dir:
             self.logger.error("Image directory path not configured.")
             return None
        return os.path.join(self.images_dir, image_name)

    def _click_element(self, image_name, description, timeout=None, click=True, conf=None, region=None):
        """Clicks an element identified by an image, adapted for module context."""
        if pyautogui is None: raise RuntimeError("PyAutoGUI is not installed.")

        timeout = timeout if timeout is not None else self.pyautogui_timeout
        conf = conf if conf is not None else self.pyautogui_confidence
        img_file = self._get_image_path(image_name)

        if not img_file or not os.path.exists(img_file):
            self.logger.error(f"Image file not found for '{description}': {img_file or image_name}")
            raise FileNotFoundError(f"Image file not found: {img_file or image_name}")

        self.logger.debug(f"Attempting to find '{description}' using image: {img_file}")
        start_time = time.time()
        location = None
        original_pause = pyautogui.PAUSE
        pyautogui.PAUSE = 0.05

        try:
            while time.time() - start_time < timeout:
                try:
                    locate_args = {'confidence': conf}
                    if region: locate_args['region'] = region
                    location = pyautogui.locateCenterOnScreen(img_file, **locate_args)
                except pyautogui.ImageNotFoundException:
                    location = None
                except Exception as e:
                    self.logger.warning(f"PyAutoGUI Exception locating '{description}' ({img_file}): {e}")
                    location = None

                if location:
                    self.logger.info(f"Found '{description}' at {location}.")
                    if click:
                        try:
                            pyautogui.click(location)
                            self.logger.debug(f"Clicked '{description}' at {location}.")
                            time.sleep(0.1)
                        except Exception as e_click:
                             self.logger.error(f"Failed to click '{description}' at {location}: {e_click}")
                             raise RuntimeError(f"Failed to click {description}") from e_click
                    return location
                time.sleep(0.3)
        finally:
            pyautogui.PAUSE = original_pause

        self.logger.warning(f"Timeout: Could not find '{description}' ({img_file}) within {timeout} seconds.")
        raise TimeoutError(f"Timeout finding {description}")

    def _click_and_type(self, image_name, text_to_type, description, timeout=None, conf=None, region=None):
        """Clicks an element and types text, adapted for module context."""
        if pyautogui is None: raise RuntimeError("PyAutoGUI is not installed.")
        location = self._click_element(image_name, description, timeout, click=True, conf=conf, region=region)
        self.logger.debug(f"Typing '{text_to_type}' into '{description}'.")
        try:
            pyautogui.typewrite(str(text_to_type), interval=0.05)
            time.sleep(0.1)
        except Exception as e:
            self.logger.error(f"Failed to type '{text_to_type}' into '{description}': {e}")
            raise RuntimeError(f"Failed to type into {description}") from e

    def _process_traffic_entry(self, customer_name, stock_item):
        """Automates a single traffic entry, adapted for module context."""
        if pyautogui is None: raise RuntimeError("PyAutoGUI is not installed.")
        self.logger.info(f"--- Processing entry: Customer='{customer_name}', Stock='{stock_item}' ---")
        original_pause = pyautogui.PAUSE
        pyautogui.PAUSE = self.pyautogui_pause
        try:
            self.logger.info("Step 1: Click New Traffic Button")
            self._click_element("new_traffic.png", "New Traffic Button", timeout=15)
            time.sleep(0.5)
            self.logger.info("Step 2: Click Customer Lookup")
            self._click_element("customer_lookup.png", "Customer Lookup Button")
            time.sleep(0.5)
            self.logger.info(f"Step 3: Search for customer '{customer_name}'")
            try:
                self._click_and_type("customer_name_field.png", customer_name, "Customer Name Field")
            except (FileNotFoundError, TimeoutError, RuntimeError) as e:
                self.logger.warning(f"Could not click Customer Name Field based on image ({e}), attempting direct type.")
                pyautogui.typewrite(customer_name, interval=0.05)
                time.sleep(0.2)
            self._click_element("search_button.png", "Search Button")
            time.sleep(1)
            self.logger.info("Step 4: Click Customer Result")
            self._click_element("select_result.png", "Customer Result Row")
            time.sleep(0.5)
            self.logger.info("Step 5: Click Save (waiting for CONV01)")
            self._click_element("conv01.png", "CONV01 Confirmation", click=False, timeout=15)
            self._click_element("save.png", "Save Button (after CONV01)")
            time.sleep(2)
            self.logger.info(f"Step 6: Enter stock number '{stock_item}'")
            self._click_and_type("stock_number.png", stock_item, "Stock Number Field", timeout=15)
            time.sleep(0.3)
            self.logger.info("Step 7: Click Save (after stock number)")
            self._click_element("save.png", "Save Button (after stock number)")
            time.sleep(0.5)
            self.logger.info("Step 8: Select Comp/Pay status")
            self._click_element("pending.png", "Pending Dropdown")
            time.sleep(0.3)
            pyautogui.press("c")
            self.logger.info("Pressed 'c' for Comp/Pay status.")
            time.sleep(0.3)
            self.logger.info("Step 9: Enter Trucker")
            self._click_and_type("trucker.png", "BRITRK", "Trucker Field")
            self.logger.info("Step 10: Click Final Save")
            self._click_element("save.png", "Final Save Button", timeout=15)
            time.sleep(2)
            self.logger.info(f"--- Successfully processed entry for Customer='{customer_name}', Stock='{stock_item}' ---")
            return True
        except Exception as e:
            self.logger.error(f"Failed processing entry for Customer='{customer_name}', Stock='{stock_item}': {e}", exc_info=True)
            raise
        finally:
            pyautogui.PAUSE = original_pause

    # --- Worker Task ---
    def _run_automation_task(self, progress_callback, status_callback):
        """Runs the automation task in a background thread."""
        if pyautogui is None: raise RuntimeError("PyAutoGUI is not installed.")
        status_callback.emit("Starting automation...")
        self.logger.info(f"Worker thread started. Reading tasks from: {self.tasks_csv_file}")
        tasks = []
        try:
            data_dir = os.path.dirname(self.tasks_csv_file)
            if not os.path.exists(data_dir):
                 raise FileNotFoundError(f"Data directory for tasks CSV not found: {data_dir}")
            with open(self.tasks_csv_file, mode='r', newline='', encoding='utf-8-sig') as file:
                reader = csv.DictReader(file)
                required_headers = {'customername', 'stockitem'} # Lowercase
                actual_headers = {h.lower().strip() for h in (reader.fieldnames or [])}
                if not required_headers.issubset(actual_headers):
                     raise ValueError(f"CSV file '{self.tasks_csv_file}' must contain 'CustomerName' and 'StockItem' columns (case-insensitive). Found: {reader.fieldnames}")
                tasks = list(reader)
        except FileNotFoundError:
            self.logger.error(f"Tasks CSV file not found: {self.tasks_csv_file}")
            raise
        except Exception as e:
            self.logger.error(f"Error reading CSV file {self.tasks_csv_file}: {e}", exc_info=True)
            raise

        if not tasks:
            return {"summary": "No tasks found in the CSV file.", "details": []}

        total_tasks = len(tasks)
        results_details = []
        success_count = 0
        fail_count = 0
        status_callback.emit(f"Found {total_tasks} tasks. Starting in 3 seconds...")
        time.sleep(3)

        for i, task in enumerate(tasks):
            # Use case-insensitive lookup for headers
            customer_name = next((task[h] for h in task if h.lower().strip() == 'customername'), '').strip()
            stock_item = next((task[h] for h in task if h.lower().strip() == 'stockitem'), '').strip()
            task_num = i + 1
            progress = int((task_num / total_tasks) * 100)
            progress_callback.emit(progress)

            if not customer_name or not stock_item:
                self.logger.warning(f"Skipping task {task_num}: Missing CustomerName or StockItem.")
                results_details.append({'Task': task_num, 'Customer': customer_name, 'Stock': stock_item, 'Status': 'Skipped - Missing Data'})
                continue

            status_callback.emit(f"Processing Task {task_num}/{total_tasks}: {customer_name} - {stock_item}")
            try:
                success = self._process_traffic_entry(customer_name, stock_item)
                if success:
                    results_details.append({'Task': task_num, 'Customer': customer_name, 'Stock': stock_item, 'Status': 'Success'})
                    success_count += 1
            except Exception as e:
                 results_details.append({'Task': task_num, 'Customer': customer_name, 'Stock': stock_item, 'Status': f'Failed: {e}'})
                 fail_count += 1
                 # Continue processing remaining tasks

            time.sleep(0.5) # Small pause between entries

        summary = f"Automation complete. Processed {total_tasks} tasks. Success: {success_count}, Failed: {fail_count}."
        self.logger.info(summary)
        return {"summary": summary, "details": results_details}


    # --- Slots for UI Interaction and Worker Signals ---
    @pyqtSlot()
    def start_automation_task(self):
        """Slot connected to the 'Run Automation' button."""
        if pyautogui is None:
             QMessageBox.critical(self, "Error", "PyAutoGUI library is not installed. Cannot run automation.")
             return
        if not self.thread_pool:
            self.logger.error("ThreadPool not available. Cannot start background task.")
            QMessageBox.critical(self, "Error", "ThreadPool is not initialized. Cannot run automation.")
            return
        if not Worker: # Check if Worker class was imported
            self.logger.error("Worker class not available. Cannot start background task.")
            QMessageBox.critical(self, "Error", "Worker class is not available. Cannot run automation.")
            return

        self.logger.info("'Run Automation' button clicked.")
        self.run_button.setEnabled(False)
        self.status_display.setText("Starting...")
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.output_log.clear()
        self.output_log.append("Starting automation task...")

        # Use the run_background_task helper from main_window if available
        if hasattr(self.main_window, 'run_background_task') and callable(self.main_window.run_background_task):
             self.main_window.run_background_task(
                 self._run_automation_task,
                 # Pass our specific result/error handlers
                 on_result=self.handle_automation_result,
                 on_error=self.handle_automation_error,
                 on_finished=self.handle_automation_finished,
                 progress_callback_slot=self.update_progress_bar, # Pass progress slot
                 loading_message="Running Traffic Automation..." # Custom loading message
             )
        else:
             # Fallback: Run directly in thread pool if helper not available (less ideal)
             self.logger.warning("main_window.run_background_task not found. Running worker directly.")
             worker = Worker(self._run_automation_task)
             worker.signals.status.connect(self.update_status_display)
             worker.signals.progress.connect(self.update_progress_bar)
             worker.signals.result.connect(self.handle_automation_result)
             worker.signals.error.connect(self.handle_automation_error) # Worker emits tuple now
             worker.signals.finished.connect(self.handle_automation_finished)
             self.thread_pool.start(worker)


    @pyqtSlot(str)
    def update_status_display(self, status_text):
        """Update the status label and append to log."""
        self.status_display.setText(status_text)
        self.output_log.append(status_text)

    @pyqtSlot(int)
    def update_progress_bar(self, progress_value):
        """Update the progress bar."""
        self.progress_bar.setValue(progress_value)

    @pyqtSlot(object) # Result is now a dictionary {"summary": ..., "details": ...}
    def handle_automation_result(self, result):
        """Handle the result when the automation finishes successfully."""
        self.logger.info(f"Automation task finished successfully.")
        summary = "Automation finished."
        if isinstance(result, dict):
            summary = result.get("summary", summary)
            details = result.get("details", [])
            self.output_log.append("\n--- Results ---")
            for item in details:
                 # Use .get() with defaults for safety
                 self.output_log.append(f"Task {item.get('Task', '?')}: {item.get('Customer','?')}/{item.get('Stock','?')} -> {item.get('Status', 'Unknown')}")
            self.output_log.append(f"\n{summary}")
        elif isinstance(result, str): # Handle simple string result if needed
            summary = result
            self.output_log.append(f"\nSummary: {summary}")

        self.status_display.setText(summary)
        self.progress_bar.setValue(100) # Ensure progress is full on success
        if self.notification_manager and hasattr(self.notification_manager, 'show_message'):
             self.notification_manager.show_message(self.get_title(), summary) # Use get_title()


    @pyqtSlot(tuple) # Worker emits tuple (exc_type, exc_value, tb_str)
    def handle_automation_error(self, error_info):
        """Handle errors emitted by the worker thread."""
        exc_type, exc_value, tb_str = error_info
        error_message = str(exc_value) # Get the string representation of the exception
        self.logger.error(f"Automation task failed: {error_message}\nTraceback:\n{tb_str}")
        self.status_display.setText(f"Error: {error_message}")
        self.output_log.append(f"\n--- ERROR ---")
        self.output_log.append(error_message)
        # Show traceback in log for debugging, not usually in message box
        # self.output_log.append(f"\nTraceback:\n{tb_str}")
        QMessageBox.warning(self, "Automation Error", f"An error occurred during automation:\n\n{error_message}")

    @pyqtSlot()
    def handle_automation_finished(self):
        """Handle cleanup when the worker thread finishes."""
        # This slot is connected to the worker's finished signal
        # It runs regardless of success or error
        self.logger.info("Automation worker finished.")
        self.run_button.setEnabled(True)
        # Optionally hide progress bar or reset it
        # self.progress_bar.setVisible(False)
        # self.progress_bar.setValue(0) # Optionally reset progress bar

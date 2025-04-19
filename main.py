# main.py - V3 - Added missing _on_task_result method
import os
import sys
import logging
import traceback
from logging.handlers import RotatingFileHandler
from PyQt5.QtWidgets import (QApplication, QMainWindow, QStackedWidget, QWidget,
                            QVBoxLayout, QLabel, QMessageBox, QToolBar, QAction,
                            QSizePolicy, QStatusBar, QDesktopWidget) # Added QDesktopWidget
# Added QTimer, pyqtSlot, QSize, QThreadPool, QIcon, QPixmap
from PyQt5.QtCore import Qt, QTimer, pyqtSlot, QSize, QThreadPool, QPoint # Added QPoint
from PyQt5.QtGui import QIcon, QPixmap

# --- Core Utilities & Managers ---
try:
    from utils.config import Config
    from utils.theme_manager import ThemeManager
    from utils.general_utils import get_resource_path
    from utils.cache_handler import CacheHandler
    from utils.csv_handler import CSVHandler
    from utils.oauth_client import JohnDeereOAuthClient
    from utils.jd_auth_manager import JDAuthManager
    from utils.worker import Worker
except ImportError as e:
     print(f"CRITICAL ERROR: Failed to import core utility/manager: {e}", file=sys.stderr)
     traceback.print_exc(file=sys.stderr)
     sys.exit(1)

# Import SharePointManager carefully
try:
    from modules.sharepoint_manager import SharePointManager
except ImportError as e:
    print(f"WARNING: SharePointManager could not be imported. SharePoint features will fail. Error: {e}", file=sys.stderr)
    SharePointManager = None

# --- UI Elements ---
try:
    from ui.splash_screen import SplashScreen
except ImportError as e:
    print(f"WARNING: SplashScreen could not be imported. Error: {e}", file=sys.stderr)
    SplashScreen = None
try:
    from ui.loading_widget import LoadingWidget
except ImportError as e:
     print(f"WARNING: LoadingWidget could not be imported. Error: {e}", file=sys.stderr)
     LoadingWidget = None
try:
    from ui.notification import Notification
except ImportError as e:
     print(f"WARNING: Notification could not be imported. Error: {e}", file=sys.stderr)
     Notification = None

# --- API Integrations ---
try:
    from api.QuoteIntegration import QuoteIntegration
except ImportError as e:
    print(f"WARNING: QuoteIntegration could not be imported. JD Quotes module may fail. Error: {e}", file=sys.stderr)
    QuoteIntegration = None
try:
    from api.MaintainQuotesAPI import MaintainQuotesAPI
except ImportError as e:
    print(f"WARNING: MaintainQuotesAPI could not be imported. QuoteIntegration/JD Quotes module may fail. Error: {e}", file=sys.stderr)
    MaintainQuotesAPI = None


# --- Application Modules ---
IMPORTED_MODULES = {}
MODULE_IMPORT_FAILED = False
MODULE_CLASSES_TO_LOAD = [
    ("modules.home_module", "HomeModule"),
    ("modules.deal_form_module", "DealFormModule"),
    ("modules.calendar_module", "CalendarModule"),
    ("modules.jd_quotes_module", "JDQuotesModule"),
    ("modules.calculator_module", "CalculatorModule"),
    ("modules.price_book_module", "PriceBookModule"),
    ("modules.recent_deals_module", "RecentDealsModule"),
    ("modules.receiving_module", "ReceivingModule"),
    ("modules.used_inventory_module", "UsedInventoryModule"),
]

for module_path, class_name in MODULE_CLASSES_TO_LOAD:
    try:
        module = __import__(module_path, fromlist=[class_name])
        IMPORTED_MODULES[class_name] = getattr(module, class_name)
        print(f"Successfully imported {class_name} from {module_path}")
    except ImportError as e:
         print(f"CRITICAL: Failed to import application module '{class_name}' from '{module_path}': {e}", file=sys.stderr)
         traceback.print_exc(file=sys.stderr)
         MODULE_IMPORT_FAILED = True
    except AttributeError as e:
         print(f"CRITICAL: Class '{class_name}' not found in module '{module_path}': {e}", file=sys.stderr)
         traceback.print_exc(file=sys.stderr)
         MODULE_IMPORT_FAILED = True


# --- Logging Setup ---
def setup_logging(log_dir="logs", log_level_str="INFO"):
    """Configure application logging."""
    if not os.path.exists(log_dir):
        try:
            os.makedirs(log_dir, exist_ok=True)
        except OSError as e:
            print(f"ERROR: Failed to create log directory {log_dir}: {e}", file=sys.stderr)
            log_dir = "."
            print(f"WARNING: Attempting to log to current directory: {os.path.abspath(log_dir)}", file=sys.stderr)

    log_file = os.path.join(log_dir, "application.log")
    log_level = getattr(logging, log_level_str.upper(), logging.INFO)
    log_format = logging.Formatter(
        '%(asctime)s - %(name)s [%(levelname)s] (%(threadName)s) %(message)s'
    )
    root_logger = logging.getLogger()
    root_logger.handlers.clear()
    log_setup_success = False
    try:
        file_handler = RotatingFileHandler(
            log_file, maxBytes=5*1024*1024, backupCount=3, encoding='utf-8'
        )
        file_handler.setFormatter(log_format)
        file_handler.setLevel(log_level)
        root_logger.addHandler(file_handler)
        log_setup_success = True
    except Exception as e:
        print(f"Error setting up file logger '{log_file}': {e}", file=sys.stderr)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(log_format)
    console_handler.setLevel(log_level)
    root_logger.addHandler(console_handler)
    root_logger.setLevel(log_level)

    if log_setup_success:
         logging.info(f"Logging initialized. Log file: {log_file}")
    else:
         logging.error(f"File logging failed. Logging to console only. Attempted log file: {log_file}")

    return logging.getLogger("main_app")


# --- Exception Handling ---
def log_exception_hook(exc_type, exc_value, exc_traceback):
    """Handle uncaught exceptions by logging them."""
    logger_hook = logging.getLogger("GlobalExceptionHook")
    if not logger_hook.hasHandlers() or not logging.getLogger().hasHandlers():
        print("FATAL ERROR: Unhandled exception occurred before/during logging setup.", file=sys.stderr)
        traceback.print_exception(exc_type, exc_value, exc_traceback, file=sys.stderr)
    else:
        if issubclass(exc_type, KeyboardInterrupt):
            logger_hook.warning("Application interrupted by user (KeyboardInterrupt).")
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        tb_lines = traceback.format_exception(exc_type, exc_value, exc_traceback)
        error_message = f"Unhandled exception:\n{''.join(tb_lines)}"
        logger_hook.critical(error_message)
    try:
        app_instance = QApplication.instance()
        if app_instance:
             QMessageBox.critical(None, "Critical Error", f"An unexpected error occurred: {exc_value}\nPlease check application.log for details.")
        else:
             print(f"CRITICAL ERROR (GUI not ready): {exc_value}", file=sys.stderr)
    except Exception as mb_error:
        logger_hook.error(f"Could not display Qt error message box: {mb_error}")
        print(f"ERROR: Could not display Qt error message box: {mb_error}", file=sys.stderr)


# --- Main Application Class ---
class MainWindow(QMainWindow):
    HOME_MODULE = "HomeModule"
    DEAL_FORM_MODULE = "DealFormModule"
    CALENDAR_MODULE = "CalendarModule"
    JD_QUOTES_MODULE = "JDQuotesModule"
    CALCULATOR_MODULE = "CalculatorModule"
    PRICE_BOOK_MODULE = "PriceBookModule"
    RECENT_DEALS_MODULE = "RecentDealsModule"
    RECEIVING_MODULE = "ReceivingModule"
    USED_INVENTORY_MODULE = "UsedInventoryModule"

    def __init__(self, config, logger, cache_handler, csv_handler, thread_pool,
                 oauth_client, jd_auth_manager, sharepoint_manager,
                 quote_integration):
        """Initialize the main window."""
        super().__init__()
        self.config = config
        self.logger = logger if logger else logging.getLogger("MainWindow")
        self.logger.info("Initializing MainWindow...")

        self.cache_handler = cache_handler
        self.csv_handler = csv_handler
        self.thread_pool = thread_pool
        self.oauth_client = oauth_client
        self.jd_auth_manager = jd_auth_manager
        self.sharepoint_manager = sharepoint_manager
        self.quote_integration = quote_integration

        self.active_notifications = []
        self.setWindowTitle(self.config.get("app_title", "BC Application"))
        self.resources_dir = getattr(self.config, 'resources_dir', None)
        if not self.resources_dir:
             self.logger.error("config.resources_dir not found! Resource loading will likely fail.")
        else:
             self.logger.info(f"Using resources directory: {self.resources_dir}")

        app_icon_filename = self.config.get("app_icon", "app_icon.png")
        app_icon_path = get_resource_path(app_icon_filename, self.resources_dir)
        if app_icon_path and os.path.exists(app_icon_path):
             self.setWindowIcon(QIcon(app_icon_path))
             self.logger.debug(f"App icon set from: {app_icon_path}")
        else:
             self.logger.warning(f"App icon not found or path invalid: {app_icon_path}")
        self.resize(self.config.get("window_width", 1200), self.config.get("window_height", 800))

        self.stackedWidget = QStackedWidget()
        self.setCentralWidget(self.stackedWidget)

        if LoadingWidget:
             self.loading_widget = LoadingWidget(self)
             self.loading_widget.setGeometry(self.rect())
             self.loading_widget.hide()
        else:
             self.loading_widget = None
             self.logger.warning("LoadingWidget class not available.")

        self.modules = {}
        self.module_actions = {}
        self._create_toolbar()
        self._create_status_bar()

        self.logger.info("MainWindow basic initialization complete (module loading deferred).")


    def _load_modules_and_init(self):
        """Loads modules, creates actions, and sets initial state."""
        self.logger.info("Loading application modules...")
        self.show_loading("Initializing modules...")

        self._load_modules()

        self.hide_loading()
        self.logger.info("Module loading process finished.")

        if self.HOME_MODULE in self.modules:
            self.switch_module(self.HOME_MODULE)
            self.update_status("Ready")
        elif self.modules:
             first_available_module = next(iter(self.modules.keys()), None)
             if first_available_module:
                  self.logger.warning(f"Home module failed or not available. Switching to first loaded module: {first_available_module}")
                  self.switch_module(first_available_module)
                  self.update_status(f"{first_available_module} Ready")
             else:
                  self._show_critical_load_failure("No modules initialized successfully.")
        else:
             self._show_critical_load_failure("No application modules found or loaded.")

        self.logger.info("Initial module set.")

    def _show_critical_load_failure(self, message):
        """Handles cases where no modules could be loaded."""
        self.logger.error(message)
        placeholder = QLabel(f"CRITICAL ERROR:\n{message}\n\nPlease check logs.")
        placeholder.setAlignment(Qt.AlignCenter)
        placeholder.setStyleSheet("color: red; font-size: 16px; font-weight: bold;")
        if hasattr(self, 'stackedWidget'):
            while self.stackedWidget.count() > 0:
                 widget = self.stackedWidget.widget(0)
                 self.stackedWidget.removeWidget(widget)
                 if widget: widget.deleteLater()

            self.stackedWidget.addWidget(placeholder)
            self.stackedWidget.setCurrentWidget(placeholder)
        self.update_status(f"Error: {message}")
        QMessageBox.critical(self, "Module Load Failure", message)


    def _create_toolbar(self):
        """Create the main application toolbar."""
        self.logger.debug("Creating toolbar...")
        self.toolbar = QToolBar("Main Toolbar")
        icon_size = self.config.get("toolbar_icon_size", 24)
        self.toolbar.setIconSize(QSize(icon_size, icon_size))
        self.addToolBar(Qt.LeftToolBarArea, self.toolbar)
        self.toolbar.setMovable(False)


    def _create_status_bar(self):
        """Create the status bar."""
        self.logger.debug("Creating status bar...")
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)


    def _load_modules(self):
        """Load and initialize all application modules, passing only necessary dependencies."""
        global IMPORTED_MODULES

        if not IMPORTED_MODULES:
             self.logger.error("No application module classes were successfully imported. Cannot load modules.")
             return

        module_keys_to_load = list(IMPORTED_MODULES.keys())
        toolbar_actions = []

        for module_key in module_keys_to_load:
            ModuleClass = IMPORTED_MODULES[module_key]
            module_instance = None
            try:
                self.logger.info(f"Attempting to initialize module: {module_key}")
                module_args = {"main_window": self}

                if module_key == self.DEAL_FORM_MODULE:
                    module_args["sharepoint_manager"] = self.sharepoint_manager
                elif module_key == self.JD_QUOTES_MODULE:
                    module_args["logger"] = self.logger.getChild(module_key)
                    module_args["quote_integration"] = self.quote_integration
                    if not self.quote_integration: self.logger.warning(f"QuoteIntegration dependency missing for {module_key}")
                elif module_key == self.RECENT_DEALS_MODULE:
                    module_args["data_path"] = getattr(self.config, 'data_dir', None)
                elif module_key == self.USED_INVENTORY_MODULE:
                    module_args["sharepoint_manager"] = self.sharepoint_manager
                elif module_key == self.PRICE_BOOK_MODULE:
                    module_args["sharepoint_manager"] = self.sharepoint_manager

                self.logger.debug(f"Instantiating {module_key} with args: {list(module_args.keys())}")
                module_instance = ModuleClass(**module_args)

                self.stackedWidget.addWidget(module_instance)
                self.modules[module_key] = module_instance
                self.logger.info(f"Module '{module_key}' initialized and added to stacked widget.")

                display_name = module_key.replace("Module", "").replace("_", " ").title()
                if hasattr(module_instance, 'get_title') and callable(module_instance.get_title):
                     display_name = module_instance.get_title()

                icon_name = f"{module_key.replace('Module', '').lower()}_icon.png"
                if hasattr(module_instance, 'get_icon_name') and callable(module_instance.get_icon_name):
                     icon_name_from_module = module_instance.get_icon_name()
                     if icon_name_from_module: icon_name = icon_name_from_module

                icon_path = get_resource_path(icon_name, self.resources_dir)
                action_icon = QIcon()
                if icon_path and os.path.exists(icon_path):
                    action_icon = QIcon(icon_path)
                    self.logger.debug(f"Found icon for {module_key} at {icon_path}")
                else:
                    self.logger.warning(f"Icon not found for module '{module_key}'. Looked for '{icon_name}' at path: {icon_path}")

                action = QAction(action_icon, display_name, self)
                action.setStatusTip(f"Switch to {display_name}")
                action.triggered.connect(lambda checked=False, key=module_key: self.switch_module(key))
                self.module_actions[module_key] = action
                toolbar_actions.append(action)
                self.logger.debug(f"Toolbar action created for '{module_key}'.")

            except AttributeError as ae:
                 self.logger.error(f"FAILED to initialize module '{module_key}' due to AttributeError: {ae}", exc_info=True)
                 error_message = f"Attribute Error loading module:\n{module_key}\n\n{ae}\n\nCheck logs for details."
                 error_label = QLabel(error_message)
                 error_label.setAlignment(Qt.AlignCenter)
                 error_label.setStyleSheet("color: red; padding: 20px;")
                 error_label.setWordWrap(True)
                 self.stackedWidget.addWidget(error_label)
                 if module_key in self.module_actions:
                      self.module_actions[module_key].setEnabled(False)
                      self.module_actions[module_key].setStatusTip(f"{display_name} failed to load")
                 QMessageBox.warning(self, "Module Load Error", error_message)

            except Exception as e:
                self.logger.error(f"FAILED to initialize module '{module_key}': {e}", exc_info=True)
                error_message = f"Error loading module:\n{module_key}\n\n{e}\n\nCheck logs for details."
                error_label = QLabel(error_message)
                error_label.setAlignment(Qt.AlignCenter)
                error_label.setStyleSheet("color: red; padding: 20px;")
                error_label.setWordWrap(True)
                self.stackedWidget.addWidget(error_label)
                if module_key in self.module_actions:
                     self.module_actions[module_key].setEnabled(False)
                     self.module_actions[module_key].setStatusTip(f"{display_name} failed to load")
                QMessageBox.warning(self, "Module Load Error", error_message)

        for action in toolbar_actions:
            self.toolbar.addAction(action)
        self.logger.debug("Toolbar actions added.")


    @pyqtSlot(str)
    def switch_module(self, module_key: str):
        """Switch the visible module in the stacked widget."""
        if module_key in self.modules:
            widget_to_show = self.modules[module_key]
            if self.stackedWidget.currentWidget() != widget_to_show:
                self.logger.info(f"Switching to module: {module_key}")
                self.stackedWidget.setCurrentWidget(widget_to_show)
                display_name = module_key.replace("Module", "")
                if hasattr(widget_to_show, 'get_title') and callable(widget_to_show.get_title):
                     display_name = widget_to_show.get_title()
                self.update_status(f"{display_name} loaded")
                if hasattr(widget_to_show, 'refresh') and callable(widget_to_show.refresh):
                    self.logger.debug(f"Calling refresh for {module_key}")
                    try:
                        widget_to_show.refresh()
                    except Exception as e:
                         self.logger.error(f"Error calling refresh on module {module_key}: {e}", exc_info=True)
        else:
            self.logger.warning(f"Attempted to switch to unknown or failed module key: {module_key}")
            if module_key in self.module_actions and not self.module_actions[module_key].isEnabled():
                 QMessageBox.warning(self, "Module Error", f"The '{module_key.replace('Module','')}' module failed to load correctly.")


    def resizeEvent(self, event):
        """Handle window resize events."""
        if self.loading_widget:
             self.loading_widget.setGeometry(self.rect())
        self._reposition_notifications()
        super().resizeEvent(event)

    def moveEvent(self, event):
        """Handle window move events."""
        self._reposition_notifications()
        super().moveEvent(event)


    @pyqtSlot(str, int)
    def update_status(self, message: str, timeout: int = 5000):
        """Update the status bar message."""
        if hasattr(self, 'statusBar') and self.statusBar:
             self.statusBar.showMessage(message, timeout)


    @pyqtSlot(str)
    def show_loading(self, message: str = "Loading..."):
        """Show the loading overlay if available."""
        if self.loading_widget:
             self.loading_widget.set_message(message)
             self.loading_widget.setGeometry(self.rect())
             self.loading_widget.show()
             self.loading_widget.raise_()
             QApplication.processEvents()


    @pyqtSlot()
    def hide_loading(self):
        """Hide the loading overlay if available."""
        if self.loading_widget:
             self.loading_widget.hide()


    @pyqtSlot(str, str, str, int)
    def show_notification(self, title: str, message: str,
                          notification_type: str = Notification.INFO,
                          duration: int = 5000):
        """Creates and shows a new notification widget."""
        if Notification:
            try:
                self.logger.debug(f"Creating notification: {title} - {message}")
                notification = Notification(parent=self,
                                            title=title,
                                            message=message,
                                            notification_type=notification_type,
                                            duration=duration)
                notification.closed.connect(lambda n=notification: self._remove_notification(n))
                self.active_notifications.append(notification)
                self._position_notification(notification)
                notification.show()
            except Exception as e:
                self.logger.error(f"Failed to create or show notification: {e}", exc_info=True)
                QMessageBox.information(self, title, message)
        else:
            self.logger.warning(f"Notification class not available. Showing MessageBox instead for: {title}")
            QMessageBox.information(self, title, message)


    def _position_notification(self, notification: Notification):
        """Calculate and set the position for a new notification."""
        spacing = 10
        window_rect = self.geometry()
        screen_rect = QDesktopWidget().availableGeometry(self)
        top_right = screen_rect.topRight()
        base_x = top_right.x() - notification.width() - spacing
        base_y = top_right.y() + spacing
        vertical_offset = 0
        for existing_notification in self.active_notifications[:-1]:
            if existing_notification and not existing_notification.isHidden():
                vertical_offset += existing_notification.height() + spacing
        notification.move(QPoint(base_x, base_y + vertical_offset))


    def _reposition_notifications(self):
        """Repositions all active notifications, e.g., after move/resize."""
        spacing = 10
        window_rect = self.geometry()
        screen_rect = QDesktopWidget().availableGeometry(self)
        top_right = screen_rect.topRight()
        base_x = top_right.x() - (self.active_notifications[0].width() if self.active_notifications else 300) - spacing
        current_y = top_right.y() + spacing
        visible_notifications = [n for n in self.active_notifications if n and not n.isHidden()]
        for notification in visible_notifications:
            notification.move(QPoint(base_x, current_y))
            current_y += notification.height() + spacing


    def _remove_notification(self, notification: Notification):
        """Remove notification from the tracking list when closed."""
        if notification in self.active_notifications:
             self.active_notifications.remove(notification)
             self._reposition_notifications()


    # --- Background Task Runner (TypeError Fixed V2 + AttributeError Fix) ---
    def run_background_task(self, task_function, *args, **kwargs):
        """Runs a function in a background thread using QThreadPool and Worker."""
        if not self.thread_pool:
             self.logger.error("ThreadPool not available for background task.")
             QMessageBox.critical(self, "Error", "Background task runner is not available.")
             return
        if not Worker:
             self.logger.error("Worker class not available for background task.")
             QMessageBox.critical(self, "Error", "Background worker class is not available.")
             return

        # *** FIX: Use getattr with default None for the default callback ***
        on_result_slot = kwargs.pop('on_result', getattr(self, '_on_task_result', None))
        # If the default method doesn't exist, on_result_slot will be None
        # Connect only if the slot is callable
        # *** End Fix ***

        on_error_slot = kwargs.pop('on_error', self._on_task_error) # Slot expects tuple
        on_finished_slot = kwargs.pop('on_finished', self.hide_loading)
        progress_slot = kwargs.pop('progress_callback_slot', None)
        loading_msg = kwargs.pop('loading_message', "Processing...")

        self.show_loading(loading_msg)
        worker = Worker(task_function, *args, **kwargs)

        # Connect signals
        # Connect result signal only if a valid slot is provided or found
        if callable(on_result_slot):
             worker.signals.result.connect(on_result_slot)
        else:
             self.logger.debug("No valid 'on_result' slot provided or found for background task.")

        try:
             worker.signals.error.connect(on_error_slot)
             self.logger.debug("Connected worker error signal directly.")
        except TypeError as te:
             # Log error but don't crash the setup here, let the global hook catch it if it happens later
             self.logger.error(f"Potential issue connecting worker error signal: {te}. Slot: {on_error_slot}")
             # Fallback lambda (might still cause issues if slot signature is wrong)
             try:
                  worker.signals.error.connect(lambda err_info: on_error_slot(err_info))
                  self.logger.warning("Used lambda fallback for worker error signal connection.")
             except Exception as lambda_e:
                   self.logger.error(f"CRITICAL: Lambda fallback for error signal failed: {lambda_e}")
                   # If even lambda fails, don't proceed
                   self.hide_loading()
                   QMessageBox.critical(self, "Error", f"Failed setup background task error handler:\n{lambda_e}")
                   return


        worker.signals.finished.connect(on_finished_slot)
        if progress_slot:
            worker.signals.progress.connect(progress_slot)

        self.thread_pool.start(worker)
        self.logger.debug(f"Started background task: {getattr(task_function, '__name__', 'unknown')}")


    # --- Slot for task result (default) ---
    # Added the missing method
    @pyqtSlot(object)
    def _on_task_result(self, result):
        """Default handler for background task success."""
        self.logger.info(f"Background task finished successfully. Result: {result}")
        # Example: Show a success notification
        # self.show_notification("Task Complete", "Background process finished.", notification_type=Notification.SUCCESS)

    # Ensure slot accepts tuple
    @pyqtSlot(tuple)
    def _on_task_error(self, error_info: tuple):
        """Default handler for background task error."""
        self.logger.debug(f"Received error signal with type: {type(error_info)}, value: {error_info}")
        if isinstance(error_info, tuple) and len(error_info) == 3:
             exc_type, exc_value, tb_str = error_info
             self.logger.error(f"Background task failed: {exc_value}\nTraceback:\n{tb_str}")
             self.update_status(f"Error during task: {exc_value}", 10000)
             self.show_notification("Task Error", f"An error occurred:\n{exc_value}", notification_type=Notification.ERROR, duration=0)
        else:
             self.logger.error(f"Background task failed with unexpected error signal type: {type(error_info)}. Value: {error_info}")
             self.update_status(f"Unknown background task error: {error_info}", 10000)
             self.show_notification("Task Error", f"An unexpected error occurred:\n{error_info}", notification_type=Notification.ERROR, duration=0)
        self.hide_loading()


    def closeEvent(self, event):
        """Handle application close event."""
        self.logger.info("Close event triggered. Shutting down...")
        current_widget = self.stackedWidget.currentWidget()
        current_module_key = None
        for key, mod_instance in self.modules.items():
            if mod_instance == current_widget:
                current_module_key = key
                break

        if current_module_key and current_widget and hasattr(current_widget, 'save_state') and callable(current_widget.save_state):
            self.logger.info(f"Attempting to save state for active module '{current_module_key}' before full shutdown...")
            try:
                current_widget.save_state()
                self.logger.info(f"Successfully saved state for active module '{current_module_key}'.")
            except Exception as e:
                 self.logger.error(f"Error saving state for active module '{current_module_key}': {e}", exc_info=False)
                 if isinstance(e, RuntimeError) and "deleted" in str(e).lower():
                      self.logger.warning(f"Active module '{current_module_key}' widget likely deleted even during early save attempt.")
                 QMessageBox.warning(self, "Save State Warning", f"Could not fully save state for {current_module_key} during shutdown.\nError: {e}")

        if self.sharepoint_manager and hasattr(self.sharepoint_manager, 'stop_background_update'):
            self.logger.info("Stopping SharePoint Manager background tasks...")
            try:
                self.sharepoint_manager.stop_background_update()
            except Exception as e:
                 self.logger.error(f"Error stopping SharePoint Manager: {e}")
        if self.thread_pool:
            self.logger.info("Waiting for background threads to finish...")
            self.thread_pool.clear()
            if not self.thread_pool.waitForDone(2000):
                 self.logger.warning("Some background threads did not finish cleanly.")

        self.logger.info("Saving state for potentially non-active modules (best effort)...")
        for key, module_instance in self.modules.items():
            if module_instance != current_widget:
                 if module_instance and hasattr(module_instance, 'save_state') and callable(module_instance.save_state):
                     try:
                         self.logger.debug(f"Saving state for non-active module '{key}'...")
                         module_instance.save_state()
                     except Exception as e:
                         self.logger.error(f"Error saving state for non-active module '{key}': {e}", exc_info=False)

        self.logger.debug("Processing final events before exit...")
        QApplication.processEvents()
        self.logger.info("Application closing")
        event.accept()


# --- Main Execution ---
def main():
    """Main application entry point."""
    QApplication.setApplicationName("BCApp")
    QApplication.setOrganizationName("YourOrganization")
    QApplication.setApplicationVersion("1.0.0")

    config = None
    try:
        config = Config()
        if not hasattr(config, 'base_path') or not config.base_path:
             raise ValueError("Config loaded, but 'base_path' is missing or empty.")
    except Exception as e:
        print(f"CRITICAL: FATAL: Failed to initialize configuration: {e}", file=sys.stderr)
        traceback.print_exc(file=sys.stderr)
        try:
            _app_temp = QApplication.instance() or QApplication(sys.argv)
            QMessageBox.critical(None, "Configuration Error", f"Failed to load configuration:\n{e}\nApplication cannot start.")
        except Exception as qe:
             print(f"CRITICAL: Could not display Qt error message: {qe}", file=sys.stderr)
        sys.exit(1)

    log_level_str = config.get('log_level', 'INFO')
    logger = None
    try:
        log_dir_path = getattr(config, 'log_dir', None)
        if log_dir_path is None:
            print(f"WARNING: config.log_dir not found, attempting default 'logs' directory relative to {config.base_path}.", file=sys.stderr)
            log_dir_path = os.path.join(config.base_path, 'logs')
        logger = setup_logging(log_dir=log_dir_path, log_level_str=log_level_str)
    except Exception as e:
        print(f"CRITICAL: FATAL: Failed to initialize logging: {e}", file=sys.stderr)
        traceback.print_exc(file=sys.stderr)
        try:
            _app_temp = QApplication.instance() or QApplication(sys.argv)
            QMessageBox.critical(None, "Logging Error", f"Failed to initialize logging:\n{e}\nApplication cannot start.")
        except Exception as qe:
             print(f"CRITICAL: Could not display Qt error message: {qe}", file=sys.stderr)
        sys.exit(1)

    logger.info("--- Application Start ---")
    logger.info(f"Version: {QApplication.applicationVersion()}")
    logger.info(f"Base path: {getattr(config, 'base_path', 'N/A')}")
    logger.info(f"Log level set to: {log_level_str}")

    logger.info("Initializing core handlers and managers...")
    cache_handler = None
    csv_handler = None
    thread_pool = None
    oauth_client = None
    jd_auth_manager = None
    sharepoint_manager_instance = None
    quote_integration_instance = None
    maintain_quotes_api = None

    try:
        cache_dir_path = getattr(config, 'cache_dir', None)
        if cache_dir_path:
             cache_handler = CacheHandler(cache_dir=cache_dir_path)
             logger.debug(f"CacheHandler initialized with path: {cache_dir_path}")
        else:
             logger.warning("config.cache_dir not found, CacheHandler not initialized.")

        data_dir_path = getattr(config, 'data_dir', None)
        if data_dir_path:
             csv_handler = CSVHandler(data_path=data_dir_path)
             logger.debug(f"CSVHandler initialized with path: {data_dir_path}")
        else:
             logger.warning("config.data_dir not found, CSVHandler not initialized.")

        thread_pool = QThreadPool.globalInstance()
        logger.debug(f"Using global QThreadPool instance. Max threads: {thread_pool.maxThreadCount()}")

        jd_id = getattr(config, 'jd_client_id', None)
        jd_secret = getattr(config, 'jd_client_secret', None)
        if jd_id and jd_secret:
             jd_cache_path = getattr(config, 'cache_dir', None)
             if jd_cache_path:
                  oauth_client = JohnDeereOAuthClient(client_id=jd_id, client_secret=jd_secret, cache_path=jd_cache_path, logger=logger.getChild("JDAuthClient"))
                  logger.debug("JohnDeereOAuthClient initialized.")
             else:
                  logger.warning("config.cache_dir not found, cannot initialize JohnDeereOAuthClient with caching.")
        else:
             logger.warning("JD Client ID or Secret missing in config, JohnDeereOAuthClient not initialized.")

        jd_auth_manager = JDAuthManager(config=config, logger=logger.getChild("JDAuthMan"))
        logger.debug("JDAuthManager initialized.")

        if SharePointManager:
             sharepoint_manager_instance = SharePointManager(config=config, logger=logger.getChild("SPMan"))
             logger.debug("SharePointManager initialized.")
        else:
             logger.error("SharePointManager class not available, cannot initialize instance.")

        if MaintainQuotesAPI and oauth_client:
            maintain_quotes_api = MaintainQuotesAPI(logger=logger.getChild("MaintainQuotesAPI"))
            token = None
            try:
                token = oauth_client.get_token()
            except Exception as token_error:
                 logger.error(f"Failed to get initial token for MaintainQuotesAPI: {token_error}", exc_info=False)
            if token and hasattr(maintain_quotes_api, 'set_access_token'):
                 maintain_quotes_api.set_access_token(token)
                 logger.debug("MaintainQuotesAPI initialized and token set.")
            elif not token:
                 logger.error("Failed to get token for MaintainQuotesAPI, API calls will likely fail.")
            elif not hasattr(maintain_quotes_api, 'set_access_token'):
                 logger.error("MaintainQuotesAPI instance does not have set_access_token method!")
        elif not MaintainQuotesAPI:
            logger.warning("MaintainQuotesAPI class not imported.")
        elif not oauth_client:
            logger.warning("OAuth client not available for MaintainQuotesAPI.")

        if QuoteIntegration and maintain_quotes_api:
             quote_integration_instance = QuoteIntegration(quotes_api=maintain_quotes_api, sharepoint_manager=sharepoint_manager_instance, logger=logger.getChild("QuoteIntegration"), config=config)
             logger.debug("QuoteIntegration initialized.")
        elif not QuoteIntegration:
             logger.warning("QuoteIntegration class not imported.")
        elif not maintain_quotes_api:
             logger.warning("MaintainQuotesAPI not available, cannot initialize QuoteIntegration.")

        logger.info("Core handlers and managers initialized (or skipped where necessary).")

    except Exception as e:
         logger.critical(f"FATAL: Failed to initialize core managers: {e}", exc_info=True)
         try:
              _app_temp = QApplication.instance() or QApplication(sys.argv)
              QMessageBox.critical(None, "Initialization Error", f"Failed to initialize core components:\n{e}\nApplication cannot start.")
         except Exception as qe:
              print(f"CRITICAL: Could not display Qt error message: {qe}", file=sys.stderr)
         sys.exit(1)

    app = QApplication(sys.argv)

    # Load the QSS stylesheet here
    try:
        qss_path = "C:\\Users\\smorley\\BC\\resources\\styles\\light.qss"
        if os.path.exists(qss_path):
            with open(qss_path, "r") as f:
                app.setStyleSheet(f.read())
            logger.info(f"Applied QSS stylesheet from: {qss_path}")
        else:
            logger.warning(f"QSS file not found: {qss_path}")
    except Exception as e:
        logger.error(f"Failed to load QSS stylesheet: {e}", exc_info=True)

    ui_theme = config.get('ui_theme', 'light')
    logger.info(f"Applying UI theme: {ui_theme}")
    try:
        ThemeManager.apply_theme(ui_theme)
    except Exception as e:
         logger.error(f"Failed to apply theme '{ui_theme}': {e}", exc_info=True)

    splash = None
    if SplashScreen:
        resources_dir_path = getattr(config, 'resources_dir', None)
        splash_image_filename = config.get("splash_image", "splash.png")
        splash_image_path = get_resource_path(splash_image_filename, resources_dir_path)
        logger.debug(f"Attempting to load splash image from: {splash_image_path}")
        try:
            splash_pixmap = QPixmap(splash_image_path)
            if splash_pixmap.isNull():
                 logger.error(f"Splash pixmap is null. Check image format/path: {splash_image_path}")
                 splash = None
            else:
                 splash = SplashScreen(splash_pixmap)
                 splash.show()
                 splash.showMessage("Loading core components...", Qt.AlignBottom | Qt.AlignHCenter, Qt.white)
                 QApplication.processEvents()
                 logger.debug("Splash screen shown.")
        except Exception as e:
             logger.error(f"Failed to load or show splash screen image '{splash_image_path}': {e}")
             splash = None
    else:
        logger.warning("SplashScreen class not available, skipping splash screen.")

    logger.info("Creating MainWindow...")
    window = None
    try:
        window = MainWindow(config=config, logger=logger, cache_handler=cache_handler, csv_handler=csv_handler, thread_pool=thread_pool, oauth_client=oauth_client, jd_auth_manager=jd_auth_manager, sharepoint_manager=sharepoint_manager_instance, quote_integration=quote_integration_instance)
    except Exception as e:
        logger.critical(f"FATAL: Failed to create MainWindow: {e}", exc_info=True)
        if splash: splash.finish(None)
        QMessageBox.critical(None, "Initialization Error", f"Failed to create main window:\n{e}\nApplication cannot start.")
        sys.exit(1)

    def finish_startup():
        """Closes splash, shows window, loads modules, sets initial view."""
        nonlocal window, splash
        try:
            if splash:
                 splash.finish(window)
                 logger.debug("Splash screen finished.")
                 splash = None
            if window:
                 window.show()
                 logger.debug("Main window shown.")
                 window._load_modules_and_init()
            else:
                 logger.critical("Main window was not created successfully. Exiting.")
                 if QApplication.instance(): QApplication.instance().quit()
        except Exception as finish_error:
             logger.critical(f"Error during final startup sequence: {finish_error}", exc_info=True)
             if QApplication.instance(): QApplication.instance().quit()

    if splash:
        splash.showMessage("Launching application...", Qt.AlignBottom | Qt.AlignHCenter, Qt.white)
        QTimer.singleShot(1500, finish_startup)
    else:
        QTimer.singleShot(50, finish_startup)

    sys.excepthook = log_exception_hook
    logger.info("Global exception hook set.")

    logger.info("Starting application event loop.")
    exit_code = 0
    try:
        exit_code = app.exec_()
        logger.info(f"Application exited with code: {exit_code}")
    except Exception as e_exec:
         logger.critical(f"FATAL: Unhandled exception during app.exec_(): {e_exec}", exc_info=True)
         exit_code = 1
    finally:
        logger.info("Exiting application.")
        sys.exit(exit_code)


if __name__ == '__main__':
    main()
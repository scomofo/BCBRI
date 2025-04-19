# File: ui/notification.py
from PyQt5.QtWidgets import QWidget, QLabel, QVBoxLayout, QHBoxLayout, QPushButton
from PyQt5.QtCore import Qt, QTimer, pyqtSignal
from PyQt5.QtGui import QIcon

class Notification(QWidget):
    """Popup notification widget."""
    
    # Signal emitted when notification is closed
    closed = pyqtSignal()
    
    # Notification types
    INFO = "info"
    SUCCESS = "success"
    WARNING = "warning"
    ERROR = "error"
    
    def __init__(self, parent=None, title="", message="", 
                 notification_type=INFO, duration=5000):
        """Initialize notification.
        
        Args:
            parent: Parent widget
            title: Notification title
            message: Notification message
            notification_type: Type of notification (info, success, warning, error)
            duration: Display duration in ms (0 for no auto-close)
        """
        super().__init__(parent)
        
        self.title = title
        self.message = message
        self.notification_type = notification_type
        self.duration = duration
        
        # Set up UI
        self.init_ui()
        
        # Set up timer for auto-close
        if duration > 0:
            self.timer = QTimer(self)
            self.timer.setSingleShot(True)
            self.timer.timeout.connect(self.close_notification)
            self.timer.start(duration)
    
    def init_ui(self):
        """Initialize the user interface."""
        # Set window flags
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        # Main layout
        main_layout = QVBoxLayout(self)
        
        # Header layout
        header_layout = QHBoxLayout()
        
        # Title label
        title_label = QLabel(self.title)
        title_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        
        # Close button
        close_button = QPushButton("×")
        close_button.setFlat(True)
        close_button.setFixedSize(20, 20)
        close_button.clicked.connect(self.close_notification)
        
        # Add to header layout
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        header_layout.addWidget(close_button)
        
        # Message label
        message_label = QLabel(self.message)
        message_label.setWordWrap(True)
        
        # Add to main layout
        main_layout.addLayout(header_layout)
        main_layout.addWidget(message_label)
        
        # Set style based on notification type
        self.set_style()
        
        # Set fixed width
        self.setFixedWidth(300)
    
    def set_style(self):
        """Set widget style based on notification type."""
        base_style = """
            QWidget {
                border-radius: 4px;
                padding: 10px;
            }
        """
        
        type_styles = {
            self.INFO: """
                background-color: #E3F2FD;
                border: 1px solid #2196F3;
            """,
            self.SUCCESS: """
                background-color: #E8F5E9;
                border: 1px solid #4CAF50;
            """,
            self.WARNING: """
                background-color: #FFF8E1;
                border: 1px solid #FFC107;
            """,
            self.ERROR: """
                background-color: #FFEBEE;
                border: 1px solid #F44336;
            """
        }
        
        self.setStyleSheet(base_style + type_styles.get(self.notification_type, type_styles[self.INFO]))
    
    def close_notification(self):
        """Close the notification."""
        self.closed.emit()
        self.deleteLater()
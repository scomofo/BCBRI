# File: ui/loading_widget.py
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QProgressBar
from PyQt5.QtCore import Qt, QTimer

class LoadingWidget(QWidget):
    """Widget to show loading status."""
    
    def __init__(self, parent=None, message="Loading..."):
        super().__init__(parent)
        
        # Set up layout
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        
        # Add message label
        self.message_label = QLabel(message)
        self.message_label.setAlignment(Qt.AlignCenter)
        self.message_label.setStyleSheet("font-size: 14pt;")
        
        # Add progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # Indeterminate mode
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedSize(300, 15)
        
        # Add to layout
        layout.addWidget(self.message_label)
        layout.addWidget(self.progress_bar)
        
        # Set widget properties
        self.setAutoFillBackground(True)
        self.setStyleSheet("background-color: rgba(240, 240, 240, 240);")
        
    def set_message(self, message):
        """Update the loading message.
        
        Args:
            message: The message to display
        """
        self.message_label.setText(message)
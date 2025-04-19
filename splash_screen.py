# File: ui/splash_screen.py
from PyQt5.QtWidgets import QSplashScreen, QProgressBar, QVBoxLayout, QWidget, QApplication
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt, QTimer

class SplashScreen(QSplashScreen):
    def __init__(self, pixmap):
        super().__init__(pixmap)
        
        # Create a progress bar
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setGeometry(0, pixmap.height() - 20, pixmap.width(), 20)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        
    def progress(self, value, message=""):
        """Update the progress bar and display a message.
        
        Args:
            value: Progress value (0-100)
            message: Status message to display
        """
        self.progress_bar.setValue(value)
        self.showMessage(message, Qt.AlignBottom | Qt.AlignHCenter, Qt.white)
        
        # Process events to update display
        QApplication.processEvents()
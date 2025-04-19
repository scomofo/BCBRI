from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QPalette, QColor
from PyQt5.QtCore import Qt

class ThemeManager:
    """Manage application themes."""
    
    THEMES = {
        'light': {
            'window': '#F0F0F0',
            'text': '#202020',
            'accent': '#0078D7',
            'highlight': '#E5F1FB',
            'sidebar': '#E0E0E0',
            'button': '#DDDDDD',
            'button_text': '#202020',
        },
        'dark': {
            'window': '#2D2D30',
            'text': '#FFFFFF',
            'accent': '#007ACC',
            'highlight': '#3E3E42',
            'sidebar': '#252526',
            'button': '#333333',
            'button_text': '#FFFFFF',
        },
        'blue': {
            'window': '#ECF4FF',
            'text': '#333333',
            'accent': '#1E88E5',
            'highlight': '#BBDEFB',
            'sidebar': '#DCEDC8',
            'button': '#90CAF9',
            'button_text': '#333333',
        }
    }
    
    @classmethod
    def apply_theme(cls, theme_name):
        """Apply a theme to the application.
        
        Args:
            theme_name: Name of the theme to apply
        """
        # Get theme colors
        theme = cls.THEMES.get(theme_name, cls.THEMES['light'])
        
        # Create a palette
        palette = QPalette()
        
        # Set colors based on theme
        palette.setColor(QPalette.Window, QColor(theme['window']))
        palette.setColor(QPalette.WindowText, QColor(theme['text']))
        palette.setColor(QPalette.Base, QColor(theme['window']))
        palette.setColor(QPalette.AlternateBase, QColor(theme['highlight']))
        palette.setColor(QPalette.ToolTipBase, QColor(theme['window']))
        palette.setColor(QPalette.ToolTipText, QColor(theme['text']))
        palette.setColor(QPalette.Text, QColor(theme['text']))
        palette.setColor(QPalette.Button, QColor(theme['button']))
        palette.setColor(QPalette.ButtonText, QColor(theme['button_text']))
        palette.setColor(QPalette.BrightText, Qt.white)
        palette.setColor(QPalette.Link, QColor(theme['accent']))
        palette.setColor(QPalette.Highlight, QColor(theme['accent']))
        palette.setColor(QPalette.HighlightedText, Qt.white)
        
        # Apply palette to application
        QApplication.setPalette(palette)
        
        # Return the theme dictionary for other custom styling
        return theme
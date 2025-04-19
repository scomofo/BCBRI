# File: ui/base_module.py
from PyQt5.QtWidgets import QWidget

class BaseModule(QWidget):
    """Base class for all application modules."""
    
    def __init__(self, main_window=None):
        """Initialize the base module.
        
        Args:
            main_window: Reference to the main application window
        """
        super().__init__()
        self.main_window = main_window
            
    def init_ui(self):
        """Initialize the user interface.
        
        To be implemented by child classes.
        """
        pass
    
    def close(self):
        """Perform cleanup operations.
        
        To be extended by child classes if needed.
        """
        pass
        
    def get_title(self):
        """Return the module title for display purposes.
        
        Returns:
            str: The title of the module
        """
        return self.__class__.__name__.replace('Module', '')
        
    def refresh(self):
        """Refresh module data.
        
        To be implemented by child classes.
        """
        pass
        
    def search(self, search_text):
        """Search for content within this module.
        
        Args:
            search_text: Text to search for
            
        Returns:
            List of search results or empty list if not implemented
        """
        return []
        
    def navigate_to(self, result):
        """Navigate to a specific item within this module.
        
        Args:
            result: Search result data
        """
        pass
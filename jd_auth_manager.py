import os
import logging
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                             QPushButton, QFormLayout, QMessageBox, QApplication)
from PyQt5.QtCore import Qt

from .oauth_client import JohnDeereOAuthClient

class JDAuthManager:
    """Manager for John Deere API authentication."""
    
    def __init__(self, config=None, logger=None):
        """Initialize the auth manager.
        
        Args:
            config: Application configuration
            logger: Logger instance
        """
        self.config = config
        self.logger = logger or logging.getLogger(__name__)
        self.cache_path = getattr(config, 'cache_path', None) if config else None
        
        # Get credentials from environment or config
        self.client_id = os.getenv('JD_CLIENT_ID', '')
        self.client_secret = os.getenv('DEERE_CLIENT_SECRET', '')
        
        # Override with config if available
        if config:
            if hasattr(config, 'jd_client_id') and config.jd_client_id:
                self.client_id = config.jd_client_id
            if hasattr(config, 'jd_client_secret') and config.jd_client_secret:
                self.client_secret = config.jd_client_secret
        
        # Initialize the OAuth client if we have credentials
        self.oauth_client = None
        if self.client_id and self.client_secret:
            self.oauth_client = JohnDeereOAuthClient(
                client_id=self.client_id,
                client_secret=self.client_secret,
                cache_path=self.cache_path,
                logger=self.logger
            )
    
    def get_access_token(self, force_refresh=False):
        """Get an access token for John Deere API.
        
        Args:
            force_refresh: Force getting a new token even if cached one exists
            
        Returns:
            Access token string or None if failed
        """
        # Check if we have an OAuth client
        if not self.oauth_client:
            # Try to create one if we have credentials
            if self.client_id and self.client_secret:
                self.oauth_client = JohnDeereOAuthClient(
                    client_id=self.client_id,
                    client_secret=self.client_secret,
                    cache_path=self.cache_path,
                    logger=self.logger
                )
            else:
                # Show dialog to get credentials
                creds = self._show_credentials_dialog()
                if creds:
                    self.client_id = creds.get('client_id', '')
                    self.client_secret = creds.get('client_secret', '')
                    
                    # Save to config if available
                    if self.config:
                        self.config.jd_client_id = self.client_id
                        self.config.jd_client_secret = self.client_secret
                        if hasattr(self.config, 'save'):
                            self.config.save()
                    
                    # Create the OAuth client
                    self.oauth_client = JohnDeereOAuthClient(
                        client_id=self.client_id,
                        client_secret=self.client_secret,
                        cache_path=self.cache_path,
                        logger=self.logger
                    )
                else:
                    self.logger.error("No credentials provided")
                    return None
        
        # Get the token
        try:
            token = self.oauth_client.get_token(force_refresh=force_refresh)
            return token
        except Exception as e:
            self.logger.error(f"Error getting access token: {str(e)}")
            return None
    
    def _show_credentials_dialog(self):
        """Show dialog to get client credentials.
        
        Returns:
            dict with client_id and client_secret or None if canceled
        """
        dialog = QDialog()
        dialog.setWindowTitle("John Deere API Credentials")
        dialog.resize(500, 200)
        
        layout = QVBoxLayout(dialog)
        
        # Instructions
        instructions = QLabel(
            "Please enter your John Deere API credentials.\n"
            "These can be found in your Developer.Deere.com application settings."
        )
        instructions.setWordWrap(True)
        layout.addWidget(instructions)
        
        # Form for credentials
        form_layout = QFormLayout()
        
        # Client ID
        client_id_edit = QLineEdit()
        client_id_edit.setText(self.client_id)
        form_layout.addRow("Client ID:", client_id_edit)
        
        # Client Secret
        client_secret_edit = QLineEdit()
        client_secret_edit.setText(self.client_secret)
        client_secret_edit.setEchoMode(QLineEdit.Password)
        form_layout.addRow("Client Secret:", client_secret_edit)
        
        layout.addLayout(form_layout)
        
        # Buttons
        button_layout = QHBoxLayout()
        
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(dialog.reject)
        
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(dialog.accept)
        ok_button.setDefault(True)
        
        button_layout.addWidget(cancel_button)
        button_layout.addWidget(ok_button)
        
        layout.addLayout(button_layout)
        
        # Execute the dialog
        if dialog.exec_() == QDialog.Accepted:
            return {
                'client_id': client_id_edit.text().strip(),
                'client_secret': client_secret_edit.text().strip()
            }
        
        return None
    
    def setup_api_client(self, api_client):
        """Set up an API client with authentication.
        
        Args:
            api_client: API client instance to set up
            
        Returns:
            True if successful, False otherwise
        """
        if not api_client:
            self.logger.error("No API client provided")
            return False
        
        # Get an access token
        token = self.get_access_token()
        if not token:
            self.logger.error("Failed to get access token")
            return False
        
        # Set the token on the API client
        if hasattr(api_client, 'set_access_token'):
            api_client.set_access_token(token)
            return True
        else:
            self.logger.error("API client does not have set_access_token method")
            return False
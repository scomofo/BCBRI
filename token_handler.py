# utils/token_handler.py - Update token handling logic

import os
import json
import time
import logging
import requests
from datetime import datetime, timedelta

class TokenHandler:
    """Handler for OAuth token management."""
    
    def __init__(self, cache_path=None, logger=None):
        """Initialize the token handler.
        
        Args:
            cache_path: Path to cache directory for storing tokens
            logger: Logger instance
        """
        self.cache_path = cache_path
        self.logger = logger or logging.getLogger(__name__)
        
        # Ensure cache directory exists
        if self.cache_path:
            os.makedirs(self.cache_path, exist_ok=True)
            self.token_file = os.path.join(self.cache_path, "jd_token.json")
        else:
            self.token_file = None
    
    def clean_token(self, token):
        """Clean a token string to remove any whitespace or newlines.
        
        Args:
            token: Raw token string
            
        Returns:
            Cleaned token string
        """
        if not isinstance(token, str):
            return None
            
        # Remove any whitespace, newlines, etc.
        cleaned = token.strip()
        
        # Check if it looks like a valid token (non-empty after cleaning)
        if not cleaned:
            self.logger.warning("Token is empty after cleaning")
            return None
            
        return cleaned
    
    def save_token(self, token, expires_in=43200):
        """Save a token to the cache.
        
        Args:
            token: Token string
            expires_in: Seconds until token expiration (default: 12 hours)
            
        Returns:
            True if saved successfully, False otherwise
        """
        if not self.token_file:
            self.logger.error("No cache path configured for token storage")
            return False
            
        try:
            # Clean the token
            cleaned_token = self.clean_token(token)
            if not cleaned_token:
                self.logger.error("Invalid token format, not saving")
                return False
                
            # Create cache directory if it doesn't exist
            os.makedirs(os.path.dirname(self.token_file), exist_ok=True)
            
            # Calculate expiry time
            expires_at = time.time() + expires_in
            
            # Create token data
            token_data = {
                "access_token": cleaned_token,
                "expires_in": expires_in,
                "expires_at": expires_at,
                "saved_at": time.time()
            }
            
            # Write to file
            with open(self.token_file, 'w') as f:
                json.dump(token_data, f)
                
            self.logger.info(f"Token saved to {self.token_file}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error saving token: {str(e)}")
            return False
    
    def load_token(self):
        """Load a token from the cache.
        
        Returns:
            Token string if valid, None otherwise
        """
        if not self.token_file or not os.path.exists(self.token_file):
            self.logger.info("No cached token file found")
            return None
            
        try:
            with open(self.token_file, 'r') as f:
                token_data = json.load(f)
                
            # Check if token has expired
            if 'expires_at' not in token_data or token_data['expires_at'] <= time.time() + 300:  # 5 minute buffer
                self.logger.info("Cached token has expired or is about to expire")
                return None
                
            # Get the token
            if 'access_token' not in token_data:
                self.logger.error("Cached token data is invalid (missing access_token)")
                return None
                
            # Clean and validate the token
            cleaned_token = self.clean_token(token_data['access_token'])
            if not cleaned_token:
                self.logger.error("Cached token is invalid")
                return None
                
            self.logger.info("Using cached token")
            return cleaned_token
            
        except Exception as e:
            self.logger.error(f"Error loading token: {str(e)}")
            return None
    
    def test_token(self, token, api_base_url=None):
        """Test if a token is valid by making a request to the API.
        
        Args:
            token: Token to test
            api_base_url: Base URL for the API
            
        Returns:
            (success, message) tuple
        """
        if not token:
            return False, "No token provided"
            
        # Clean the token
        cleaned_token = self.clean_token(token)
        if not cleaned_token:
            return False, "Invalid token format"
            
        # Default API base URL if not provided
        if not api_base_url:
            api_base_url = "https://jdquote2-api-sandbox.deere.com/om/cert/maintainquote"
            
        # Make a simple request to test the token
        test_url = f"{api_base_url}/api/v1/dealers/X000000/maintain-quotes"
        
        headers = {
            "Authorization": f"Bearer {cleaned_token}",
            "Accept": "application/json",
            "Content-Type": "application/json"
        }
        
        data = {"dealerRacfID": "X000000"}
        
        try:
            self.logger.info(f"Testing token against {test_url}")
            response = requests.post(test_url, headers=headers, json=data, timeout=15)
            
            # Even 401/404 responses can indicate the token format is valid but permissions/paths are wrong
            # So we consider this a valid test unless it's a connection error
            if response.status_code == 200:
                self.logger.info("Token test successful")
                return True, "Token is valid"
            elif response.status_code == 401:
                self.logger.warning(f"Token authentication failed: {response.text}")
                return False, f"Authentication failed: {response.text}"
            else:
                self.logger.warning(f"Token test failed with status {response.status_code}: {response.text}")
                # Still return True for format validity if we get any response
                return True, f"Token format is valid, but API returned: {response.status_code}"
                
        except requests.exceptions.RequestException as e:
            self.logger.error(f"Request error during token test: {str(e)}")
            return False, f"Connection error: {str(e)}"
        except Exception as e:
            self.logger.error(f"Error testing token: {str(e)}")
            return False, f"Error: {str(e)}"
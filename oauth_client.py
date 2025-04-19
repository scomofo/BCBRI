# utils/oauth_client.py - Consolidated version
import requests
import base64
import json
import time
import logging
import os
import traceback

class JohnDeereOAuthClient:
    """Client for handling John Deere OAuth authentication with consolidated token handling."""
    
    # OAuth endpoints from documentation
    TOKEN_URL = "https://signin.johndeere.com/oauth2/aus78tnlaysMraFhC1t7/v1/token"
    WELL_KNOWN_URL = "https://signin.johndeere.com/oauth2/aus78tnlaysMraFhC1t7/.well-known/oauth-authorization-server"
    
    def __init__(self, client_id, client_secret, cache_path=None, logger=None):
        """Initialize the OAuth client.
        
        Args:
            client_id: Application Client ID
            client_secret: Application Client Secret
            cache_path: Path to cache directory for storing tokens
            logger: Logger instance
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.cache_path = cache_path
        self.logger = logger or logging.getLogger(__name__)
        
        # Ensure cache directory exists
        if self.cache_path:
            os.makedirs(self.cache_path, exist_ok=True)
            self.token_file = os.path.join(self.cache_path, "jd_token.json")
        else:
            self.token_file = None
            
        # Log initialization
        self.logger.info(f"Initialized JohnDeereOAuthClient with client_id: {client_id[:5]}...")
    
    def get_oauth_endpoints(self):
        """Get OAuth endpoints from the well-known URL."""
        try:
            self.logger.info(f"Fetching OAuth endpoints from {self.WELL_KNOWN_URL}")
            response = requests.get(self.WELL_KNOWN_URL, timeout=30)
            if response.status_code == 200:
                data = response.json()
                endpoints = {
                    'token_endpoint': data.get('token_endpoint', self.TOKEN_URL),
                    'authorization_endpoint': data.get('authorization_endpoint'),
                    'scopes_supported': data.get('scopes_supported', [])
                }
                self.logger.info(f"Fetched endpoints: {endpoints['token_endpoint']}")
                return endpoints
            else:
                self.logger.error(f"Failed to get OAuth endpoints: {response.status_code} - {response.text}")
                return {
                    'token_endpoint': self.TOKEN_URL
                }
        except Exception as e:
            self.logger.error(f"Error getting OAuth endpoints: {str(e)}")
            self.logger.error(traceback.format_exc())
            return {
                'token_endpoint': self.TOKEN_URL
            }
    
    def get_client_credentials_token(self, scope="offline_access"):
        """Get a token using the client credentials grant type.
        
        Args:
            scope: OAuth scope to request
            
        Returns:
            dict with token information or None if failed
        """
        # Get token endpoint URL
        endpoints = self.get_oauth_endpoints()
        token_url = endpoints.get('token_endpoint', self.TOKEN_URL)
        
        # Create authorization header
        auth_string = f"{self.client_id}:{self.client_secret}"
        auth_bytes = auth_string.encode('ascii')
        auth_b64 = base64.b64encode(auth_bytes).decode('ascii')
        
        # Set up headers and data
        headers = {
            "Authorization": f"Basic {auth_b64}",
            "Accept": "application/json",
            "Content-Type": "application/x-www-form-urlencoded"
        }
        
        data = {
            "grant_type": "client_credentials",
            "scope": scope
        }
        
        try:
            self.logger.info(f"Requesting access token from {token_url}")
            response = requests.post(token_url, headers=headers, data=data, timeout=30)
            
            # Log response details for debugging
            self.logger.debug(f"Token response status: {response.status_code}")
            if hasattr(response, 'text'):
                self.logger.debug(f"Token response content: {response.text[:100]}...")
            
            if response.status_code == 200:
                token_data = response.json()
                # Add expiry timestamp
                if 'expires_in' in token_data:
                    token_data['expires_at'] = time.time() + token_data['expires_in']
                
                self.logger.info(f"Successfully obtained access token (expires in {token_data.get('expires_in', 'unknown')} seconds)")
                
                # Save token to cache
                if self.token_file:
                    try:
                        with open(self.token_file, 'w') as f:
                            json.dump(token_data, f)
                        self.logger.info(f"Saved token to {self.token_file}")
                    except Exception as e:
                        self.logger.error(f"Error saving token to cache: {str(e)}")
                
                return token_data
            else:
                self.logger.error(f"Failed to get token: {response.status_code} - {response.text}")
                return None
        except Exception as e:
            self.logger.error(f"Error getting token: {str(e)}")
            self.logger.error(traceback.format_exc())
            return None
    
    def load_cached_token(self):
        """Load a token from cache if available and not expired.
        
        Returns:
            Token string or None if no valid cached token
        """
        if not self.token_file or not os.path.exists(self.token_file):
            self.logger.info(f"No cached token file found at {self.token_file}")
            return None
        
        try:
            with open(self.token_file, 'r') as f:
                token_data = json.load(f)
            
            # Check if token has expired (with 5-minute buffer)
            if 'expires_at' in token_data and token_data['expires_at'] > time.time() + 300:
                self.logger.info("Using cached token that expires at " + 
                               time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(token_data['expires_at'])))
                return token_data.get('access_token')
            else:
                self.logger.info("Cached token has expired or will expire soon")
                return None
        except Exception as e:
            self.logger.error(f"Error loading cached token: {str(e)}")
            return None
    
    def get_token(self, force_refresh=False):
        """Get a valid token, using cache if available.
        
        Args:
            force_refresh: Force getting a new token even if cached one exists
            
        Returns:
            Token string or None if failed
        """
        if not force_refresh:
            # Try to load from cache first
            cached_token = self.load_cached_token()
            if cached_token:
                return cached_token
        
        # Get a new token
        self.logger.info("Getting new token using client credentials flow")
        token_data = self.get_client_credentials_token()
        if token_data:
            return token_data.get('access_token')
        
        return None
    
    # Methods moved from token_handler.py
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
# File: utils/config.py - CORRECTED
import os
import json
import logging # Import logging
from dotenv import load_dotenv

# It's better practice to get the logger instance here if methods use it
logger = logging.getLogger(__name__)

class Config:
    """Application configuration manager."""

    # Default configuration
    DEFAULTS = {
        'weather_refresh_interval': 1,  # hours
        'exchange_refresh_interval': 6,  # hours
        'commodities_refresh_interval': 4,  # hours
        'api_timeout': 15,  # seconds
        'ui_theme': 'light',
        'enable_high_dpi': True,
        'log_level': 'INFO',
        'splash_image': 'splash.png', # Add default splash image name
        'app_icon': 'app_icon.png',   # Add default app icon name
        'app_title': 'BRIDeal', # Add default app title
        'window_width': 1200,
        'window_height': 800,
        'toolbar_icon_size': 24,
        # Add defaults for traffic auto if needed
        # 'traffic_images_dir_name': 'traffic_images', # Example: Subdirectory name in resources
        # 'traffic_csv_filename': 'traffic_tasks.csv', # Example: Filename in data dir
        # 'pyautogui_pause': 0.2,
        # 'pyautogui_timeout': 10,
        # 'pyautogui_confidence': 0.8,
    }

    def __init__(self, base_path=None):
        """Initialize the configuration manager."""
        # Load environment variables first
        try:
            dotenv_path = os.path.expanduser("~/.env")
            loaded = load_dotenv(dotenv_path)
            # Use print here as logger might not be configured yet
            print(f"DEBUG: Attempted to load .env from '{dotenv_path}'. Loaded: {loaded}")
        except Exception as e:
            print(f"WARNING: Error loading .env file from {dotenv_path}: {e}")

        # Set base path
        self.base_path = base_path or os.getcwd()
        print(f"DEBUG: Config using base path: {self.base_path}") # Keep this debug print

        # Initialize with default values
        self.config = self.DEFAULTS.copy()

        # Try to load from config file (config.json)
        self._load_from_file() # Handles its own errors

        # Create derived paths (ensure this runs without error)
        try:
            print("DEBUG: Calling _setup_paths()") # Add this debug print
            self._setup_paths()
            print("DEBUG: _setup_paths() finished") # Add this debug print
        except Exception as e:
            print(f"CRITICAL ERROR during _setup_paths: {e}")
            # Decide how to handle this - maybe raise e? Or set paths to None?
            # For now, critical attributes might be missing.
            raise # Re-raise the exception to make it clear setup failed

        # Load API credentials
        try:
            print("DEBUG: Calling _load_credentials()") # Add this debug print
            self._load_credentials()
            print("DEBUG: _load_credentials() finished") # Add this debug print
        except Exception as e:
             print(f"CRITICAL ERROR during _load_credentials: {e}")
             raise # Re-raise

        # Log initialized configuration (uses print as logging setup happens later)
        # This should only run if everything above succeeded
        self._log_config()

    def _load_from_file(self):
        """Load configuration from JSON file."""
        config_file = os.path.join(self.base_path, 'config.json')
        print(f"DEBUG: Checking for config file at: {config_file}")
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r') as f:
                    loaded_config = json.load(f)
                self.config.update(loaded_config)
                print(f"INFO: Loaded configuration from {config_file}")
            except Exception as e:
                print(f"ERROR: Error loading config file {config_file}: {str(e)}")
        else:
            print(f"DEBUG: Config file {config_file} not found, using defaults and environment variables.")


    def _setup_paths(self):
        """Setup application paths using consistent _dir suffix."""
        print(f"DEBUG: Setting up paths based on base_path: {self.base_path}")
        # Main directories
        self.data_dir = os.path.join(self.base_path, 'data')
        self.log_dir = os.path.join(self.base_path, 'logs')
        self.cache_dir = os.path.join(self.base_path, 'cache')
        self.assets_dir = os.path.join(self.base_path, 'assets')
        self.resources_dir = os.path.join(self.base_path, 'resources')

        # Create modules/exports directory for exports
        self.exports_dir = os.path.join(self.base_path, 'modules', 'exports')
        print(f"DEBUG: Calculated paths - data: {self.data_dir}, log: {self.log_dir}, cache: {self.cache_dir}, assets: {self.assets_dir}, resources: {self.resources_dir}, exports: {self.exports_dir}")

        # Ensure directories exist
        paths_to_check = [self.data_dir, self.log_dir, self.cache_dir,
                          self.assets_dir, self.resources_dir, self.exports_dir]
        print(f"DEBUG: Ensuring directories exist: {paths_to_check}")
        for path in paths_to_check:
            try:
                os.makedirs(path, exist_ok=True)
            except OSError as e:
                print(f"ERROR: Failed to create directory {path}: {e}")
                # Consider raising an error or logging more severely if dir creation is critical
        print("DEBUG: Directory existence check complete.")


    def _load_credentials(self):
        """Load API credentials from environment variables."""
        # Note: load_dotenv was already called in __init__

        # SharePoint/Azure credentials - Use keys matching your .env file
        print("DEBUG: Loading Azure/SharePoint credentials...")
        self.azure_client_id = os.getenv('AZURE_CLIENT_ID')
        self.azure_client_secret = os.getenv('AZURE_CLIENT_SECRET')
        self.azure_tenant_id = os.getenv('AZURE_TENANT_ID')
        self.sharepoint_site_id = os.getenv('SHAREPOINT_SITE_ID')
        self.sharepoint_site_name = os.getenv('SHAREPOINT_SITE_NAME')
        self.sharepoint_file_path = os.getenv('FILE_PATH')
        print(f"DEBUG: Azure Client ID loaded: {bool(self.azure_client_id)}")

        # John Deere Credentials
        print("DEBUG: Loading John Deere credentials...")
        self.jd_client_id = os.getenv('JD_CLIENT_ID')
        self.jd_client_secret = os.getenv('DEERE_CLIENT_SECRET') # Using the name found previously
        print(f"DEBUG: JD Client ID loaded: {bool(self.jd_client_id)}")

        # Other API keys
        print("DEBUG: Loading other API keys...")
        self.finnhub_api_key = os.getenv('FINNHUB_API_KEY')
        print(f"DEBUG: Finnhub API Key loaded: {bool(self.finnhub_api_key)}")

        # Check for required credentials
        self._check_required_credentials() # Logs warnings if missing


    def _check_required_credentials(self):
        """Check if required credentials are present."""
        # Define required variables based on expected environment variable names
        required_vars = [
            'AZURE_CLIENT_ID',
            'AZURE_CLIENT_SECRET',
            'AZURE_TENANT_ID',
            'SHAREPOINT_SITE_ID',
            'SHAREPOINT_SITE_NAME',
            'FILE_PATH',
            'JD_CLIENT_ID',          # Added check
            'DEERE_CLIENT_SECRET',   # Added check
            # Add 'FINNHUB_API_KEY' if it's strictly required
        ]

        missing_vars = []
        for var_name in required_vars:
            # Check if the corresponding attribute on self is None or empty
            # Attribute names are derived by lowercasing the env var name
            attribute_name = var_name.lower()
            if not getattr(self, attribute_name, None):
                # Check if the environment variable itself was missing
                if not os.getenv(var_name):
                    missing_vars.append(var_name)

        if missing_vars:
            # Use print as logger might not be ready
            print(f"WARNING: Missing required environment variables: {', '.join(missing_vars)}")
        else:
            print("DEBUG: All checked required environment variables seem present.")


    def _log_config(self):
        """Log configuration to console (use print as logging not set up yet)."""
        print("--- Configuration Summary ---")
        print(f"Base Directory: {getattr(self, 'base_path', 'N/A')}") # Use getattr for safety
        print(f"Data Directory: {getattr(self, 'data_dir', 'N/A')}")
        print(f"Log Directory: {getattr(self, 'log_dir', 'N/A')}")
        print(f"Cache Directory: {getattr(self, 'cache_dir', 'N/A')}")
        print(f"Assets Directory: {getattr(self, 'assets_dir', 'N/A')}")
        print(f"Resources Directory: {getattr(self, 'resources_dir', 'N/A')}")
        print(f"Exports Directory: {getattr(self, 'exports_dir', 'N/A')}")

        # Log API credentials (masked for security)
        # Use getattr for safety in case they weren't loaded
        jd_client_id = getattr(self, 'jd_client_id', None)
        if jd_client_id:
            print(f"JD Client ID: {jd_client_id}") # Or mask if sensitive

        jd_secret = getattr(self, 'jd_client_secret', None)
        if jd_secret:
            masked_jd_secret = '*' * (len(jd_secret) - 4) + jd_secret[-4:] if len(jd_secret) > 4 else "****"
            print(f"JD Client Secret: {masked_jd_secret}")

        azure_id = getattr(self, 'azure_client_id', None)
        if azure_id:
            print(f"Azure Client ID: {azure_id}")

        azure_secret = getattr(self, 'azure_client_secret', None)
        if azure_secret:
            masked_secret = '*' * (len(azure_secret) - 4) + azure_secret[-4:] if len(azure_secret) > 4 else "****"
            print(f"Azure Client Secret: {masked_secret}")

        azure_tenant = getattr(self, 'azure_tenant_id', None)
        if azure_tenant:
            print(f"Azure Tenant ID: {azure_tenant}")

        sp_site_id = getattr(self, 'sharepoint_site_id', None)
        if sp_site_id:
            print(f"SharePoint Site ID: {sp_site_id}")

        sp_site_name = getattr(self, 'sharepoint_site_name', None)
        if sp_site_name:
            print(f"SharePoint Site Name: {sp_site_name}")

        sp_file_path = getattr(self, 'sharepoint_file_path', None)
        if sp_file_path:
            print(f"SharePoint File Path: {sp_file_path}")

        finnhub_key = getattr(self, 'finnhub_api_key', None)
        if finnhub_key:
            masked_key = finnhub_key[:4] + '*' * (len(finnhub_key) - 8) + finnhub_key[-4:] if len(finnhub_key) > 8 else "****"
            print(f"Finnhub API Key: {masked_key}")


        # Log other configuration values from self.config dictionary
        print(f"Log Level: {self.config.get('log_level')}")
        print(f"UI Theme: {self.config.get('ui_theme')}")
        print(f"Enable High DPI: {self.config.get('enable_high_dpi')}")
        print(f"Splash Image: {self.config.get('splash_image')}")
        print(f"App Icon: {self.config.get('app_icon')}")
        print(f"App Title: {self.config.get('app_title')}")
        print(f"Weather Refresh Interval (Hours): {self.config.get('weather_refresh_interval')}")
        print(f"Exchange Refresh Interval (Hours): {self.config.get('exchange_refresh_interval')}")
        print(f"Commodities Refresh Interval (Hours): {self.config.get('commodities_refresh_interval')}")
        print(f"API Timeout (Seconds): {self.config.get('api_timeout')}")
        print("--- End Configuration Summary ---")

    def get(self, key, default=None):
        """Get a configuration value.

        Args:
            key: The configuration key
            default: Default value if key not found

        Returns:
            The configuration value
        """
        # Check environment variables first (convention: uppercase)
        # Allows overriding config file/defaults via environment
        env_value = os.getenv(key.upper())
        if env_value is not None:
             # Basic type casting based on default type if available
             default_type = type(self.config.get(key))
             if default_type is bool:
                  return env_value.lower() in ('true', '1', 'yes', 'y')
             elif default_type is int:
                  try: return int(env_value)
                  except ValueError: pass # Fallback to string or config value
             elif default_type is float:
                  try: return float(env_value)
                  except ValueError: pass # Fallback to string or config value
             return env_value # Return as string if no type match or unknown

        # Fallback to loaded config (from file or defaults)
        return self.config.get(key, default)

    def save(self):
        """Save the current configuration dictionary (excluding env vars and defaults not overridden) to config.json."""
        config_file = os.path.join(self.base_path, 'config.json')
        config_to_save = {k: v for k, v in self.config.items() if k not in self.DEFAULTS or v != self.DEFAULTS[k]}

        try:
            with open(config_file, 'w') as f:
                json.dump(config_to_save, f, indent=2)
            print(f"INFO: Configuration saved to {config_file}") # Use print as logger might not be ready
        except Exception as e:
            print(f"ERROR: Error saving configuration to {config_file}: {str(e)}")
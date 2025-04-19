# modules/sharepoint_manager.py - V5 - Change default sheet & Add Search Fallback
import os
import json
import time
import logging
import requests
from datetime import datetime, timedelta
import traceback
import msal # Microsoft Authentication Library
import threading
import csv
import io
import re # Added re import
from typing import List, Dict, Any, Optional, Union


# --- Cache Handling (Simple file-based cache) ---
class SimpleCache:
    def __init__(self, cache_dir='cache', default_ttl_seconds=3600):
        self.cache_dir = cache_dir
        self.default_ttl = timedelta(seconds=default_ttl_seconds)
        os.makedirs(self.cache_dir, exist_ok=True)
        self.logger = logging.getLogger(__name__).getChild("Cache")

    def _get_cache_filepath(self, key):
        safe_key = "".join(c if c.isalnum() or c in ('-', '_') else '_' for c in key)
        return os.path.join(self.cache_dir, f"{safe_key}.cache.json")

    def get(self, key):
        filepath = self._get_cache_filepath(key)
        if not os.path.exists(filepath):
            self.logger.debug(f"Cache miss for key '{key}' (file not found)")
            return None
        try:
            file_mod_time = datetime.fromtimestamp(os.path.getmtime(filepath))
            if datetime.now() - file_mod_time > self.default_ttl:
                self.logger.info(f"Cache expired for key '{key}'")
                os.remove(filepath)
                return None
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
                self.logger.info(f"Cache hit for key '{key}'")
                return data
        except (IOError, json.JSONDecodeError, Exception) as e:
            self.logger.error(f"Error reading cache file '{filepath}' for key '{key}': {e}")
            try:
                if os.path.exists(filepath): os.remove(filepath)
            except OSError as rm_err:
                self.logger.error(f"Failed to remove corrupted cache file '{filepath}': {rm_err}")
            return None

    def set(self, key, value):
        filepath = self._get_cache_filepath(key)
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(value, f, indent=4)
            self.logger.info(f"Cache set for key '{key}'")
        except (IOError, Exception) as e:
            self.logger.error(f"Error writing cache file '{filepath}' for key '{key}': {e}")

# --- Main SharePoint Manager Class ---
class SharePointManager:
    """Manages SharePoint interactions for the application."""

    GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0'
    DEFAULT_SCOPES = ['https://graph.microsoft.com/.default']

    def __init__(self, config=None, logger=None):
        """Initialize the SharePoint manager."""
        self.logger = logger or logging.getLogger(__name__).getChild("SPMan")
        self.config = config
        self.tenant_id = None
        self.client_id = None
        self.client_secret = None
        self.site_id = None
        self.site_name = None
        self.file_path = None # This is the PREFERRED path used first
        self.target_filename = None # Extracted filename for searching
        self.cache_dir = None

        if config:
            self.tenant_id = getattr(config, 'azure_tenant_id', None)
            self.client_id = getattr(config, 'azure_client_id', None)
            self.client_secret = getattr(config, 'azure_client_secret', None)
            self.site_id = getattr(config, 'sharepoint_site_id', None)
            self.site_name = getattr(config, 'sharepoint_site_name', None)
            self.file_path = getattr(config, 'sharepoint_file_path', None)
            self.cache_dir = getattr(config, 'cache_dir', 'cache')
            self.logger.debug("Loaded SharePoint config from Config object.")
        else:
            # Load from environment variables as fallback
            self.tenant_id = os.getenv('AZURE_TENANT_ID')
            self.client_id = os.getenv('AZURE_CLIENT_ID')
            self.client_secret = os.getenv('AZURE_CLIENT_SECRET')
            self.site_id = os.getenv('SHAREPOINT_SITE_ID')
            self.site_name = os.getenv('SHAREPOINT_SITE_NAME')
            self.file_path = os.getenv('FILE_PATH') # Uses FILE_PATH from .env
            self.cache_dir = os.getenv('CACHE_DIR', 'cache')
            self.logger.debug("Loaded SharePoint config from environment variables.")

        missing_configs = []
        if not self.tenant_id: missing_configs.append("Azure Tenant ID")
        if not self.client_id: missing_configs.append("Azure Client ID")
        if not self.client_secret: missing_configs.append("Azure Client Secret")
        if not self.site_id: missing_configs.append("SharePoint Site ID")
        if not self.file_path: missing_configs.append("SharePoint File Path")

        if missing_configs:
             msg = f"Missing required SharePoint configuration: {', '.join(missing_configs)}. SharePoint features will likely fail."
             self.logger.error(msg)
        else:
            self.logger.info("SharePoint configuration loaded successfully.")
            self.logger.debug(f"Site ID: {self.site_id}, Preferred File Path: {self.file_path}")
            # Extract filename from path for searching
            if self.file_path:
                 self.target_filename = os.path.basename(self.file_path.replace('\\', '/')) # Handle windows/unix separators
                 self.logger.debug(f"Target filename for search: {self.target_filename}")

        self.msal_app = self._init_msal_app()
        self.token_cache_file = os.path.join(self.cache_dir or '.', "sharepoint_token.dat")
        self.cache = SimpleCache(cache_dir=self.cache_dir)
        self.access_token = None
        self.token_expiry = None
        self._drive_id = None
        self._file_item_id = None
        self._file_path_used = None # Store the path that actually found the file

        # Add a lock for thread-safe token acquisition/cache access if needed
        # self._token_lock = threading.Lock()

        # Call diagnostic function on init
        if not missing_configs:
             self.list_drive_root_children()

        self.logger.info("SharePoint Manager initialized")

    # ...(rest of __init__ methods like _init_msal_app, _load_token_from_cache etc remain the same)...
    def _init_msal_app(self):
        """Initialize the MSAL confidential client application."""
        if not self.tenant_id or not self.client_id or not self.client_secret:
            self.logger.error("Cannot initialize MSAL app: Missing Tenant ID, Client ID, or Client Secret.")
            return None
        authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        try:
            app = msal.ConfidentialClientApplication(
                client_id=self.client_id,
                authority=authority,
                client_credential=self.client_secret,
            )
            self.logger.debug("MSAL Confidential Client Application initialized.")
            return app
        except Exception as e:
            self.logger.error(f"Failed to initialize MSAL application: {e}", exc_info=True)
            return None

    def _load_token_from_cache(self):
        """Loads token from a simple file cache if it exists and hasn't expired."""
        if not os.path.exists(self.token_cache_file):
            return None
        try:
            # with self._token_lock: # Optional locking
            with open(self.token_cache_file, 'r') as f:
                data = json.load(f)
            expiry_time = datetime.fromisoformat(data.get("expires_on"))
            if expiry_time > datetime.now() + timedelta(minutes=5):
                self.logger.info("Loaded valid token from cache.")
                self.access_token = data.get("access_token")
                self.token_expiry = expiry_time
                return self.access_token
            else:
                self.logger.info("Token found in cache but has expired.")
                # with self._token_lock: # Optional locking
                if os.path.exists(self.token_cache_file): os.remove(self.token_cache_file)
                return None
        except (IOError, json.JSONDecodeError, KeyError, ValueError, Exception) as e:
            self.logger.error(f"Failed to load token from cache file {self.token_cache_file}: {e}")
            try:
                # with self._token_lock: # Optional locking
                if os.path.exists(self.token_cache_file): os.remove(self.token_cache_file)
            except OSError as rm_err:
                self.logger.error(f"Failed to remove corrupt token cache file: {rm_err}")
            return None

    def _save_token_to_cache(self, token_response):
        """Saves token information to a simple file cache."""
        if not token_response or 'access_token' not in token_response or 'expires_in' not in token_response:
            self.logger.warning("Invalid token response received, cannot save to cache.")
            return
        try:
            expires_in_seconds = int(token_response['expires_in'])
            expiry_time = datetime.now() + timedelta(seconds=expires_in_seconds)
            cache_data = {
                "access_token": token_response['access_token'],
                "expires_on": expiry_time.isoformat(),
                "scopes": token_response.get('scope', '')
            }
            # with self._token_lock: # Optional locking
            os.makedirs(os.path.dirname(self.token_cache_file), exist_ok=True)
            with open(self.token_cache_file, 'w') as f:
                json.dump(cache_data, f)
            self.logger.info("Token saved to cache.")
        except (IOError, TypeError, Exception) as e:
            self.logger.error(f"Failed to save token to cache file {self.token_cache_file}: {e}")

    def get_access_token(self):
        """Acquires a new access token using MSAL client credentials flow."""
        # with self._token_lock: # Optional locking
        if not self.msal_app:
            self.logger.error("MSAL app not initialized. Cannot acquire token.")
            return None

        cached_token = self._load_token_from_cache()
        if cached_token:
            return cached_token

        self.logger.info(f"Acquiring new access token for scopes: {self.DEFAULT_SCOPES}")
        result = None
        try:
            result = self.msal_app.acquire_token_for_client(scopes=self.DEFAULT_SCOPES)
        except Exception as e:
             self.logger.error(f"Exception during acquire_token_for_client: {e}", exc_info=True)
             return None

        if "access_token" in result:
            self.access_token = result['access_token']
            self.token_expiry = datetime.now() + timedelta(seconds=int(result.get('expires_in', 3599)))
            self.logger.info(f"Successfully acquired new access token, expires around: {self.token_expiry}")
            self._save_token_to_cache(result)
            return self.access_token
        else:
            error = result.get("error")
            error_description = result.get("error_description")
            self.logger.error(f"Failed to acquire token. Error: {error}. Description: {error_description}")
            self.access_token = None
            self.token_expiry = None
            return None

    def ensure_authenticated(self):
        """Ensures a valid, non-expired token is available, acquiring one if needed."""
        if self.access_token and self.token_expiry and self.token_expiry > datetime.now() + timedelta(minutes=5):
             self.logger.debug("Existing token is valid.")
             return True
        else:
             self.logger.info("Existing token missing or expired. Attempting to get/refresh token.")
             new_token = self.get_access_token()
             return new_token is not None

    def make_graph_request(self, method, url_suffix, **kwargs):
        """Makes a request to the Microsoft Graph API with enhanced logging."""
        if not self.ensure_authenticated():
            self.logger.error(f"Authentication failed. Cannot make Graph request to {url_suffix}.")
            return None, 503 # Return None and a simulated service unavailable status

        api_url = f"{self.GRAPH_ENDPOINT}/{url_suffix}"
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Accept': 'application/json'
        }
        if method.upper() in ['POST', 'PATCH', 'PUT'] and 'json' in kwargs and 'Content-Type' not in headers:
            headers['Content-Type'] = 'application/json'

        self.logger.info(f"Making Graph API request: {method} {api_url}")
        if 'json' in kwargs:
            payload_str = json.dumps(kwargs['json'])
            log_payload = payload_str[:1000] + ('...' if len(payload_str) > 1000 else '')
            self.logger.debug(f"Request JSON payload (truncated): {log_payload}")

        response_data = None
        status_code = 500 # Default internal error
        try:
            response = requests.request(method, api_url, headers=headers, timeout=30, **kwargs)
            status_code = response.status_code # Store actual status code
            self.logger.info(f"Graph API response status: {status_code} {response.reason}")

            try:
                response_data = response.json()
                self.logger.debug(f"Graph API Response JSON: {json.dumps(response_data, indent=2)}")
            except json.JSONDecodeError:
                response_data = response.text
                self.logger.debug(f"Graph API Response Text: {response_data}")

            response.raise_for_status() # Raise HTTPError for bad responses AFTER logging

            self.logger.debug(f"Graph API request successful: {method} {api_url}")
            return response_data if status_code != 204 else {}, status_code # Return data and status

        except requests.exceptions.HTTPError as http_err:
            self.logger.error(f"HTTP error calling Graph API {method} {api_url}: {http_err}")
            return response_data, status_code # Return potentially parsed error details and status code
        except requests.exceptions.RequestException as req_err:
            self.logger.error(f"Request exception calling Graph API {method} {api_url}: {req_err}", exc_info=True)
            return None, 500 # Return None and internal error status
        except Exception as e:
            self.logger.error(f"Unexpected error during Graph API call {method} {api_url}: {e}", exc_info=True)
            return None, 500 # Return None and internal error status


    def get_site_drive_id(self):
        """Gets the default drive ID for the configured SharePoint site."""
        if self._drive_id:
            return self._drive_id
        if not self.site_id:
            self.logger.error("Cannot get drive ID: SharePoint Site ID is not configured.")
            return None

        url_suffix = f"sites/{self.site_id}/drive"
        response, status = self.make_graph_request('GET', url_suffix)
        if status == 200 and response and isinstance(response, dict) and 'id' in response:
            self._drive_id = response['id']
            self.logger.info(f"Retrieved drive ID: {self._drive_id} for site ID: {self.site_id}")
            return self._drive_id
        else:
            self.logger.error(f"Failed to get drive ID for site ID: {self.site_id}. Status: {status}, Response: {response}")
            return None

    # --- Function to list root children ---
    def list_drive_root_children(self):
        """Lists files and folders at the root of the site's default drive."""
        drive_id = self.get_site_drive_id()
        if not drive_id:
            self.logger.error("Cannot list root children: Failed to get drive ID.")
            return None

        url_suffix = f"drives/{drive_id}/root/children"
        self.logger.info(f"Listing items in drive root: /{url_suffix}")
        response, status = self.make_graph_request('GET', url_suffix)

        if status == 200 and response and isinstance(response, dict) and 'value' in response:
            children = response['value']
            if children:
                self.logger.info("Found items at drive root:")
                for item in children:
                    item_type = "Folder" if 'folder' in item else "File" if 'file' in item else "Item"
                    self.logger.info(f"- Name: '{item.get('name', 'N/A')}', Type: {item_type}, ID: {item.get('id', 'N/A')}")
            else:
                self.logger.info("Drive root appears to be empty.")
            return children # Return the list of item dicts
        else:
            self.logger.error(f"Failed to list drive root children. Status: {status}, Response: {response}")
            return None

    # --- New Search Function ---
    def search_file_by_name(self, filename):
        """Searches for a file by name within the site's default drive."""
        if not filename:
            self.logger.error("Cannot search: Filename not provided.")
            return None
        drive_id = self.get_site_drive_id()
        if not drive_id:
            self.logger.error(f"Cannot search for '{filename}': Failed to get drive ID.")
            return None

        # Note: Graph search can be slow and might require specific permissions depending on scope.
        # URL encoding the filename is important.
        encoded_filename = requests.utils.quote(filename)
        url_suffix = f"drives/{drive_id}/root/search(q='{encoded_filename}')"
        self.logger.info(f"Searching for file '{filename}' using Graph API path: /{url_suffix}")
        response, status = self.make_graph_request('GET', url_suffix)

        if status == 200 and response and isinstance(response, dict) and 'value' in response:
            search_results = response['value']
            if not search_results:
                self.logger.warning(f"Search for '{filename}' returned no results.")
                return None
            # Assume the first result matching the exact name is the correct one
            # Graph search might return partial matches, so filter carefully
            for item in search_results:
                if item.get('name', '').lower() == filename.lower():
                    file_id = item.get('id')
                    file_path_found = item.get('parentReference', {}).get('path', 'UnknownPath') + '/' + item.get('name', '')
                    # Clean up drive prefix if present (e.g., /drive/root:)
                    file_path_found = re.sub(r'^/drive/root:', '', file_path_found)
                    self.logger.info(f"Search found file '{filename}' with ID: {file_id} at approximate path: {file_path_found}")
                    return file_id, file_path_found # Return ID and the path reported by search
            self.logger.warning(f"Search results found, but no exact match for '{filename}'. Results: {[item.get('name') for item in search_results]}")
            return None # No exact match found
        else:
            self.logger.error(f"Failed to search for file '{filename}'. Status: {status}, Response: {response}")
            return None

    # --- Modified get_file_item with Search Fallback ---
    def get_file_item(self):
        """Gets the drive item ID. Tries direct path first, then search as fallback."""
        # Return cached value if available
        if self._file_item_id and self._file_path_used:
            self.logger.debug(f"Returning cached file item ID: {self._file_item_id} for verified path: {self._file_path_used}")
            return self._file_item_id, self._file_path_used

        if not self.file_path:
            self.logger.error("Cannot get file item: Preferred file path (FILE_PATH in .env) not configured.")
            return None, None

        drive_id = self.get_site_drive_id()
        if not drive_id:
             self.logger.error("Cannot get file item: Failed to get drive ID.")
             return None, None

        # --- Attempt 1: Direct Path Lookup ---
        graph_path = self.file_path.lstrip('/')
        if not graph_path:
             self.logger.error(f"Invalid preferred file path format for Graph API: '{self.file_path}'")
             return None, None

        url_suffix = f"drives/{drive_id}/root:/{graph_path}"
        self.logger.info(f"Attempt 1: Get file item using direct Graph API path: '/{url_suffix}'")
        response, status = self.make_graph_request('GET', url_suffix)

        if status == 200 and response and isinstance(response, dict) and 'id' in response:
            self._file_item_id = response['id']
            self._file_path_used = self.file_path # Store the path that worked
            self.logger.info(f"Direct path lookup SUCCESS. Found ID: {self._file_item_id} for path: {self._file_path_used}")
            return self._file_item_id, self._file_path_used
        else:
            self.logger.warning(f"Direct path lookup failed for '{self.file_path}'. Status: {status}. Attempting search fallback...")

            # --- Attempt 2: Search Fallback ---
            if not self.target_filename:
                 self.logger.error("Cannot search for file: Filename could not be determined from configured path.")
                 return None, None

            search_result = self.search_file_by_name(self.target_filename)
            if search_result:
                file_id, found_path = search_result
                self._file_item_id = file_id
                # Decide whether to update self.file_path or just store the path used
                self._file_path_used = found_path # Store the path found via search
                self.logger.info(f"Search fallback SUCCESS. Found ID: {self._file_item_id}. Using path from search: {self._file_path_used}")
                # Maybe update self.file_path if we want future direct lookups to use the found path?
                # self.file_path = found_path # Optional: Update primary path
                return self._file_item_id, self._file_path_used
            else:
                self.logger.error(f"Both direct path lookup and search failed for file '{self.target_filename}' (based on path '{self.file_path}').")
                return None, None


    def read_worksheet_data(self, worksheet_name):
        """Reads all data from a specific worksheet in the configured Excel file."""
        file_item_id, file_path_used = self.get_file_item()
        if not file_item_id:
            self.logger.error(f"Cannot read worksheet '{worksheet_name}': Failed to get file item ID.")
            return None

        url_suffix = f"sites/{self.site_id}/drive/items/{file_item_id}/workbook/worksheets/{worksheet_name}/usedRange?$select=text"
        self.logger.info(f"Reading used range from worksheet: '{worksheet_name}' in file identified as '{file_path_used}'")
        response, status = self.make_graph_request('GET', url_suffix)

        if status == 200 and response and isinstance(response, dict) and 'text' in response:
            self.logger.info(f"Successfully read data from worksheet '{worksheet_name}'. Rows: {len(response['text'])}")
            return response['text']
        else:
            self.logger.error(f"Failed to read data from worksheet '{worksheet_name}'. Status: {status}, Response: {response}")
            return None


    def update_excel_data(self, data: List[Dict[str, Any]], sheet_name=None, **kwargs):
        """
        Updates or appends rows to the specified Excel sheet.
        Uses the file item ID found by get_file_item (direct path or search).
        """
        if kwargs:
            self.logger.warning(f"update_excel_data received unexpected keyword arguments: {list(kwargs.keys())}")

        if not data:
            self.logger.warning("No data provided to update_excel_data.")
            return False

        target_sheet = sheet_name or "App" # Default sheet is "App"
        self.logger.info(f"Preparing to update sheet '{target_sheet}'...")

        # Use get_file_item which now includes search fallback
        file_item_id, file_path_used = self.get_file_item()
        if not file_item_id:
            # Error already logged by get_file_item
            self.logger.error(f"Cannot update worksheet '{target_sheet}': Failed to obtain file item ID. Save failed.")
            return False

        used_range_url = f"sites/{self.site_id}/drive/items/{file_item_id}/workbook/worksheets/{target_sheet}/usedRange(valuesOnly=true)?$select=address,rowCount"
        self.logger.debug(f"Getting used range address for sheet '{target_sheet}' in file item {file_item_id}")
        range_response, status = self.make_graph_request('GET', used_range_url)

        last_row_index = 0
        start_cell_address = f"{target_sheet}!A1"

        if status == 200 and range_response and isinstance(range_response, dict) and 'address' in range_response:
            # ... (Keep the same logic as V4 for determining last_row_index) ...
            used_address = range_response['address']
            self.logger.debug(f"Sheet '{target_sheet}' used range address: {used_address}")
            last_row_index = range_response.get('rowCount', 0)
            if last_row_index > 0:
                 start_cell_address = f"{target_sheet}!A{last_row_index + 1}"
                 self.logger.info(f"Determined last used row via rowCount: {last_row_index}. Will append data starting at row {last_row_index + 1}.")
            else:
                 try:
                    end_cell = used_address.split('!')[-1].split(':')[-1]
                    last_row_index = int(re.findall(r'\d+$', end_cell)[0])
                    start_cell_address = f"{target_sheet}!A{last_row_index + 1}"
                    self.logger.info(f"Determined last used row via address parse: {last_row_index}. Will append data starting at row {last_row_index + 1}.")
                 except (IndexError, ValueError, Exception) as e:
                    self.logger.warning(f"Could not parse used range address '{used_address}' reliably. Assuming empty or starting at A1. Error: {e}")
                    last_row_index = 0
                    start_cell_address = f"{target_sheet}!A1"
        else:
             # Handle case where usedRange call itself failed
            self.logger.warning(f"Could not get used range for sheet '{target_sheet}'. Status: {status}. Assuming empty sheet, writing from A1.")
            last_row_index = 0
            start_cell_address = f"{target_sheet}!A1"

        if not isinstance(data, list) or not data:
             self.logger.error("Invalid data format: Expected a non-empty list of dictionaries.")
             return False
        if not isinstance(data[0], dict):
            self.logger.error("Invalid data format: List items must be dictionaries.")
            return False

        headers = list(data[0].keys())
        values_payload = []
        for row_dict in data:
            row_list = [row_dict.get(header, "") for header in headers]
            values_payload.append(row_list)

        if not values_payload:
            self.logger.error("Failed to prepare data payload for update.")
            return False

        num_rows_to_add = len(values_payload)
        num_cols = len(headers)
        def col_to_excel(col_num):
            excel_col = ""
            while col_num > 0:
                col_num, remainder = divmod(col_num - 1, 26)
                excel_col = chr(65 + remainder) + excel_col
            return excel_col

        end_col_letter = col_to_excel(num_cols)
        target_range_address = f"{target_sheet}!A{last_row_index + 1}:{end_col_letter}{last_row_index + num_rows_to_add}"
        self.logger.info(f"Updating range: {target_range_address}")

        update_url_suffix = f"sites/{self.site_id}/drive/items/{file_item_id}/workbook/worksheets/{target_sheet}/range(address='{target_range_address}')"
        payload = { "values": values_payload }

        self.logger.info(f"Attempting PATCH request to: {update_url_suffix}")
        self.logger.debug(f"PATCH Payload (first row): {json.dumps(payload['values'][0]) if payload['values'] else 'N/A'}")

        response, status = self.make_graph_request('PATCH', update_url_suffix, json=payload)

        self.logger.info(f"Update Response Data: {response}")
        self.logger.info(f"Update Response Status Code: {status}")

        if status is not None and 200 <= status < 300: # Check for 2xx success status
            self.logger.info(f"Successfully updated data in worksheet '{target_sheet}'.")
            return True
        else:
            self.logger.error(f"Failed to update data in worksheet '{target_sheet}'. Status: {status}, Response: {response}")
            return False


    # --- Legacy/Compatibility Methods ---
    def read_excel_sheet(self, sheet_name, use_cache=True):
        self.logger.warning("read_excel_sheet is potentially deprecated, consider read_worksheet_data.")
        return self.read_sheet(sheet_name, use_cache)

    def read_sheet(self, sheet_name, use_cache=True, cache_ttl=None):
        cache_key = f"sheet_{sheet_name}_{self.site_id}_{self.file_path}"
        if use_cache:
            cached_data = self.cache.get(cache_key)
            if cached_data is not None:
                self.logger.info(f"Returning cached data for sheet '{sheet_name}'")
                return cached_data
        self.logger.info(f"Fetching fresh data for sheet '{sheet_name}'")
        live_data = self.read_worksheet_data(sheet_name)
        if live_data is not None:
            self.cache.set(cache_key, live_data)
        return live_data

    # --- Background Update (Placeholder) ---
    # (Keep V4 code for background update, start, stop)
    def _background_update_task(self, sheet_names):
        self.logger.info(f"Background update thread started for sheets: {sheet_names}")
        while not self._stop_event.is_set():
            for sheet_name in sheet_names:
                if self._stop_event.is_set(): break
                self.logger.info(f"[Background] Updating data for sheet: {sheet_name}")
                try:
                    data = self.read_sheet(sheet_name, use_cache=False)
                    if data is None:
                         self.logger.warning(f"[Background] Failed to fetch fresh data for sheet '{sheet_name}'.")
                except Exception as e:
                    self.logger.error(f"[Background] Error updating sheet '{sheet_name}': {e}", exc_info=True)
                time.sleep(5)

            if self._stop_event.is_set(): break
            wait_time = self._update_interval_seconds
            self.logger.debug(f"[Background] Update cycle complete. Waiting {wait_time} seconds.")
            self._stop_event.wait(wait_time)
        self.logger.info("Background update thread stopped.")

    def start_background_update(self, sheet_names, interval_hours=6):
        if hasattr(self, '_update_thread') and self._update_thread.is_alive():
            self.logger.warning("Background update thread already running.")
            return
        # --- Added stop event initialization if missing ---
        if not hasattr(self, '_stop_event'):
             self._stop_event = threading.Event()
        # --- End change ---
        self._update_interval_seconds = interval_hours * 3600
        self._stop_event.clear() # Ensure event is clear before starting
        self._update_thread = threading.Thread(target=self._background_update_task, args=(sheet_names,), daemon=True)
        self._update_thread.start()
        self.logger.info(f"Background update thread initiated for sheets {sheet_names} with interval {interval_hours} hours.")

    def stop_background_update(self):
        if hasattr(self, '_stop_event'):
            self.logger.info("Requesting background update thread to stop...")
            self._stop_event.set()
        if hasattr(self, '_update_thread') and self._update_thread.is_alive():
            self.logger.debug("Waiting for background thread to join...")
            self._update_thread.join(timeout=10)
            if self._update_thread.is_alive():
                 self.logger.warning("Background thread did not stop gracefully after 10 seconds.")
            else:
                 self.logger.info("Background thread joined successfully.")
        else:
             self.logger.info("Background update thread was not running.")


# --- Compatibility Wrapper Class ---
# (Keep V4 code)
class SharePointExcelManager:
    """Compatibility wrapper for older AMSDealForm usage."""
    def __init__(self, config=None, logger=None):
        self.logger = logger or logging.getLogger(__name__).getChild("SPExcelMgrCompat")
        self.sp_manager = SharePointManager(config=config, logger=self.logger)
        self.logger.info("SharePointExcelManager initialized (compatibility wrapper)")

    def read_excel_sheet(self, sheet_name, use_cache=True):
        return self.sp_manager.read_excel_sheet(sheet_name, use_cache=use_cache)

    def update_excel_data(self, data, sheet_name=None, **kwargs):
        return self.sp_manager.update_excel_data(data, sheet_name=sheet_name, **kwargs)

    def ensure_authenticated(self):
        return self.sp_manager.ensure_authenticated()

    def read_sheet(self, sheet_name, use_cache=True, cache_ttl=None):
        return self.sp_manager.read_sheet(sheet_name, use_cache=use_cache, cache_ttl=cache_ttl)

    def start_background_update(self, sheet_names, interval_hours=6):
        return self.sp_manager.start_background_update(sheet_names, interval_hours=interval_hours)

    def stop_background_update(self):
        return self.sp_manager.stop_background_update()
# home_module.py - Finnhub Calls Commented Out
import os
import logging
import time
import json
import re  # Added missing import
import traceback
import requests
from datetime import datetime, timedelta
import pandas as pd
import numpy as np
# import finnhub # Keep import for now, client init might still be useful later
try:
    import finnhub
except ImportError:
    finnhub = None # Define as None if import fails
    print("WARNING: finnhub library not installed or found.")

from dotenv import load_dotenv
import matplotlib
matplotlib.use('Qt5Agg')  # Use Qt5 backend for matplotlib
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.pyplot as plt
import seaborn as sns
# Use try-except for optional dependencies
try:
    from forex_python.converter import CurrencyRates
except ImportError:
    CurrencyRates = None
    print("WARNING: forex-python library not installed or found. Forex conversion will fail.")
try:
    from forex_python.bitcoin import BtcConverter
except ImportError:
    BtcConverter = None
    print("WARNING: forex-python library not installed or found. Bitcoin conversion will fail.")


from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox,
                            QPushButton, QGridLayout, QFrame, QGroupBox, QScrollArea,
                            QSizePolicy, QSpacerItem, QTabWidget, QTableWidget,
                            QTableWidgetItem, QHeaderView, QProgressBar, QListWidget)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, pyqtSlot, QTimer, QSize
from PyQt5.QtGui import QFont, QColor, QIcon, QPixmap

# Import base module
try:
    from ui.base_module import BaseModule
except ImportError:
    print("ERROR: Failed to import BaseModule from ui.base_module.")
    # Define a dummy BaseModule if import fails, allowing app to potentially run
    class BaseModule(QWidget):
        def __init__(self, main_window=None):
            super().__init__()
            self.main_window = main_window
            print("WARNING: Using dummy BaseModule.")
        def get_title(self): return "Dummy Module"
        def get_icon_name(self): return None
        def refresh(self): pass
        def save_state(self): pass
        def close(self): pass


# Try to import from config, handle failure gracefully
try:
    from utils.config import Config
    logger = logging.getLogger(__name__)
    # logger.info("Successfully imported Config from utils.config") # Less verbose
except ImportError:
    # Fallback config class if import fails
    logger = logging.getLogger(__name__)
    logger.warning("Could not import from config.py. Using dummy values if needed.")

    class Config:
        """Fallback Config class with default values"""
        def __init__(self, base_path=None):
            self.base_path = base_path or os.getcwd()
            self.data_dir = os.path.join(self.base_path, 'data')
            self.log_dir = os.path.join(self.base_path, 'logs')
            self.cache_dir = os.path.join(self.base_path, 'cache')
            self.resources_dir = os.path.join(self.base_path, 'resources')

            # Create directories
            for path in [self.data_dir, self.log_dir, self.cache_dir, self.resources_dir]:
                try:
                    os.makedirs(path, exist_ok=True)
                except OSError as e:
                     logger.error(f"Failed to create fallback directory {path}: {e}")

        def get(self, key, default=None):
            # Basic fallback get method
            return getattr(self, key, default)


# Try to import general_utils, provide fallbacks if needed
try:
    from utils.general_utils import format_currency, parse_env_variable, get_resource_path
except ImportError:
    logger.warning("Could not import from general_utils. Using dummy functions.")

    def format_currency(value, currency='$', decimals=2):
        """Format a value as currency"""
        try:
            return f"{currency}{float(value):,.{decimals}f}" # Added comma formatting
        except (ValueError, TypeError):
             return f"{currency}N/A"

    def parse_env_variable(var, default=None):
        """Parse environment variable with fallback"""
        return os.environ.get(var, default)

    def get_resource_path(filename, resource_dir):
         """Basic fallback resource path getter"""
         if resource_dir and filename:
              return os.path.join(resource_dir, filename)
         return filename # Return filename if dir is missing


# Load environment variables
load_dotenv()

# Constants
REFRESH_INTERVAL_MS = 3600000  # 1 hour in milliseconds
API_REFRESH_INTERVAL_MS = 21600000  # 6 hours in milliseconds
API_TIMEOUT = 15  # API timeout in seconds

# Weather cities
WEATHER_CITIES = [
    {"name": "Camrose", "country_code": "CA", "province": "AB"}, # Use AB for Alberta
    {"name": "Provost", "country_code": "CA", "province": "AB"},
    {"name": "Wainwright", "country_code": "CA", "province": "AB"},
    {"name": "Killam", "country_code": "CA", "province": "AB"}
]

# Financial data to track
STOCK_SYMBOLS = ['DE']  # John Deere
FOREX_PAIRS = ['USDCAD']  # USD/CAD
CRYPTO = ['BTC']  # Bitcoin


class CommodityDataFetcher:
    """Fetches commodity prices from agricultural data APIs"""

    def __init__(self, api_key=None):
        """Initialize the commodity data fetcher.

        Args:
            api_key: API key for commodity data access (optional)
        """
        self.api_key = api_key or os.getenv('COMMODITY_API_KEY')
        self.logger = logging.getLogger(__name__).getChild("CommodityFetcher") # Specific logger
        self.timeout = API_TIMEOUT  # API timeout in seconds

        # Set cache expiration (4 hours) - Note: Caching not implemented here yet
        self.cache_expiration = 4 * 60 * 60  # seconds

    def fetch_commodity_data(self):
        """Fetch current commodity data for wheat and canola.

        Returns:
            Dictionary with commodity price information
        """
        commodity_data = {
            'CANOLA': {},
            'WHEAT': {}
        }

        try:
            # Try to fetch canola price from ICE Futures (Example - Requires specific API access/library)
            # canola_data = self._fetch_canola_price()
            # if canola_data:
            #     commodity_data['CANOLA'] = canola_data
            # self.logger.warning("Canola price fetching from ICE not implemented.")
            canola_data = {} # Placeholder

            # Try to fetch wheat price from CBOT (Chicago Board of Trade) (Example - Requires specific API access/library)
            # wheat_data = self._fetch_wheat_price()
            # if wheat_data:
            #     commodity_data['WHEAT'] = wheat_data
            # self.logger.warning("Wheat price fetching from CBOT not implemented.")
            wheat_data = {} # Placeholder


            # If primary API calls failed/not implemented, try to use agriculture.com data
            # Check if either price is missing
            if not canola_data.get('price_cad_bu') or not wheat_data.get('price_cad_bu'):
                self.logger.info("Attempting to fetch commodity prices from Agriculture.com fallback...")
                fallback_data = self._fetch_agriculture_com_prices()

                # Use fallback data if needed and primary data is empty
                if not canola_data.get('price_cad_bu') and 'CANOLA' in fallback_data:
                    self.logger.info("Using Agriculture.com fallback for Canola.")
                    commodity_data['CANOLA'] = fallback_data['CANOLA']

                if not wheat_data.get('price_cad_bu') and 'WHEAT' in fallback_data:
                    self.logger.info("Using Agriculture.com fallback for Wheat.")
                    commodity_data['WHEAT'] = fallback_data['WHEAT']

            # If still no data, use fallback static values with warning
            if not commodity_data.get('CANOLA') or not commodity_data['CANOLA'].get('price_cad_bu'):
                self.logger.warning("Using static fallback value for Canola price")
                commodity_data['CANOLA'] = {
                    'price_cad_bu': 750.25, # Example static value
                    'timestamp': datetime.now().timestamp(),
                    'source': 'Static Fallback'
                }

            if not commodity_data.get('WHEAT') or not commodity_data['WHEAT'].get('price_cad_bu'):
                self.logger.warning("Using static fallback value for Wheat price")
                commodity_data['WHEAT'] = {
                    'price_cad_bu': 450.75, # Example static value
                    'timestamp': datetime.now().timestamp(),
                    'source': 'Static Fallback'
                }

            return commodity_data

        except Exception as e:
            self.logger.error(f"Error fetching commodity data: {str(e)}", exc_info=True) # Added exc_info

            # Return fallback values on error
            return {
                'CANOLA': {
                    'price_cad_bu': 750.25,
                    'timestamp': datetime.now().timestamp(),
                    'source': 'Error Fallback'
                },
                'WHEAT': {
                    'price_cad_bu': 450.75,
                    'timestamp': datetime.now().timestamp(),
                    'source': 'Error Fallback'
                }
            }

    # --- Placeholder/Example Fetch Methods ---
    # These would need actual implementation using specific APIs/libraries
    # def _fetch_canola_price(self): return {}
    # def _fetch_wheat_price(self): return {}
    # --- End Placeholder ---

    def _fetch_agriculture_com_prices(self):
        """Scrape commodity prices from Agriculture.com as a fallback.

        Returns:
            Dictionary with commodity price information
        """
        # NOTE: Web scraping is fragile and might break if the website structure changes.
        # Using a proper API is always preferred.
        try:
            url = "https://www.agriculture.com/markets/futures"
            headers = { # Set a realistic User-Agent
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
            }
            response = requests.get(url, headers=headers, timeout=self.timeout)
            response.raise_for_status() # Check for HTTP errors (like 404)

            html_content = response.text
            commodities = {}

            # --- Improved Parsing (Example - needs adjustment based on actual site structure) ---
            # This requires inspecting the current HTML structure of agriculture.com/markets/futures
            # Example using simplified regex (likely needs refinement)

            # Look for a pattern like: <a ...>Chicago Wheat</a> ... <span class="value">6.50</span>
            # This is highly dependent on the website's current HTML
            wheat_pattern = r'Chicago Wheat.*?<span[^>]*>([\d\.\,]+)</span'
            wheat_match = re.search(wheat_pattern, html_content, re.IGNORECASE | re.DOTALL)

            if wheat_match:
                try:
                    wheat_price_usd = float(wheat_match.group(1).replace(',', '')) # Handle commas
                    usd_cad_rate = self._get_usd_cad_rate()
                    wheat_price_cad = wheat_price_usd * usd_cad_rate
                    commodities['WHEAT'] = {
                        'price_usd_bu': wheat_price_usd,
                        'price_cad_bu': wheat_price_cad,
                        'timestamp': datetime.now().timestamp(),
                        'source': 'Agriculture.com'
                    }
                    self.logger.info(f"Successfully scraped Wheat price: ${wheat_price_cad:.2f} CAD/bu")
                except Exception as parse_err:
                     self.logger.warning(f"Could not parse scraped Wheat price: {parse_err}")
            else:
                 self.logger.warning("Could not find Wheat price pattern on Agriculture.com")


            # Look for ICE Canola pattern (Example)
            canola_pattern = r'ICE Canola.*?<span[^>]*>([\d\.\,]+)</span'
            canola_match = re.search(canola_pattern, html_content, re.IGNORECASE | re.DOTALL)

            if canola_match:
                try:
                    # Canola is typically quoted in CAD per ton
                    canola_price_cad_ton = float(canola_match.group(1).replace(',', ''))
                    # Convert from CAD/metric ton to CAD/bushel
                    # 1 metric ton of canola = approximately 44.092 bushels (more precise)
                    canola_price_cad_bushel = canola_price_cad_ton / 44.092
                    commodities['CANOLA'] = {
                        'price_cad_ton': canola_price_cad_ton,
                        'price_cad_bu': canola_price_cad_bushel,
                        'timestamp': datetime.now().timestamp(),
                        'source': 'Agriculture.com'
                    }
                    self.logger.info(f"Successfully scraped Canola price: ${canola_price_cad_bushel:.2f} CAD/bu")
                except Exception as parse_err:
                    self.logger.warning(f"Could not parse scraped Canola price: {parse_err}")
            else:
                 self.logger.warning("Could not find Canola price pattern on Agriculture.com")
            # --- End Improved Parsing ---

            return commodities

        except requests.exceptions.RequestException as e:
            self.logger.error(f"Error scraping Agriculture.com: {e}")
            return {}
        except Exception as e: # Catch other potential errors
             self.logger.error(f"Unexpected error during Agriculture.com scraping: {e}", exc_info=True)
             return {}

    def _get_usd_cad_rate(self):
        """Helper to get USD to CAD rate, with fallback."""
        if CurrencyRates:
            try:
                c = CurrencyRates()
                return c.get_rate('USD', 'CAD')
            except Exception as e:
                self.logger.warning(f"Failed to get live USD/CAD rate from forex-python: {e}. Using fallback.")
                return 1.35 # Fallback rate
        else:
             self.logger.warning("forex-python not available. Using fallback USD/CAD rate.")
             return 1.35


class MatplotlibCanvas(FigureCanvas):
    """Canvas for Matplotlib figures in Qt"""
    def __init__(self, figure=None, parent=None, width=5, height=4, dpi=100):
        if figure is None:
            figure = Figure(figsize=(width, height), dpi=dpi)

        self.figure = figure
        self.axes = self.figure.add_subplot(111)

        # Use ggplot style for better aesthetics
        try:
            plt.style.use('ggplot')
        except Exception as e:
             logger.warning(f"Could not apply 'ggplot' style: {e}")

        # Set seaborn style if available
        try:
             sns.set_style("whitegrid")
        except Exception as e:
             logger.warning(f"Could not apply seaborn style: {e}")


        super(MatplotlibCanvas, self).__init__(self.figure)
        self.setParent(parent)

        # Make the canvas expandable
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.updateGeometry()

    def clear(self):
        """Clear the figure"""
        try:
            self.axes.clear()
            self.figure.tight_layout() # May raise error if layout is complex
            self.draw()
        except Exception as e:
             logger.error(f"Error clearing Matplotlib canvas: {e}")


class WeatherFetcherThread(QThread):
    """Thread for fetching weather data from OpenWeather API"""
    weather_fetched = pyqtSignal(dict)
    error_occurred = pyqtSignal(str)

    def __init__(self, cities=None):
        super().__init__()
        self.cities = cities or WEATHER_CITIES
        self.api_key = os.getenv('OPENWEATHER_API_KEY')
        self.timeout = API_TIMEOUT  # API timeout in seconds
        self.logger = logging.getLogger(__name__).getChild("WeatherFetcher") # Specific logger

        if not self.api_key:
            self.logger.error("OpenWeather API key not found in environment variables")
            # Don't emit here, let run handle it
            # self.error_occurred.emit("OpenWeather API key not found")
            return

        self.running = True

    def run(self):
        """Run the thread to fetch weather data"""
        if not self.api_key:
            self.logger.error("Aborting weather fetch: API key missing.")
            self.error_occurred.emit("OpenWeather API key missing")
            return

        try:
            base_url = "https://api.openweathermap.org/data/2.5/weather"
            all_weather_data = {}
            self.logger.info(f"Starting weather fetch for {len(self.cities)} cities...")

            for city in self.cities:
                if not self.running:
                    self.logger.info("Weather fetch thread stopped.")
                    break

                try:
                    city_name = city["name"]
                    country_code = city["country_code"]
                    province = city.get("province", "") # Use get for safety

                    # Build API URL with province if available
                    if province:
                        query = f"{city_name},{province},{country_code}"
                    else:
                        query = f"{city_name},{country_code}"

                    url = f"{base_url}?q={query}&units=metric&appid={self.api_key}"
                    self.logger.debug(f"Fetching weather for: {query} from {url}")

                    # Make the API request
                    response = requests.get(url, timeout=self.timeout)
                    response.raise_for_status()  # Raise an exception for 4xx/5xx status codes

                    # Parse the JSON response
                    weather_data = response.json()
                    self.logger.debug(f"Received weather data for {city_name}")

                    # Store the data
                    all_weather_data[city_name] = weather_data

                    # Sleep to avoid API rate limits
                    time.sleep(0.5)

                except requests.exceptions.RequestException as e:
                    error_msg = f"Network/API error fetching weather for {city_name}: {e}"
                    self.logger.error(error_msg)
                    all_weather_data[city_name] = {'error': error_msg} # Store specific error
                except Exception as e:
                    error_msg = f"Unexpected error fetching weather for {city_name}: {e}"
                    self.logger.error(error_msg, exc_info=True)
                    all_weather_data[city_name] = {'error': error_msg}

            if self.running: # Only emit if not stopped prematurely
                 self.logger.info("Weather fetch complete.")
                 self.weather_fetched.emit(all_weather_data)

        except Exception as e:
            error_msg = f"Critical error in weather fetcher thread: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            self.error_occurred.emit(error_msg)

    def stop(self):
        """Stop the thread"""
        self.logger.info("Stopping weather fetcher thread...")
        self.running = False


class FinancialDataFetcherThread(QThread):
    """Thread for fetching financial data (stocks, forex, crypto, commodities)"""
    data_fetched = pyqtSignal(dict)
    error_occurred = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.alpha_vantage_api_key = os.getenv('ALPHA_VANTAGE_API_KEY')
        self.exchange_rate_api_key = os.getenv('EXCHANGE_RATE_API_KEY')
        self.commodity_api_key = os.getenv('COMMODITY_API_KEY')
        self.timeout = API_TIMEOUT  # API timeout in seconds
        self.running = True
        self.logger = logging.getLogger(__name__).getChild("FinancialFetcher") # Specific logger
        self.finnhub_client = None # Initialize as None

        if not self.alpha_vantage_api_key:
            self.logger.error("Alpha Vantage API key not found in environment variables")
            
        if not self.exchange_rate_api_key:
            self.logger.error("Exchange Rate API key not found in environment variables")

    def run(self):
        """Run the thread to fetch financial data"""
        self.logger.info("Starting financial data fetch...")
        try:
            financial_data = {
                'stocks': {},
                'forex': {},
                'crypto': {},
                'commodities': {}
            }

            # Fetch stock data (John Deere) using Alpha Vantage
            for symbol in STOCK_SYMBOLS:
                if not self.running: break
                self.logger.debug(f"Fetching stock data for {symbol}...")
                try:
                    if self.alpha_vantage_api_key:
                        url = f"https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol={symbol}&apikey={self.alpha_vantage_api_key}"
                        response = requests.get(url, timeout=self.timeout)
                        
                        if response.status_code == 200:
                            data = response.json()
                            quote = data.get('Global Quote', {})
                            
                            if quote:
                                # Create structure similar to what the app expects
                                financial_data['stocks'][symbol] = {
                                    'quote': {
                                        'c': float(quote.get('05. price', 0)) if quote.get('05. price') else None,
                                        'pc': float(quote.get('08. previous close', 0)) if quote.get('08. previous close') else None,
                                        'd': float(quote.get('09. change', 0)) if quote.get('09. change') else None,
                                        'dp': float(quote.get('10. change percent', '0').rstrip('%')) if quote.get('10. change percent') else None
                                    },
                                    'profile': {'name': symbol},
                                    'candles': {'s': 'ok'}  # Placeholder for candles data
                                }
                                self.logger.info(f"Successfully fetched {symbol} stock data from Alpha Vantage")
                            else:
                                financial_data['stocks'][symbol] = {'error': 'No quote data returned'}
                        else:
                            financial_data['stocks'][symbol] = {'error': f'API returned status {response.status_code}'}
                    else:
                        financial_data['stocks'][symbol] = {'error': 'Alpha Vantage API key missing'}
                        
                    time.sleep(0.5)  # Avoid rate limiting

                except Exception as e:
                    self.logger.error(f"Error fetching stock data for {symbol}: {e}", exc_info=False)
                    financial_data['stocks'][symbol] = {'error': str(e)}

            # Fetch forex data (USD/CAD) using Exchange Rate API
            try:
                self.logger.debug("Fetching forex data (USD/CAD)...")
                if self.exchange_rate_api_key:
                    url = f"https://v6.exchangerate-api.com/v6/{self.exchange_rate_api_key}/pair/USD/CAD"
                    response = requests.get(url, timeout=self.timeout)
                    
                    if response.status_code == 200:
                        data = response.json()
                        if data.get('result') == 'success':
                            rate = data.get('conversion_rate')
                            financial_data['forex']['USDCAD'] = {
                                'rate': rate,
                                'timestamp': datetime.now().timestamp()
                            }
                            self.logger.info(f"Fetched USD/CAD rate: {rate:.4f}")
                        else:
                            financial_data['forex']['USDCAD'] = {'error': data.get('error', 'Unknown error')}
                    else:
                        financial_data['forex']['USDCAD'] = {'error': f'API returned status {response.status_code}'}
                else:
                    financial_data['forex']['USDCAD'] = {'error': 'Exchange Rate API key missing'}
            except Exception as e:
                self.logger.error(f"Error fetching USD/CAD rate: {e}", exc_info=True)
                financial_data['forex']['USDCAD'] = {'error': str(e)}

            # Fetch crypto data (BTC in CAD)
            try:
                self.logger.debug("Fetching crypto data (BTC/CAD)...")
                if self.exchange_rate_api_key:
                    # First get the BTC/USD rate from a free API
                    btc_url = "https://api.coinbase.com/v2/prices/BTC-USD/spot"
                    btc_response = requests.get(btc_url, timeout=self.timeout)
                    
                    if btc_response.status_code == 200:
                        btc_data = btc_response.json()
                        btc_usd_price = float(btc_data.get('data', {}).get('amount', 0))
                        
                        # Then convert to CAD using our USD/CAD rate
                        if 'USDCAD' in financial_data.get('forex', {}) and 'rate' in financial_data['forex']['USDCAD']:
                            usd_cad_rate = financial_data['forex']['USDCAD']['rate']
                            btc_cad_price = btc_usd_price * usd_cad_rate
                            financial_data['crypto']['BTC'] = {
                                'price_cad': btc_cad_price,
                                'timestamp': datetime.now().timestamp()
                            }
                            self.logger.info(f"Calculated BTC/CAD price: {btc_cad_price:,.2f}")
                        else:
                            financial_data['crypto']['BTC'] = {'error': 'USD/CAD rate not available for conversion'}
                    else:
                        financial_data['crypto']['BTC'] = {'error': f'API returned status {btc_response.status_code}'}
                else:
                    financial_data['crypto']['BTC'] = {'error': 'API keys missing'}
            except Exception as e:
                self.logger.error(f"Error fetching BTC price in CAD: {e}", exc_info=True)
                financial_data['crypto']['BTC'] = {'error': str(e)}


            # Fetch commodity data (Canola and Wheat)
            try:
                self.logger.debug("Fetching commodity data...")
                commodity_fetcher = CommodityDataFetcher(api_key=self.commodity_api_key)
                commodity_data = commodity_fetcher.fetch_commodity_data()

                # Add commodity data to financial data
                if 'CANOLA' in commodity_data:
                    financial_data['commodities']['CANOLA'] = commodity_data['CANOLA']
                    self.logger.info("Commodity data for CANOLA obtained.")
                if 'WHEAT' in commodity_data:
                    financial_data['commodities']['WHEAT'] = commodity_data['WHEAT']
                    self.logger.info("Commodity data for WHEAT obtained.")

            except Exception as e:
                self.logger.error(f"Error fetching commodity data: {e}", exc_info=True)
                # Use fallback values if fetching fails
                financial_data['commodities']['CANOLA'] = {
                    'price_cad_bu': 750.25,
                    'timestamp': datetime.now().timestamp(),
                    'source': 'Error Fallback'
                }
                financial_data['commodities']['WHEAT'] = {
                    'price_cad_bu': 450.75,
                    'timestamp': datetime.now().timestamp(),
                    'source': 'Error Fallback'
                }

            if self.running: # Only emit if not stopped
                 self.logger.info("Financial data fetch complete.")
                 self.data_fetched.emit(financial_data)

        except Exception as e:
            error_msg = f"Critical error in financial data fetcher thread: {str(e)}"
            self.logger.error(error_msg, exc_info=True)
            if self.running: self.error_occurred.emit(error_msg)

    def stop(self):
        """Stop the thread"""
        self.logger.info("Stopping financial fetcher thread...")
        self.running = False


class ScheduledRefreshThread(QThread):
    """Thread for performing scheduled data refreshes"""
    refresh_signal = pyqtSignal()

    def __init__(self, interval_ms=API_REFRESH_INTERVAL_MS):
        super().__init__()
        self.interval_ms = interval_ms
        self.running = True
        self.logger = logging.getLogger(__name__).getChild("Scheduler") # Specific logger

    def run(self):
        """Run the thread to perform scheduled refreshes"""
        self.logger.info(f"Scheduled refresh started. Interval: {self.interval_ms / 1000 / 60:.1f} minutes.")
        while self.running:
            # Sleep for the specified interval, checking running flag periodically
            sleep_chunk = 1000 # Check every second
            total_slept = 0
            while self.running and total_slept < self.interval_ms:
                 sleep_duration = min(sleep_chunk, self.interval_ms - total_slept)
                 time.sleep(sleep_duration / 1000.0)
                 total_slept += sleep_duration

            if self.running:
                # Emit the refresh signal
                self.logger.info("Scheduled refresh triggered.")
                self.refresh_signal.emit()
        self.logger.info("Scheduled refresh thread finished.")


    def stop(self):
        """Stop the thread"""
        self.logger.info("Stopping scheduled refresh thread...")
        self.running = False


class HomeModule(BaseModule):
    """Home module for displaying dashboard with market data, weather, and charts"""
    # Class attributes for the module loader
    MODULE_DISPLAY_NAME = "Dashboard" # Changed to class attribute
    MODULE_ICON_NAME = "home_icon.png" # Changed to class attribute

    def __init__(self, main_window=None): # Added default None
        # Initialize the base class first
        super().__init__(main_window=main_window) # Pass main_window # CALL SUPER FIRST
        self.setObjectName("HomeModule") # Set object name

        # Get logger from main_window (NOW self.main_window EXISTS)
        parent_logger = getattr(self.main_window, 'logger', None)
        self.logger = parent_logger.getChild("Home") if parent_logger else logging.getLogger(__name__).getChild("Home")
        self.logger.debug("Initializing HomeModule...") # Logger now exists

        # Get the config from main_window if available
        self.config = getattr(self.main_window, 'config', None)
        if not self.config:
            self.logger.warning("MainWindow has no config object, using fallback.")
            self.config = Config(os.path.dirname(os.path.dirname(__file__))) # Basic fallback

        # Initialize data containers
        self.weather_data = {}
        self.financial_data = {
            'stocks': {},
            'forex': {},
            'crypto': {},
            'commodities': {}
        }
        self.current_city = WEATHER_CITIES[0]["name"] if WEATHER_CITIES else None

        # Initialize UI elements to None initially
        # ... (rest of UI element initializations to None) ...
        self.last_updated_label = None

        # Initialize threads
        self.weather_fetcher = None
        self.financial_fetcher = None
        self.scheduled_refresh = None

        # Setup UI (Call it explicitly AFTER logger is ready)
        self.init_ui() # <-- ADD THIS CALL HERE

        # Start data loading after UI is set up
        self.init_data_loading()
        self.logger.debug("HomeModule initialization complete.")


    def init_ui(self):
        """Initialize the user interface"""
        self.logger.debug("Initializing HomeModule UI...")
        if self.layout() is not None:
             self.logger.warning("UI already initialized for HomeModule, skipping.")
             return

        # Main layout
        main_layout = QVBoxLayout(self) # Set layout on self
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        # Title and refresh section
        title_layout = QHBoxLayout()
        title_label = QLabel(self.MODULE_DISPLAY_NAME) # Use constant
        title_label.setFont(QFont('Arial', 16, QFont.Bold))
        title_layout.addWidget(title_label)
        title_layout.addStretch() # Push buttons to the right

        self.last_updated_label = QLabel("Last updated: Never")
        self.last_updated_label.setStyleSheet("font-style: italic; color: grey;")
        title_layout.addWidget(self.last_updated_label)

        self.refresh_button = QPushButton("Refresh Data")
        self.refresh_button.setToolTip("Fetch latest weather and financial data")

        # Safely get resources path from config
        resources_dir = getattr(self.config, 'resources_dir', None) # Use stored config
        icon_path = get_resource_path("refresh.png", resources_dir) if resources_dir else "refresh.png"

        if icon_path and os.path.exists(icon_path):
            self.refresh_button.setIcon(QIcon(icon_path))
            self.refresh_button.setIconSize(QSize(16,16)) # Smaller icon
        else:
             self.logger.warning(f"Refresh icon not found at {icon_path}")
             self.refresh_button.setText("🔄 Refresh") # Fallback text

        self.refresh_button.clicked.connect(self.refresh_data)
        title_layout.addWidget(self.refresh_button)

        # Add title layout to main layout
        main_layout.addLayout(title_layout)

        # Add separator line
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(separator)

        # Create tab widget for different views
        self.tab_widget = QTabWidget()

        # --- Overview Tab ---
        overview_tab = QWidget()
        overview_layout = QHBoxLayout(overview_tab) # Use QHBoxLayout for side-by-side

        # Left side: Financial Overview
        financial_overview_group = QGroupBox("Financial Market Overview")
        financial_overview_layout = QGridLayout(financial_overview_group)

        financial_overview_layout.addWidget(QLabel("John Deere (DE):"), 0, 0)
        self.de_stock_label = QLabel("Loading...")
        self.de_stock_label.setFont(QFont('Arial', 11)) # Slightly smaller font
        financial_overview_layout.addWidget(self.de_stock_label, 0, 1)

        financial_overview_layout.addWidget(QLabel("USD/CAD:"), 1, 0)
        self.usd_cad_label = QLabel("Loading...")
        self.usd_cad_label.setFont(QFont('Arial', 11))
        financial_overview_layout.addWidget(self.usd_cad_label, 1, 1)

        financial_overview_layout.addWidget(QLabel("Bitcoin (CAD):"), 2, 0)
        self.btc_cad_label = QLabel("Loading...")
        self.btc_cad_label.setFont(QFont('Arial', 11))
        financial_overview_layout.addWidget(self.btc_cad_label, 2, 1)

        financial_overview_layout.addWidget(QLabel("Canola (CAD/bu):"), 3, 0)
        self.canola_label = QLabel("Loading...")
        self.canola_label.setFont(QFont('Arial', 11))
        financial_overview_layout.addWidget(self.canola_label, 3, 1)

        financial_overview_layout.addWidget(QLabel("Wheat (CAD/bu):"), 4, 0)
        self.wheat_label = QLabel("Loading...")
        self.wheat_label.setFont(QFont('Arial', 11))
        financial_overview_layout.addWidget(self.wheat_label, 4, 1)

        financial_overview_layout.setColumnStretch(1, 1) # Allow value column to stretch
        financial_overview_layout.setRowStretch(5, 1) # Add stretch at the bottom

        overview_layout.addWidget(financial_overview_group, 1) # Add group with stretch factor 1

        # Right side: Weather Overview
        weather_overview_group = QGroupBox("Alberta Weather Overview")
        weather_overview_layout = QVBoxLayout(weather_overview_group)

        city_selector_layout = QHBoxLayout()
        city_selector_layout.addWidget(QLabel("City:"))
        self.city_selector = QComboBox()
        if WEATHER_CITIES:
             for city in WEATHER_CITIES:
                  self.city_selector.addItem(city["name"])
             self.city_selector.currentIndexChanged.connect(self.city_changed)
        else:
             self.city_selector.addItem("No cities configured")
             self.city_selector.setEnabled(False)
        city_selector_layout.addWidget(self.city_selector, 1) # Allow combobox to stretch
        weather_overview_layout.addLayout(city_selector_layout)

        weather_display_layout = QHBoxLayout() # Layout for icon+temp and details

        self.current_weather_group = QFrame() # Use QFrame for better layout control
        self.current_weather_group.setFrameShape(QFrame.StyledPanel)
        current_weather_layout = QVBoxLayout(self.current_weather_group)
        current_weather_layout.setAlignment(Qt.AlignCenter)

        self.weather_icon_label = QLabel("?")
        self.weather_icon_label.setFixedSize(64, 64) # Fixed size for icon
        self.weather_icon_label.setScaledContents(True)
        self.weather_icon_label.setAlignment(Qt.AlignCenter)
        current_weather_layout.addWidget(self.weather_icon_label)

        self.weather_temp_label = QLabel("--°C")
        self.weather_temp_label.setFont(QFont('Arial', 18, QFont.Bold))
        self.weather_temp_label.setAlignment(Qt.AlignCenter)
        current_weather_layout.addWidget(self.weather_temp_label)

        self.weather_desc_label = QLabel("Loading...")
        self.weather_desc_label.setFont(QFont('Arial', 10))
        self.weather_desc_label.setAlignment(Qt.AlignCenter)
        current_weather_layout.addWidget(self.weather_desc_label)

        weather_display_layout.addWidget(self.current_weather_group)

        self.weather_details_group = QFrame() # Use QFrame
        self.weather_details_group.setFrameShape(QFrame.StyledPanel)
        weather_details_box_layout = QVBoxLayout(self.weather_details_group)
        weather_details_box_layout.setAlignment(Qt.AlignLeft | Qt.AlignVCenter) # Align content

        self.weather_feels_like_label = QLabel("Feels Like: --°C")
        weather_details_box_layout.addWidget(self.weather_feels_like_label)
        self.weather_humidity_label = QLabel("Humidity: --%")
        weather_details_box_layout.addWidget(self.weather_humidity_label)
        self.weather_wind_label = QLabel("Wind: -- m/s")
        weather_details_box_layout.addWidget(self.weather_wind_label)
        self.weather_pressure_label = QLabel("Pressure: -- hPa")
        weather_details_box_layout.addWidget(self.weather_pressure_label)

        weather_display_layout.addWidget(self.weather_details_group)
        weather_overview_layout.addLayout(weather_display_layout)
        weather_overview_layout.addStretch(1) # Push content up

        overview_layout.addWidget(weather_overview_group, 1) # Add group with stretch factor 1

        self.tab_widget.addTab(overview_tab, "Overview")

        # --- Weather Tab ---
        weather_tab = QWidget()
        weather_layout = QVBoxLayout(weather_tab)

        self.weather_chart_group = QGroupBox("Alberta Weather Comparison")
        weather_chart_layout = QVBoxLayout(self.weather_chart_group)
        comparison_options_layout = QHBoxLayout()
        comparison_options_layout.addWidget(QLabel("Compare:"))
        self.weather_comparison_type = QComboBox()
        self.weather_comparison_type.addItems(["Temperature (°C)", "Feels Like (°C)", "Humidity (%)", "Wind Speed (m/s)"])
        self.weather_comparison_type.currentIndexChanged.connect(self.update_weather_chart)
        comparison_options_layout.addWidget(self.weather_comparison_type)
        comparison_options_layout.addStretch(1)
        weather_chart_layout.addLayout(comparison_options_layout)
        self.weather_chart_canvas = MatplotlibCanvas(width=10, height=4, dpi=90) # Adjusted size/dpi
        weather_chart_layout.addWidget(self.weather_chart_canvas)
        weather_layout.addWidget(self.weather_chart_group)

        self.weather_table_group = QGroupBox("All Cities Weather Details")
        weather_table_layout = QVBoxLayout(self.weather_table_group)
        self.weather_table = QTableWidget(0, 6)
        self.weather_table.setHorizontalHeaderLabels(["City", "Temp (°C)", "Feels Like (°C)", "Weather", "Humidity (%)", "Wind (m/s)"])
        header = self.weather_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch) # Weather description can stretch
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.weather_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.weather_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.weather_table.setAlternatingRowColors(True)
        weather_table_layout.addWidget(self.weather_table)
        weather_layout.addWidget(self.weather_table_group)

        self.tab_widget.addTab(weather_tab, "Weather Details")

        # --- Financial Tab ---
        financial_tab = QWidget()
        financial_layout = QVBoxLayout(financial_tab)

        self.stock_chart_group = QGroupBox("John Deere (DE) Stock Price (1 Year)")
        stock_chart_layout = QVBoxLayout(self.stock_chart_group)
        self.stock_chart_canvas = MatplotlibCanvas(width=10, height=4, dpi=90) # Adjusted size/dpi
        stock_chart_layout.addWidget(self.stock_chart_canvas)
        financial_layout.addWidget(self.stock_chart_group)

        self.financial_table_group = QGroupBox("Market Data Summary")
        financial_table_layout = QVBoxLayout(self.financial_table_group)
        self.financial_table = QTableWidget(5, 3) # Rows: DE, USDCAD, BTC, Canola, Wheat
        self.financial_table.setHorizontalHeaderLabels(["Item", "Current Value", "Change / Source"])
        # self.financial_table.setVerticalHeaderLabels(["John Deere (DE)", "USD/CAD", "Bitcoin (CAD)", "Canola (CAD/bu)", "Wheat (CAD/bu)"]) # Set items instead
        header = self.financial_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.financial_table.verticalHeader().setVisible(False) # Hide row numbers
        self.financial_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.financial_table.setAlternatingRowColors(True)
        financial_table_layout.addWidget(self.financial_table)
        financial_layout.addWidget(self.financial_table_group)

        self.tab_widget.addTab(financial_tab, "Financial Details")

        # Add tab widget to main layout
        main_layout.addWidget(self.tab_widget)

        # Status bar area (handled by main window)
        # self.status_label = QLabel("Ready") # Status label now part of main window status bar

        # Set the layout (already done by QVBoxLayout(self))
        # self.setLayout(main_layout)
        self.logger.debug("HomeModule UI initialization finished.")


    def init_data_loading(self):
        """Initialize data loading"""
        self.logger.debug("Initializing data loading...")
        # Safely get cache directory
        cache_dir = getattr(self.config, 'cache_dir', None) # Use stored config
        if not cache_dir:
            self.logger.warning("Cache directory not configured, caching disabled.")
        else:
             # Ensure cache subdirectories exist
             try:
                  os.makedirs(os.path.join(cache_dir, 'market_data'), exist_ok=True)
                  os.makedirs(os.path.join(cache_dir, 'weather_data'), exist_ok=True)
             except OSError as e:
                  self.logger.error(f"Could not create cache subdirectories in {cache_dir}: {e}")
                  cache_dir = None # Disable caching if dirs can't be created

        # Load cached data first
        if cache_dir:
            self.load_cached_weather_data()
            self.load_cached_financial_data()
        else:
             self.logger.info("Skipping cache load as cache directory is unavailable.")

        # Update UI with cached data (or loading state) immediately
        self.update_ui()

        # Then start data fetcher threads
        self.refresh_data() # Initial fetch

        # Start scheduled refresh thread
        # Use interval from config if available, else default
        refresh_hours = self.config.get('api_refresh_interval_hours', 6) # Example config key
        refresh_interval_ms = int(refresh_hours * 3600 * 1000)

        if hasattr(self, 'scheduled_refresh') and self.scheduled_refresh and self.scheduled_refresh.isRunning():
             self.scheduled_refresh.stop()
             self.scheduled_refresh.wait()

        self.scheduled_refresh = ScheduledRefreshThread(interval_ms=refresh_interval_ms)
        self.scheduled_refresh.refresh_signal.connect(self.refresh_data)
        self.scheduled_refresh.start()


    def load_cached_weather_data(self):
        """Load weather data from cache if available"""
        cache_dir = getattr(self.config, 'cache_dir', None)
        if not cache_dir: return # Skip if no cache dir

        cache_file = os.path.join(cache_dir, 'weather_data', 'weather_data.json')
        if os.path.exists(cache_file):
            try:
                with open(cache_file, 'r', encoding='utf-8') as f: # Specify encoding
                    cache_data = json.load(f)

                # Check cache age (e.g., less than 1 hour for weather)
                cache_age_limit_hours = self.config.get('weather_cache_hours', 1)
                if 'timestamp' in cache_data:
                    cache_time = datetime.fromtimestamp(cache_data['timestamp'])
                    if datetime.now() - cache_time < timedelta(hours=cache_age_limit_hours):
                        self.weather_data = cache_data.get('data', {}) # Use get for safety
                        self.logger.info(f"Loaded weather data for {len(self.weather_data)} cities from cache.")
                        return # Success
                    else:
                         self.logger.info("Cached weather data is too old.")
                else:
                     self.logger.warning("Cached weather data missing timestamp.")

            except json.JSONDecodeError as e:
                 self.logger.error(f"Error decoding cached weather data JSON: {e}")
                 # Optionally delete corrupted cache file
                 # os.remove(cache_file)
            except Exception as e:
                self.logger.error(f"Error loading cached weather data: {e}", exc_info=True)

        self.logger.info("No valid cached weather data found.")


    def load_cached_financial_data(self):
        """Load financial data from cache if available"""
        cache_dir = getattr(self.config, 'cache_dir', None)
        if not cache_dir: return # Skip if no cache dir

        cache_file = os.path.join(cache_dir, 'market_data', 'financial_data.json')
        if os.path.exists(cache_file):
            try:
                with open(cache_file, 'r', encoding='utf-8') as f: # Specify encoding
                    cache_data = json.load(f)

                # Check cache age (e.g., less than 4 hours for financial)
                cache_age_limit_hours = self.config.get('financial_cache_hours', 4)
                if 'timestamp' in cache_data:
                    cache_time = datetime.fromtimestamp(cache_data['timestamp'])
                    if datetime.now() - cache_time < timedelta(hours=cache_age_limit_hours):
                        self.financial_data = cache_data.get('data', {}) # Use get for safety
                        self.logger.info("Loaded financial data from cache.")
                        return # Success
                    else:
                         self.logger.info("Cached financial data is too old.")
                else:
                     self.logger.warning("Cached financial data missing timestamp.")

            except json.JSONDecodeError as e:
                 self.logger.error(f"Error decoding cached financial data JSON: {e}")
                 # Optionally delete corrupted cache file
                 # os.remove(cache_file)
            except Exception as e:
                self.logger.error(f"Error loading cached financial data: {e}", exc_info=True)

        self.logger.info("No valid cached financial data found.")


    def save_weather_data_to_cache(self):
        """Save weather data to cache"""
        cache_dir = getattr(self.config, 'cache_dir', None)
        if not cache_dir or not self.weather_data: return # Skip if no cache dir or no data

        cache_file = os.path.join(cache_dir, 'weather_data', 'weather_data.json')
        try:
            cache_data = {
                'timestamp': datetime.now().timestamp(),
                'data': self.weather_data
            }
            with open(cache_file, 'w', encoding='utf-8') as f: # Specify encoding
                json.dump(cache_data, f, indent=2) # Add indent for readability
            self.logger.info("Saved weather data to cache.")
        except Exception as e:
            self.logger.error(f"Error saving weather data to cache: {e}", exc_info=True)


    def save_financial_data_to_cache(self):
        """Save financial data to cache"""
        cache_dir = getattr(self.config, 'cache_dir', None)
        if not cache_dir or not self.financial_data: return # Skip if no cache dir or no data

        cache_file = os.path.join(cache_dir, 'market_data', 'financial_data.json')
        try:
            cache_data = {
                'timestamp': datetime.now().timestamp(),
                'data': self.financial_data
            }
            with open(cache_file, 'w', encoding='utf-8') as f: # Specify encoding
                json.dump(cache_data, f, indent=2) # Add indent for readability
            self.logger.info("Saved financial data to cache.")
        except Exception as e:
            self.logger.error(f"Error saving financial data to cache: {e}", exc_info=True)


    @pyqtSlot() # Mark as slot for signal connection
    def refresh_data(self):
        """Refresh all data by starting fetcher threads."""
        self.logger.info("Refresh data requested.")
        # Update last updated time immediately
        if hasattr(self, 'last_updated_label') and self.last_updated_label:
             self.last_updated_label.setText(f"Updating... ({datetime.now().strftime('%H:%M:%S')})")
        self.refresh_weather_data()
        self.refresh_financial_data()


    def refresh_weather_data(self):
        """Refresh weather data"""
        # Stop existing thread if running
        if self.weather_fetcher and self.weather_fetcher.isRunning():
            self.logger.debug("Stopping previous weather fetcher thread...")
            self.weather_fetcher.stop()
            if not self.weather_fetcher.wait(1000): # Wait up to 1 sec
                 self.logger.warning("Previous weather fetcher thread did not stop cleanly.")

        # Update UI status using main window's method if available
        if hasattr(self.main_window, 'update_status'):
             self.main_window.update_status("Fetching weather data...")
        if hasattr(self.main_window, 'show_loading'):
             self.main_window.show_loading("Fetching Weather...") # Show loading overlay

        if hasattr(self, 'refresh_button'): self.refresh_button.setEnabled(False)

        # Start new thread
        self.logger.debug("Starting new weather fetcher thread.")
        self.weather_fetcher = WeatherFetcherThread(WEATHER_CITIES)
        # Disconnect previous signals if any to prevent duplicates
        try: self.weather_fetcher.weather_fetched.disconnect()
        except TypeError: pass # Ignore error if not connected
        try: self.weather_fetcher.error_occurred.disconnect()
        except TypeError: pass
        try: self.weather_fetcher.finished.disconnect()
        except TypeError: pass
        # Connect signals
        self.weather_fetcher.weather_fetched.connect(self.handle_fetched_weather_data)
        self.weather_fetcher.error_occurred.connect(self.handle_fetch_error)
        self.weather_fetcher.finished.connect(lambda: self.fetch_finished("weather"))
        self.weather_fetcher.start()


    def refresh_financial_data(self):
        """Refresh financial data"""
        # Stop existing thread if running
        if self.financial_fetcher and self.financial_fetcher.isRunning():
            self.logger.debug("Stopping previous financial fetcher thread...")
            self.financial_fetcher.stop()
            if not self.financial_fetcher.wait(1000): # Wait up to 1 sec
                 self.logger.warning("Previous financial fetcher thread did not stop cleanly.")

        # Update UI status
        if hasattr(self.main_window, 'update_status'):
             self.main_window.update_status("Fetching financial data...")
        if hasattr(self.main_window, 'show_loading'):
             self.main_window.show_loading("Fetching Market Data...") # Show loading overlay

        if hasattr(self, 'refresh_button'): self.refresh_button.setEnabled(False)

        # Start new thread
        self.logger.debug("Starting new financial fetcher thread.")
        self.financial_fetcher = FinancialDataFetcherThread()
        # Disconnect previous signals
        try: self.financial_fetcher.data_fetched.disconnect()
        except TypeError: pass
        try: self.financial_fetcher.error_occurred.disconnect()
        except TypeError: pass
        try: self.financial_fetcher.finished.disconnect()
        except TypeError: pass
        # Connect signals
        self.financial_fetcher.data_fetched.connect(self.handle_fetched_financial_data)
        self.financial_fetcher.error_occurred.connect(self.handle_fetch_error)
        self.financial_fetcher.finished.connect(lambda: self.fetch_finished("financial"))
        self.financial_fetcher.start()


    @pyqtSlot(dict) # Mark as slot
    def handle_fetched_weather_data(self, data):
        """Handle fetched weather data"""
        self.logger.info(f"Received updated weather data for {len(data)} cities.")
        self.weather_data = data
        self.update_weather_ui()
        self.save_weather_data_to_cache()


    @pyqtSlot(dict) # Mark as slot
    def handle_fetched_financial_data(self, data):
        """Handle fetched financial data"""
        self.logger.info("Received updated financial data.")
        self.financial_data = data
        self.update_financial_ui()
        self.save_financial_data_to_cache()


    @pyqtSlot(str) # Mark as slot
    def handle_fetch_error(self, error_msg):
        """Handle fetch errors"""
        self.logger.error(f"Data fetch error reported: {error_msg}")
        if hasattr(self.main_window, 'update_status'):
             self.main_window.update_status(f"Error: {error_msg}", 10000) # Show error longer


    @pyqtSlot(str) # Mark as slot, accept argument
    def fetch_finished(self, data_type):
        """Handle fetch completion"""
        self.logger.info(f"Fetch finished for: {data_type}")
        # Only hide loading/enable button if BOTH fetches are done (or handle separately)
        weather_running = self.weather_fetcher and self.weather_fetcher.isRunning()
        financial_running = self.financial_fetcher and self.financial_fetcher.isRunning()

        if not weather_running and not financial_running:
            self.logger.info("All data fetches complete.")
            if hasattr(self.main_window, 'hide_loading'):
                 self.main_window.hide_loading()
            if hasattr(self, 'refresh_button'): self.refresh_button.setEnabled(True)
            if hasattr(self.main_window, 'update_status'):
                 self.main_window.update_status("Ready")
            if hasattr(self, 'last_updated_label') and self.last_updated_label:
                 self.last_updated_label.setText(f"Last updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")


    @pyqtSlot() # Mark as slot
    def city_changed(self):
        """Handle city selection change"""
        if hasattr(self, 'city_selector') and self.city_selector:
             self.current_city = self.city_selector.currentText()
             self.logger.debug(f"City changed to: {self.current_city}")
             self.update_city_weather()
        else:
             self.logger.warning("city_changed called but city_selector does not exist.")


    def update_ui(self):
         """Update all UI elements with current data."""
         self.logger.debug("Updating all HomeModule UI elements.")
         self.update_weather_ui()
         self.update_financial_ui()


    def update_weather_ui(self):
        """Update all weather UI elements"""
        self.logger.debug("Updating weather UI.")
        self.update_city_weather()
        self.update_weather_table()
        self.update_weather_chart()


    def update_financial_ui(self):
        """Update all financial UI elements"""
        self.logger.debug("Updating financial UI.")
        self.update_financial_overview()
        self.update_financial_table()
        self.update_stock_chart()


    def update_city_weather(self):
        """Update weather display for the selected city"""
        # Check if UI elements exist before updating
        if not all(hasattr(self, attr) and getattr(self, attr) is not None for attr in [
            'weather_data', 'current_city', 'weather_temp_label', 'weather_desc_label',
            'weather_icon_label', 'weather_feels_like_label', 'weather_humidity_label',
            'weather_wind_label', 'weather_pressure_label'
        ]):
            self.logger.warning("update_city_weather: Skipping, required UI elements or data missing.")
            return

        if not self.current_city or self.current_city not in self.weather_data:
            self.logger.warning(f"No weather data available for selected city: {self.current_city}")
            # Reset labels to default/error state
            self.weather_temp_label.setText("--°C")
            self.weather_desc_label.setText("Unavailable")
            self.weather_icon_label.setText("?")
            self.weather_feels_like_label.setText("Feels Like: --°C")
            self.weather_humidity_label.setText("Humidity: --%")
            self.weather_wind_label.setText("Wind: -- m/s")
            self.weather_pressure_label.setText("Pressure: -- hPa")
            return

        city_weather = self.weather_data[self.current_city]

        # Check if there was an error for this city
        if 'error' in city_weather:
            error_msg = city_weather['error']
            self.logger.error(f"Error loading weather for {self.current_city}: {error_msg}")
            # Update labels to show error
            self.weather_temp_label.setText("ERR")
            self.weather_desc_label.setText("Error")
            self.weather_icon_label.setText("!")
            self.weather_feels_like_label.setText("Feels Like: ERR")
            self.weather_humidity_label.setText("Humidity: ERR")
            self.weather_wind_label.setText("Wind: ERR")
            self.weather_pressure_label.setText("Pressure: ERR")
            if hasattr(self.main_window, 'update_status'):
                 self.main_window.update_status(f"Weather Error: {self.current_city}", 5000)
            return

        # Extract weather data safely using .get()
        main_data = city_weather.get('main', {})
        weather_desc_list = city_weather.get('weather', [{}])
        weather_desc = weather_desc_list[0] if weather_desc_list else {}
        wind_data = city_weather.get('wind', {})

        # Update temperature
        temp = main_data.get('temp') # Returns None if key missing
        self.weather_temp_label.setText(f"{temp:.1f}°C" if temp is not None else "--°C")

        # Update weather description
        desc = weather_desc.get('description')
        self.weather_desc_label.setText(desc.capitalize() if desc else "Unknown")

        # Update weather icon
        icon_code = weather_desc.get('icon')
        if icon_code:
            # Try to download the icon from OpenWeather
            try:
                icon_url = f"https://openweathermap.org/img/wn/{icon_code}@2x.png" # Use HTTPS
                response = requests.get(icon_url, timeout=5) # Add timeout
                if response.status_code == 200:
                    pixmap = QPixmap()
                    pixmap.loadFromData(response.content)
                    self.weather_icon_label.setPixmap(pixmap)
                else:
                    self.logger.warning(f"Failed to download weather icon {icon_code}, status: {response.status_code}")
                    self.weather_icon_label.setText(f"{icon_code}") # Show code as fallback
            except Exception as e:
                self.logger.error(f"Error loading weather icon {icon_code}: {e}")
                self.weather_icon_label.setText("?") # Error indicator
        else:
            self.weather_icon_label.setText("?") # No icon code

        # Update weather details
        feels_like = main_data.get('feels_like')
        self.weather_feels_like_label.setText(f"Feels Like: {feels_like:.1f}°C" if feels_like is not None else "Feels Like: --°C")

        humidity = main_data.get('humidity')
        self.weather_humidity_label.setText(f"Humidity: {humidity}%" if humidity is not None else "Humidity: --%")

        wind_speed = wind_data.get('speed') # m/s
        wind_deg = wind_data.get('deg')
        if wind_speed is not None:
             wind_speed_kmh = wind_speed * 3.6 # Convert m/s to km/h
             wind_direction = self.get_wind_direction(wind_deg) if wind_deg is not None else ""
             self.weather_wind_label.setText(f"Wind: {wind_speed_kmh:.1f} km/h {wind_direction}")
        else:
             self.weather_wind_label.setText("Wind: -- km/h")


        pressure = main_data.get('pressure')
        self.weather_pressure_label.setText(f"Pressure: {pressure} hPa" if pressure is not None else "Pressure: -- hPa")


    def update_weather_table(self):
        """Update the weather table with all cities"""
        if not hasattr(self, 'weather_table') or not self.weather_table:
             self.logger.warning("update_weather_table: Skipping, table widget missing.")
             return
        if not self.weather_data:
            self.logger.debug("update_weather_table: No weather data available.")
            self.weather_table.setRowCount(0) # Clear table if no data
            return

        self.weather_table.setSortingEnabled(False) # Disable sorting during update
        self.weather_table.setRowCount(0) # Clear existing rows

        # Add weather data for each city
        row = 0
        for city_name in sorted(self.weather_data.keys()):
            data = self.weather_data[city_name]

            # Skip if there was an error for this city
            if 'error' in data:
                self.logger.warning(f"Skipping weather table row for {city_name} due to error: {data['error']}")
                continue

            self.weather_table.insertRow(row)

            # City name
            city_item = QTableWidgetItem(city_name)
            self.weather_table.setItem(row, 0, city_item)

            # Temperature
            temp = data.get('main', {}).get('temp')
            temp_item = QTableWidgetItem(f"{temp:.1f}" if temp is not None else "N/A")
            temp_item.setTextAlignment(Qt.AlignCenter)
            self.weather_table.setItem(row, 1, temp_item)

            # Feels like
            feels_like = data.get('main', {}).get('feels_like')
            feels_like_item = QTableWidgetItem(f"{feels_like:.1f}" if feels_like is not None else "N/A")
            feels_like_item.setTextAlignment(Qt.AlignCenter)
            self.weather_table.setItem(row, 2, feels_like_item)

            # Weather description
            desc = data.get('weather', [{}])[0].get('description', 'Unknown')
            desc_item = QTableWidgetItem(desc.capitalize() if desc else "N/A")
            self.weather_table.setItem(row, 3, desc_item)

            # Humidity
            humidity = data.get('main', {}).get('humidity')
            humidity_item = QTableWidgetItem(f"{humidity}%" if humidity is not None else "N/A")
            humidity_item.setTextAlignment(Qt.AlignCenter)
            self.weather_table.setItem(row, 4, humidity_item)

            # Wind
            wind_speed = data.get('wind', {}).get('speed') # m/s
            wind_deg = data.get('wind', {}).get('deg')
            wind_text = "N/A"
            if wind_speed is not None:
                 wind_speed_kmh = wind_speed * 3.6
                 wind_direction = self.get_wind_direction(wind_deg) if wind_deg is not None else ""
                 wind_text = f"{wind_speed_kmh:.1f} km/h {wind_direction}"
            wind_item = QTableWidgetItem(wind_text)
            wind_item.setTextAlignment(Qt.AlignCenter)
            self.weather_table.setItem(row, 5, wind_item)

            row += 1
        self.weather_table.setSortingEnabled(True) # Re-enable sorting


    def update_weather_chart(self):
        """Update the weather comparison chart"""
        if not hasattr(self, 'weather_chart_canvas') or not self.weather_chart_canvas:
             self.logger.warning("update_weather_chart: Skipping, canvas missing.")
             return
        if not self.weather_data:
            self.logger.debug("update_weather_chart: No weather data available.")
            self.weather_chart_canvas.axes.clear()
            self.weather_chart_canvas.axes.text(0.5, 0.5, "No Weather Data", ha='center', va='center')
            self.weather_chart_canvas.draw()
            return

        # Get comparison type
        comparison_type = self.weather_comparison_type.currentText() if self.weather_comparison_type else "Temperature (°C)"

        # Prepare data for chart
        cities = []
        values = []
        colors = []
        ylabel = ""
        title = ""

        try:
            if comparison_type == "Temperature (°C)":
                ylabel = "Temperature (°C)"
                title = "Temperature Comparison"
                for city_name in sorted(self.weather_data.keys()):
                    data = self.weather_data[city_name]
                    if 'error' not in data and 'main' in data:
                        temp = data['main'].get('temp')
                        if temp is not None:
                             cities.append(city_name)
                             values.append(temp)
                             colors.append(self.get_temp_color(temp))

            elif comparison_type == "Feels Like (°C)":
                ylabel = "Feels Like (°C)"
                title = "Feels Like Temperature Comparison"
                for city_name in sorted(self.weather_data.keys()):
                    data = self.weather_data[city_name]
                    if 'error' not in data and 'main' in data:
                        feels_like = data['main'].get('feels_like')
                        if feels_like is not None:
                             cities.append(city_name)
                             values.append(feels_like)
                             colors.append(self.get_temp_color(feels_like))

            elif comparison_type == "Humidity (%)":
                ylabel = "Humidity (%)"
                title = "Humidity Comparison"
                for city_name in sorted(self.weather_data.keys()):
                    data = self.weather_data[city_name]
                    if 'error' not in data and 'main' in data:
                        humidity = data['main'].get('humidity')
                        if humidity is not None:
                             cities.append(city_name)
                             values.append(humidity)
                             colors.append(self.get_humidity_color(humidity))

            else:  # Wind Speed (m/s)
                ylabel = "Wind Speed (m/s)"
                title = "Wind Speed Comparison"
                for city_name in sorted(self.weather_data.keys()):
                    data = self.weather_data[city_name]
                    if 'error' not in data and 'wind' in data:
                        wind_speed = data['wind'].get('speed')
                        if wind_speed is not None:
                             cities.append(city_name)
                             values.append(wind_speed)
                             colors.append(self.get_wind_color(wind_speed))

            # Clear previous chart
            self.weather_chart_canvas.axes.clear()

            if not cities: # Handle case where no valid data was found for the selected type
                 self.weather_chart_canvas.axes.text(0.5, 0.5, f"No Data for {comparison_type}", ha='center', va='center')
            else:
                 # Create the bar chart
                 bars = self.weather_chart_canvas.axes.bar(cities, values, color=colors)

                 # Add value labels on top of bars
                 for bar in bars:
                     height = bar.get_height()
                     label = ""
                     if comparison_type == "Temperature (°C)" or comparison_type == "Feels Like (°C)":
                         label = f'{height:.1f}°C'
                     elif comparison_type == "Humidity (%)":
                         label = f'{height:.0f}%'
                     else:  # Wind Speed
                         label = f'{height:.1f} m/s'

                     self.weather_chart_canvas.axes.text(
                         bar.get_x() + bar.get_width()/2., height, label,
                         ha='center', va='bottom', fontsize=9 # Smaller font for labels
                     )

                 # Customize chart
                 self.weather_chart_canvas.axes.set_title(title, fontsize=10)
                 self.weather_chart_canvas.axes.set_ylabel(ylabel, fontsize=9)
                 self.weather_chart_canvas.axes.tick_params(axis='x', labelsize=8, rotation=15) # Rotate labels slightly
                 self.weather_chart_canvas.axes.tick_params(axis='y', labelsize=8)
                 self.weather_chart_canvas.axes.grid(True, axis='y', linestyle='--', alpha=0.6)
                 # Adjust y-axis limits slightly
                 if values:
                      min_val = min(values)
                      max_val = max(values)
                      padding = (max_val - min_val) * 0.1 + 1 # Add a small absolute padding too
                      self.weather_chart_canvas.axes.set_ylim(min(0, min_val - padding), max_val + padding)


            # Adjust layout and draw
            self.weather_chart_canvas.figure.tight_layout()
            self.weather_chart_canvas.draw()

        except Exception as e:
             self.logger.error(f"Error updating weather chart: {e}", exc_info=True)
             try: # Try to show error on canvas
                  self.weather_chart_canvas.axes.clear()
                  self.weather_chart_canvas.axes.text(0.5, 0.5, "Error generating chart", ha='center', va='center')
                  self.weather_chart_canvas.draw()
             except Exception: pass # Ignore errors during error display


    def update_financial_overview(self):
        """Update the financial overview section"""
        # Check if UI elements exist
        if not all(hasattr(self, attr) and getattr(self, attr) is not None for attr in [
             'de_stock_label', 'usd_cad_label', 'btc_cad_label', 'canola_label', 'wheat_label'
        ]):
             self.logger.warning("update_financial_overview: Skipping, required UI labels missing.")
             return

        # Default text
        na_text = "N/A"

        # John Deere stock
        de_text = na_text
        de_style = ""
        if 'DE' in self.financial_data.get('stocks', {}):
            de_data = self.financial_data['stocks']['DE']
            if 'error' not in de_data and 'quote' in de_data:
                quote = de_data['quote']
                current_price = quote.get('c')
                prev_close = quote.get('pc')
                if current_price is not None and prev_close is not None and prev_close != 0:
                    change = current_price - prev_close
                    change_pct = (change / prev_close * 100)
                    arrow = "↑" if change >= 0 else "↓"
                    color = "green" if change >= 0 else "red"
                    de_text = f"${current_price:.2f} ({arrow} {abs(change):.2f}, {change_pct:+.2f}%)"
                    de_style = f"color: {color};"
                elif current_price is not None:
                     de_text = f"${current_price:.2f}" # Show price if change unavailable
            elif 'error' in de_data:
                 de_text = f"Error: {de_data['error'][:30]}..." # Show truncated error

        self.de_stock_label.setText(de_text)
        self.de_stock_label.setStyleSheet(de_style)

        # USD/CAD rate
        usdcad_text = na_text
        if 'USDCAD' in self.financial_data.get('forex', {}):
            usdcad_data = self.financial_data['forex']['USDCAD']
            if 'error' not in usdcad_data and 'rate' in usdcad_data:
                rate = usdcad_data.get('rate')
                if rate is not None:
                     usdcad_text = f"{rate:.4f} CAD"
            elif 'error' in usdcad_data:
                 usdcad_text = f"Error: {usdcad_data['error'][:30]}..."

        self.usd_cad_label.setText(usdcad_text)
        self.usd_cad_label.setStyleSheet("") # Reset style

        # BTC in CAD
        btc_text = na_text
        if 'BTC' in self.financial_data.get('crypto', {}):
            btc_data = self.financial_data['crypto']['BTC']
            if 'error' not in btc_data and 'price_cad' in btc_data:
                price = btc_data.get('price_cad')
                if price is not None:
                     btc_text = format_currency(price, currency='$', decimals=2) # Use helper
            elif 'error' in btc_data:
                 btc_text = f"Error: {btc_data['error'][:30]}..."

        self.btc_cad_label.setText(btc_text)
        self.btc_cad_label.setStyleSheet("")

        # Canola price
        canola_text = na_text
        if 'CANOLA' in self.financial_data.get('commodities', {}):
            canola_data = self.financial_data['commodities']['CANOLA']
            if 'error' not in canola_data and 'price_cad_bu' in canola_data:
                price = canola_data.get('price_cad_bu')
                source = canola_data.get('source', '')
                if price is not None:
                     source_tag = f" ({source})" if source and source not in ['API', 'Static Fallback', 'Error Fallback'] else ""
                     canola_text = f"{format_currency(price, decimals=2)}{source_tag}"
            elif 'error' in canola_data:
                 canola_text = f"Error: {canola_data['error'][:30]}..."

        self.canola_label.setText(canola_text)
        self.canola_label.setStyleSheet("")

        # Wheat price
        wheat_text = na_text
        if 'WHEAT' in self.financial_data.get('commodities', {}):
            wheat_data = self.financial_data['commodities']['WHEAT']
            if 'error' not in wheat_data and 'price_cad_bu' in wheat_data:
                price = wheat_data.get('price_cad_bu')
                source = wheat_data.get('source', '')
                if price is not None:
                     source_tag = f" ({source})" if source and source not in ['API', 'Static Fallback', 'Error Fallback'] else ""
                     wheat_text = f"{format_currency(price, decimals=2)}{source_tag}"
            elif 'error' in wheat_data:
                 wheat_text = f"Error: {wheat_data['error'][:30]}..."

        self.wheat_label.setText(wheat_text)
        self.wheat_label.setStyleSheet("")


    def update_financial_table(self):
        """Update the financial data table"""
        if not hasattr(self, 'financial_table') or not self.financial_table:
             self.logger.warning("update_financial_table: Skipping, table widget missing.")
             return

        self.financial_table.setSortingEnabled(False)
        na_item = QTableWidgetItem("N/A")
        na_item.setTextAlignment(Qt.AlignCenter)

        # Helper to create table items
        def create_item(text, align=Qt.AlignLeft | Qt.AlignVCenter):
             item = QTableWidgetItem(str(text))
             item.setTextAlignment(align)
             return item

        # John Deere stock (Row 0)
        de_item = create_item("John Deere (DE)")
        de_val_item = create_item("N/A", Qt.AlignRight | Qt.AlignVCenter)
        de_chg_item = create_item("N/A", Qt.AlignCenter)
        if 'DE' in self.financial_data.get('stocks', {}):
            de_data = self.financial_data['stocks']['DE']
            if 'error' not in de_data and 'quote' in de_data:
                quote = de_data['quote']
                current_price = quote.get('c')
                prev_close = quote.get('pc')
                if current_price is not None:
                     de_val_item.setText(format_currency(current_price))
                if current_price is not None and prev_close is not None and prev_close != 0:
                    change = current_price - prev_close
                    change_pct = (change / prev_close * 100)
                    de_chg_item.setText(f"{change:+.2f} ({change_pct:+.2f}%)")
                    de_chg_item.setForeground(QColor('green') if change >= 0 else QColor('red'))
            elif 'error' in de_data:
                 de_val_item.setText("Error")
                 de_chg_item.setText(de_data['error'][:15]+"...") # Show truncated error

        self.financial_table.setItem(0, 0, de_item)
        self.financial_table.setItem(0, 1, de_val_item)
        self.financial_table.setItem(0, 2, de_chg_item)

        # USD/CAD rate (Row 1)
        usd_item = create_item("USD/CAD Exchange Rate")
        usd_val_item = create_item("N/A", Qt.AlignRight | Qt.AlignVCenter)
        usd_chg_item = create_item("N/A", Qt.AlignCenter) # No change data here
        if 'USDCAD' in self.financial_data.get('forex', {}):
            usdcad_data = self.financial_data['forex']['USDCAD']
            if 'error' not in usdcad_data and 'rate' in usdcad_data:
                rate = usdcad_data.get('rate')
                if rate is not None:
                     usd_val_item.setText(f"{rate:.4f}")
            elif 'error' in usdcad_data:
                 usd_val_item.setText("Error")
                 usd_chg_item.setText(usdcad_data['error'][:15]+"...")

        self.financial_table.setItem(1, 0, usd_item)
        self.financial_table.setItem(1, 1, usd_val_item)
        self.financial_table.setItem(1, 2, usd_chg_item)

        # BTC in CAD (Row 2)
        btc_item = create_item("Bitcoin Price (CAD)")
        btc_val_item = create_item("N/A", Qt.AlignRight | Qt.AlignVCenter)
        btc_chg_item = create_item("N/A", Qt.AlignCenter) # No change data here
        if 'BTC' in self.financial_data.get('crypto', {}):
            btc_data = self.financial_data['crypto']['BTC']
            if 'error' not in btc_data and 'price_cad' in btc_data:
                price = btc_data.get('price_cad')
                if price is not None:
                     btc_val_item.setText(format_currency(price))
            elif 'error' in btc_data:
                 btc_val_item.setText("Error")
                 btc_chg_item.setText(btc_data['error'][:15]+"...")

        self.financial_table.setItem(2, 0, btc_item)
        self.financial_table.setItem(2, 1, btc_val_item)
        self.financial_table.setItem(2, 2, btc_chg_item)

        # Canola price (Row 3)
        can_item = create_item("Canola (CAD/bu)")
        can_val_item = create_item("N/A", Qt.AlignRight | Qt.AlignVCenter)
        can_chg_item = create_item("N/A", Qt.AlignCenter) # Source shown here instead
        if 'CANOLA' in self.financial_data.get('commodities', {}):
            canola_data = self.financial_data['commodities']['CANOLA']
            if 'error' not in canola_data and 'price_cad_bu' in canola_data:
                price = canola_data.get('price_cad_bu')
                source = canola_data.get('source', '')
                if price is not None:
                     can_val_item.setText(format_currency(price, decimals=2))
                     can_chg_item.setText(source) # Show source
            elif 'error' in canola_data:
                 can_val_item.setText("Error")
                 can_chg_item.setText(canola_data['error'][:15]+"...")

        self.financial_table.setItem(3, 0, can_item)
        self.financial_table.setItem(3, 1, can_val_item)
        self.financial_table.setItem(3, 2, can_chg_item)

        # Wheat price (Row 4)
        wht_item = create_item("Wheat (CAD/bu)")
        wht_val_item = create_item("N/A", Qt.AlignRight | Qt.AlignVCenter)
        wht_chg_item = create_item("N/A", Qt.AlignCenter) # Source shown here instead
        if 'WHEAT' in self.financial_data.get('commodities', {}):
            wheat_data = self.financial_data['commodities']['WHEAT']
            if 'error' not in wheat_data and 'price_cad_bu' in wheat_data:
                price = wheat_data.get('price_cad_bu')
                source = wheat_data.get('source', '')
                if price is not None:
                     wht_val_item.setText(format_currency(price, decimals=2))
                     wht_chg_item.setText(source) # Show source
            elif 'error' in wheat_data:
                 wht_val_item.setText("Error")
                 wht_chg_item.setText(wheat_data['error'][:15]+"...")

        self.financial_table.setItem(4, 0, wht_item)
        self.financial_table.setItem(4, 1, wht_val_item)
        self.financial_table.setItem(4, 2, wht_chg_item)

        self.financial_table.resizeRowsToContents()
        self.financial_table.setSortingEnabled(True)


    def update_stock_chart(self):
        """Update the John Deere stock chart with real data"""
        if not hasattr(self, 'stock_chart_canvas') or not self.stock_chart_canvas:
             self.logger.warning("update_stock_chart: Skipping, canvas missing.")
             return

        self.stock_chart_canvas.axes.clear() # Clear previous plot

        if 'DE' not in self.financial_data.get('stocks', {}):
             self.stock_chart_canvas.axes.text(0.5, 0.5, "Stock Data Unavailable", ha='center', va='center')
             self.stock_chart_canvas.draw()
             return

        stock_data = self.financial_data['stocks']['DE']

        if 'error' in stock_data:
            self.logger.warning(f"Cannot plot stock chart due to error: {stock_data['error']}")
            self.stock_chart_canvas.axes.text(0.5, 0.5, f"Stock Data Error:\n{stock_data['error']}", ha='center', va='center', color='red')
            self.stock_chart_canvas.draw()
            return

        # Check if we have candle data
        if 'candles' not in stock_data or stock_data['candles'].get('s') != 'ok':
            self.logger.warning("No historical stock candle data available for plotting.")
            self.stock_chart_canvas.axes.text(0.5, 0.5, "No Historical Data", ha='center', va='center')
            # Still plot current price if available
            if 'quote' in stock_data and stock_data['quote'].get('c') is not None:
                 current_price = stock_data['quote']['c']
                 self.stock_chart_canvas.axes.plot([datetime.now()], [current_price], 'ro', markersize=8, label=f"Current: ${current_price:.2f}")
                 self.stock_chart_canvas.axes.legend()
            self.stock_chart_canvas.draw()
            return

        try:
            # Extract the candle data
            candles = stock_data['candles']
            timestamps = candles.get('t', [])
            close_prices = candles.get('c', [])
            open_prices = candles.get('o', []) # Get open prices too if available
            high_prices = candles.get('h', [])
            low_prices = candles.get('l', [])

            if not timestamps or not close_prices:
                 self.logger.warning("Empty candle data received.")
                 self.stock_chart_canvas.axes.text(0.5, 0.5, "Empty Historical Data", ha='center', va='center')
                 self.stock_chart_canvas.draw()
                 return

            # Convert timestamps to dates
            dates = [datetime.fromtimestamp(ts) for ts in timestamps]

            # Create a pandas DataFrame for easier handling
            df = pd.DataFrame({
                'date': pd.to_datetime(dates), # Ensure datetime type
                'close': close_prices,
                'open': open_prices if len(open_prices) == len(dates) else [np.nan]*len(dates), # Handle missing data
                'high': high_prices if len(high_prices) == len(dates) else [np.nan]*len(dates),
                'low': low_prices if len(low_prices) == len(dates) else [np.nan]*len(dates),
            }).set_index('date') # Set date as index for potential resampling

            # Plot the closing price line
            self.stock_chart_canvas.axes.plot(df.index, df['close'], color='royalblue', linewidth=1.5, label='Close Price')

            # Add current price point from quote data
            if 'quote' in stock_data and stock_data['quote'].get('c') is not None:
                 current_price = stock_data['quote']['c']
                 # Plot slightly offset or ensure date matches last candle date
                 last_date = df.index[-1]
                 self.stock_chart_canvas.axes.plot(last_date, current_price, 'o', color='red', markersize=6, label=f"Current: ${current_price:.2f}")

            # Customize chart
            self.stock_chart_canvas.axes.set_title("John Deere (DE) Stock Price (1 Year)", fontsize=10)
            # self.stock_chart_canvas.axes.set_xlabel("Date", fontsize=9) # Often redundant with auto-formatting
            self.stock_chart_canvas.axes.set_ylabel("Price ($)", fontsize=9)
            self.stock_chart_canvas.axes.grid(True, linestyle='--', alpha=0.6)
            self.stock_chart_canvas.axes.tick_params(axis='x', labelsize=8)
            self.stock_chart_canvas.axes.tick_params(axis='y', labelsize=8)
            self.stock_chart_canvas.axes.legend(fontsize=8)

            # Format x-axis dates nicely
            self.stock_chart_canvas.figure.autofmt_xdate()

            # Adjust layout and draw
            self.stock_chart_canvas.figure.tight_layout()
            self.stock_chart_canvas.draw()
            self.logger.debug("Stock chart updated.")

        except Exception as e:
             self.logger.error(f"Error updating stock chart: {e}", exc_info=True)
             try: # Try to show error on canvas
                  self.stock_chart_canvas.axes.clear()
                  self.stock_chart_canvas.axes.text(0.5, 0.5, "Error generating chart", ha='center', va='center', color='red')
                  self.stock_chart_canvas.draw()
             except Exception: pass # Ignore errors during error display


    def get_wind_direction(self, degrees):
        """Convert wind direction in degrees to cardinal direction"""
        if degrees is None: return ""
        try:
            directions = ["N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE",
                          "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW"]
            index = round(((float(degrees) % 360) / 22.5))
            return directions[index % 16]
        except (ValueError, TypeError):
            return ""


    def get_temp_color(self, temp):
        """Get color based on temperature in Celsius"""
        try:
            temp = float(temp)
            if temp <= -10: return '#0000FF'  # Deep blue
            elif temp <= 0: return '#00BFFF'  # Light blue
            elif temp <= 10: return '#90EE90' # Light green
            elif temp <= 20: return '#00FF00' # Green
            elif temp <= 30: return '#FFA500' # Orange
            else: return '#FF0000' # Red
        except (ValueError, TypeError):
            return '#808080' # Grey for invalid


    def get_humidity_color(self, humidity):
        """Get color based on humidity percentage"""
        try:
            humidity = float(humidity)
            if humidity <= 30: return '#FFD700'  # Gold
            elif humidity <= 50: return '#ADFF2F' # Green yellow
            elif humidity <= 70: return '#00FF00' # Green
            elif humidity <= 90: return '#00BFFF' # Light blue
            else: return '#0000FF' # Blue
        except (ValueError, TypeError):
            return '#808080' # Grey for invalid


    def get_wind_color(self, wind_speed):
        """Get color based on wind speed in m/s"""
        try:
            wind_speed = float(wind_speed)
            if wind_speed <= 1: return '#00FF00'  # Green
            elif wind_speed <= 3: return '#ADFF2F' # Green yellow
            elif wind_speed <= 6: return '#FFFF00'  # Yellow
            elif wind_speed <= 10: return '#FFA500' # Orange
            elif wind_speed <= 15: return '#FF4500' # Orange red
            else: return '#FF0000' # Red
        except (ValueError, TypeError):
            return '#808080' # Grey for invalid


    # --- BaseModule Methods ---
    def get_title(self):
        """Return the module title for display purposes."""
        return self.MODULE_DISPLAY_NAME

    def get_icon_name(self):
        """Return the icon filename for the module."""
        return self.MODULE_ICON_NAME

    def refresh(self):
        """Refresh method called from the main application when module is shown."""
        self.logger.debug("HomeModule refresh triggered.")
        # Optionally trigger a data refresh immediately when tab is selected
        # self.refresh_data()
        # Or just ensure UI is up-to-date with existing data
        self.update_ui()

    def save_state(self):
        """Save any necessary state for the home module (if any)."""
        self.logger.debug("HomeModule save_state called (no state to save).")
        pass # Nothing specific to save for this module currently

    def close(self):
        """Clean up resources before closing."""
        self.logger.debug("Closing HomeModule...")
        self.cleanup()

    def cleanup(self):
        """Stop background threads."""
        self.logger.debug("Stopping HomeModule background threads...")
        if self.weather_fetcher and self.weather_fetcher.isRunning():
            self.weather_fetcher.stop()
            self.weather_fetcher.wait()
            self.logger.debug("Weather fetcher stopped.")

        if self.financial_fetcher and self.financial_fetcher.isRunning():
            self.financial_fetcher.stop()
            self.financial_fetcher.wait()
            self.logger.debug("Financial fetcher stopped.")

        if self.scheduled_refresh and self.scheduled_refresh.isRunning():
            self.scheduled_refresh.stop()
            self.scheduled_refresh.wait()
            self.logger.debug("Scheduled refresh stopped.")
        self.logger.debug("HomeModule cleanup finished.")


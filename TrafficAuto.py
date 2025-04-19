# TrafficAuto.py - Modified
import sys
import pyautogui
import time
import pyperclip
import csv
import os
import logging
from logging.handlers import RotatingFileHandler # Added
from utils.config import Config # Added - Assuming config.py is accessible

# --- Logging Setup ---
def setup_traffic_logging(log_dir="logs", log_level_str="INFO"):
    """Configure logging specifically for this script."""
    log_file = os.path.join(log_dir, "traffic_auto.log")
    log_level = getattr(logging, log_level_str.upper(), logging.INFO)
    os.makedirs(log_dir, exist_ok=True)

    log_format = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    # Configure logger
    logger = logging.getLogger("TrafficAuto")
    logger.setLevel(log_level)
    logger.handlers.clear() # Remove default handlers if any

    # File Handler
    try:
        file_handler = RotatingFileHandler(log_file, maxBytes=1*1024*1024, backupCount=1, encoding='utf-8')
        file_handler.setFormatter(log_format)
        logger.addHandler(file_handler)
    except Exception as e:
        print(f"Error setting up file logger: {e}", file=sys.stderr) # Use print only if logger fails

    # Console Handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(log_format)
    logger.addHandler(console_handler)

    logger.info("TrafficAuto logging initialized.")
    return logger

# --- Configuration ---
try:
    config = Config()
    # Assuming these keys exist in your config setup (config.py or .env)
    images_dir = config.get('traffic_images_dir', os.path.join(os.path.dirname(os.path.abspath(__file__)), "Script Images"))
    tasks_csv_file = config.get('traffic_csv_path', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'traffic_tasks.csv'))
    log_dir_config = config.get('log_dir', 'logs')
    log_level_config = config.get('log_level', 'INFO')
except Exception as e:
    print(f"CRITICAL ERROR: Failed to load configuration. Using defaults. Error: {e}", file=sys.stderr)
    script_dir_fallback = os.path.dirname(os.path.abspath(__file__))
    images_dir = os.path.join(script_dir_fallback, "Script Images")
    tasks_csv_file = os.path.join(script_dir_fallback, 'traffic_tasks.csv')
    log_dir_config = "logs"
    log_level_config = "INFO"

# --- Initialize Logger ---
logger = setup_traffic_logging(log_dir=log_dir_config, log_level_str=log_level_config)

# --- Script Settings ---
pyautogui.PAUSE = config.get('pyautogui_pause', 0.2) # Get pause from config if desired

# --- Directory Checks ---
if not os.path.isdir(images_dir):
    logger.critical(f"Image directory not found at configured path: {images_dir}")
    # Maybe exit here if images are essential? sys.exit(1)
else:
    logger.debug(f"Using image directory: {images_dir}")

if not os.path.exists(tasks_csv_file):
    logger.critical(f"Tasks CSV file not found at configured path: {tasks_csv_file}")
    sys.exit(f"Error: Tasks file '{tasks_csv_file}' not found.")


# --- Helper Functions ---
def get_image_path(image_name):
    return os.path.join(images_dir, image_name)

def click_element(image_path, description, timeout=10, click=True, conf=0.8, region=None):
    """Clicks an element identified by an image, with logging."""
    start_time = time.time()
    location = None
    img_file = get_image_path(image_path) # Use helper to get full path

    if not os.path.exists(img_file):
        logger.error(f"Image file not found for '{description}': {img_file}")
        return None # Cannot proceed without the image file

    logger.debug(f"Attempting to find '{description}' using image: {img_file}")
    while time.time() - start_time < timeout:
        try:
            locate_args = {'confidence': conf}
            if region:
                locate_args['region'] = region

            location = pyautogui.locateCenterOnScreen(img_file, **locate_args)

        except pyautogui.ImageNotFoundException:
             # This is expected if the image isn't there, continue loop
             logger.log(5, f"ImageNotFoundException for '{description}', retrying...") # Use custom level 5 for trace/spammy debug
             location = None
        except Exception as e:
            # Log unexpected errors during location attempt
            logger.warning(f"Exception while locating '{description}' ({img_file}): {e}")
            location = None

        if location:
            logger.info(f"Found '{description}' at {location}.")
            if click:
                try:
                    pyautogui.click(location)
                    logger.debug(f"Clicked '{description}' at {location}.")
                    time.sleep(0.1) # Small pause after click
                except Exception as e:
                    logger.error(f"Failed to click '{description}' at {location}: {e}")
                    return None # Return None if click fails
            return location # Return location whether clicked or not

        time.sleep(0.5) # Wait before retrying

    logger.warning(f"Timeout: Could not find '{description}' ({img_file}) within {timeout} seconds.")
    return None

def click_and_type(image_path, text_to_type, description, timeout=10, conf=0.8, region=None):
    """Clicks an element and types text, with logging."""
    location = click_element(image_path, description, timeout, click=True, conf=conf, region=region)
    if location:
        logger.debug(f"Typing '{text_to_type}' into '{description}'.")
        try:
            # Using pyperclip might be more reliable sometimes
            # pyperclip.copy(text_to_type)
            # pyautogui.hotkey('ctrl', 'v')
            # time.sleep(0.1)
            pyautogui.typewrite(str(text_to_type), interval=0.05) # Type slowly
            time.sleep(0.1)
            return True
        except Exception as e:
            logger.error(f"Failed to type '{text_to_type}' into '{description}': {e}")
            return False
    else:
        # Error already logged by click_element
        return False

# --- Main Workflow Function ---
def process_traffic_entry(customer_name, stock_item):
    """Automates a single traffic entry."""
    logger.info(f"--- Processing entry: Customer='{customer_name}', Stock='{stock_item}' ---")

    # Step 1: Click "New Traffic Button"
    logger.info("Step 1: Click New Traffic Button")
    if not click_element("new_traffic.png", "New Traffic Button", timeout=15):
        return "Failed: Could not find 'New Traffic' button."
    time.sleep(0.5)

    # Step 2: Click "Customer Lookup"
    logger.info("Step 2: Click Customer Lookup")
    if not click_element("customer_lookup.png", "Customer Lookup Button", timeout=10):
        return "Failed: Could not find 'Customer Lookup' button."
    time.sleep(0.5)

    # Step 3: Enter customer name and click search
    logger.info(f"Step 3: Search for customer '{customer_name}'")
    if not click_and_type("customer_name_field.png", customer_name, "Customer Name Field", timeout=10):
         # Don't return failure yet, maybe the field was already focused or image slightly off
         # Try typing directly if click_and_type failed based on image
         logger.warning("Could not click Customer Name Field based on image, attempting direct type.")
         try:
             pyautogui.typewrite(customer_name, interval=0.05)
         except Exception as e:
              logger.error(f"Direct type failed for customer name: {e}")
              return f"Failed: Could not enter customer name '{customer_name}'."
         time.sleep(0.2) # Pause after typing

    if not click_element("search_button.png", "Search Button", timeout=10):
        return "Failed: Could not find 'Search' button."
    time.sleep(1) # Wait for search results

    # Step 4: Click the result (assuming first result is correct - risky!)
    logger.info("Step 4: Click Customer Result (assuming first match)")
    # Note: This assumes the result appears at a predictable location relative to the search field,
    # or uses an image specific to the result row. Using a fixed offset or generic row image is fragile.
    # If possible, use an image unique to the customer row or OCR.
    # Example using a generic 'select_result.png' which might be the first row:
    if not click_element("select_result.png", "Customer Result Row", timeout=10):
         # Alternative: try double-clicking near where the result should be
         logger.warning("Could not find specific 'Select Result' image, trying double click near field.")
         # This needs coordinates relative to the search field or a fixed screen area
         # Example: search_btn_loc = click_element("search_button.png", "Search Button", click=False)
         # if search_btn_loc: pyautogui.doubleClick(search_btn_loc.x, search_btn_loc.y + 50) # Guessing position
         # else: return "Failed: Cannot determine where to click customer result."
         # For now, just log failure if image isn't found
         return "Failed: Could not find or click customer result row."
    time.sleep(0.5)

    # Step 5: Click Save (after CONV01 appears - assumes this is a confirmation)
    logger.info("Step 5: Click Save (after CONV01)")
    # Need an image for 'conv01.png' or whatever indicates the selection worked
    if not click_element("conv01.png", "CONV01 Confirmation", timeout=15, click=False): # Find confirmation first
        logger.warning("Did not find CONV01 confirmation, proceeding to Save anyway.")
        # return "Failed: Did not find CONV01 confirmation after selecting customer."

    if not click_element("save.png", "Save Button (after CONV01)", timeout=10):
        return "Failed: Could not find 'Save' button after entering customer."
    time.sleep(2) # Wait for save/screen update

    # Step 6: Enter stock number
    logger.info(f"Step 6: Enter stock number '{stock_item}'")
    if not click_and_type("stock_number.png", stock_item, "Stock Number Field", timeout=15):
        return f"Failed: Could not enter stock number '{stock_item}'."
    time.sleep(0.3)

    # Step 7: Click Save (after stock number)
    logger.info("Step 7: Click Save (after stock number)")
    if not click_element("save.png", "Save Button (after stock number)", timeout=10):
        return "Failed: Could not find 'Save' button after entering stock number."
    time.sleep(0.5) # Wait for save

    # Step 8: Click Pending dropdown and press 'c' for Comp/Pay
    logger.info("Step 8: Click Pending dropdown and press 'c'")
    if not click_element("pending.png", "Pending Dropdown", timeout=10):
        return "Failed: Could not find 'Pending' dropdown."
    time.sleep(0.3)
    try:
        pyautogui.press("c")
        logger.info("Pressed 'c' for Comp/Pay status.")
    except Exception as e:
         logger.error(f"Failed to press 'c': {e}")
         return "Failed: Could not set status to Comp/Pay."
    time.sleep(0.3)

    # Step 9: Enter 'BRITRK' in Trucker field
    logger.info("Step 9: Enter 'BRITRK' in Trucker field")
    if not click_and_type("trucker.png", "BRITRK", "Trucker Field", timeout=10):
        return "Failed: Could not enter 'BRITRK' in Trucker field."

    # Step 10: Click Final Save
    logger.info("Step 10: Click Final Save")
    # Assuming the final save button also uses 'save.png' - might need a different image 'final_save.png'
    if not click_element("save.png", "Final Save Button", timeout=15):
        return "Failed: Could not find final 'Save' button."
    time.sleep(2) # Wait longer for final save/window close

    logger.info(f"--- Successfully processed entry for Customer='{customer_name}', Stock='{stock_item}' ---")
    return "Success"


# --- Main Loop ---
def main():
    logger.info("Starting TrafficAuto script...")
    logger.info(f"Reading tasks from: {tasks_csv_file}")

    tasks = []
    try:
        with open(tasks_csv_file, mode='r', newline='', encoding='utf-8-sig') as file: # Use utf-8-sig to handle potential BOM
            reader = csv.DictReader(file)
            if 'CustomerName' not in reader.fieldnames or 'StockItem' not in reader.fieldnames:
                 logger.critical(f"CSV file '{tasks_csv_file}' must contain 'CustomerName' and 'StockItem' columns.")
                 sys.exit(1)
            tasks = list(reader)
    except FileNotFoundError:
        logger.critical(f"Tasks CSV file not found: {tasks_csv_file}")
        sys.exit(1)
    except Exception as e:
        logger.critical(f"Error reading CSV file {tasks_csv_file}: {e}")
        sys.exit(1)

    if not tasks:
        logger.warning("No tasks found in the CSV file.")
        sys.exit(0)

    logger.info(f"Found {len(tasks)} tasks to process.")
    results = []

    # Small delay before starting automation
    logger.info("Starting automation in 5 seconds...")
    time.sleep(5)

    for i, task in enumerate(tasks):
        customer_name = task.get('CustomerName', '').strip()
        stock_item = task.get('StockItem', '').strip()

        if not customer_name or not stock_item:
            logger.warning(f"Skipping task {i+1}: Missing CustomerName or StockItem.")
            results.append({'CustomerName': customer_name, 'StockItem': stock_item, 'Status': 'Skipped - Missing Data'})
            continue

        logger.info(f"--- Starting Task {i+1}/{len(tasks)} ---")
        status = process_traffic_entry(customer_name, stock_item)
        results.append({'CustomerName': customer_name, 'StockItem': stock_item, 'Status': status})
        logger.info(f"--- Finished Task {i+1}/{len(tasks)} with status: {status} ---")
        time.sleep(1) # Pause between entries

    # --- Output Results ---
    results_csv_file = os.path.join(os.path.dirname(tasks_csv_file), 'traffic_results.csv')
    logger.info(f"Processing complete. Writing results to: {results_csv_file}")
    try:
        with open(results_csv_file, mode='w', newline='', encoding='utf-8') as file:
            if results: # Only write if there are results
                 writer = csv.DictWriter(file, fieldnames=['CustomerName', 'StockItem', 'Status'])
                 writer.writeheader()
                 writer.writerows(results)
            else:
                 logger.warning("No results to write.")
        logger.info("Results successfully saved.")
    except Exception as e:
        logger.error(f"Failed to write results CSV: {e}")

    logger.info("TrafficAuto script finished.")

if __name__ == "__main__":
    main()
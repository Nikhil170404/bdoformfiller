from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException, ElementNotInteractableException
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from typing import List, Dict, Optional, Union, Any, cast, Tuple
import pandas as pd
import numpy as np
import logging
import time
import random
import os
import re

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('form_automation.log'),
        logging.StreamHandler()
    ]
)

class FormFiller:
    def __init__(self, url: str, excel_path: str):
        """Initialize the form filler"""
        self.url = url
        self.excel_path = excel_path
        self.driver: Optional[webdriver.Edge] = None
        self.wait: Optional[WebDriverWait] = None
        self.data: Optional[pd.DataFrame] = None
        self.actions: Optional[ActionChains] = None
        
    def setup_driver(self) -> None:
        """Initialize Edge driver with custom options"""
        try:
            edge_options = Options()
            edge_options.add_argument('--start-maximized')
            edge_options.add_argument('--disable-blink-features=AutomationControlled')
            edge_options.add_argument('--disable-notifications')
            edge_options.add_argument('--inprivate')
            edge_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            edge_options.add_experimental_option('useAutomationExtension', False)
            
            service = Service(EdgeChromiumDriverManager().install())
            self.driver = webdriver.Edge(service=service, options=edge_options)
            self.wait = WebDriverWait(self.driver, 10)
            self.actions = ActionChains(self.driver)
            logging.info("Browser setup completed successfully")
            
        except Exception as e:
            logging.error(f"Error setting up browser: {e}")
            if self.driver:
                self.driver.quit()
            self.driver = None
            raise

    def normalize_string(self, text: str) -> str:
        """Normalize string for comparison"""
        return re.sub(r'[^a-zA-Z0-9]', '', text.lower())

    def strings_similar(self, str1: str, str2: str) -> bool:
        """Check if two strings are similar"""
        norm1 = self.normalize_string(str1)
        norm2 = self.normalize_string(str2)
        return norm1 in norm2 or norm2 in norm1

    def find_field_by_multiple_strategies(self, label: str) -> Optional[WebElement]:
        """Find form field using multiple strategies"""
        if not self.driver or not self.wait:
            return None

        # Only create strategies if driver exists
        driver = self.driver  # Local reference to avoid multiple checks
        if not driver:
            return None

        strategies = [
            # By label with 'for' attribute
            (By.CSS_SELECTOR, f"label[for*='{label}'] + input, label[for*='{label}'] + select, label[for*='{label}'] + textarea"),
            
            # By label text
            (By.XPATH, f"//label[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{label.lower()}')]//following::input[1]"),
            
            # By placeholder
            (By.CSS_SELECTOR, f"input[placeholder*='{label}'], textarea[placeholder*='{label}']"),
            
            # By name or id
            (By.CSS_SELECTOR, f"input[name*='{label}'], input[id*='{label}'], textarea[name*='{label}'], textarea[id*='{label}']"),
            
            # By aria-label
            (By.CSS_SELECTOR, f"input[aria-label*='{label}'], textarea[aria-label*='{label}']"),
            
            # By class containing label
            (By.CSS_SELECTOR, f"input[class*='{label}'], textarea[class*='{label}']"),
            
            # By nearby text
            (By.XPATH, f"//*[contains(text(), '{label}')]/following::input[1]"),
            
            # By data-testid or similar attributes
            (By.CSS_SELECTOR, f"input[data-testid*='{label}'], input[data-test*='{label}'], input[data-qa*='{label}']"),
        ]

        # Try each strategy
        for by, selector in strategies:
            try:
                field = driver.find_element(by, selector)
                if field and field.is_displayed() and field.is_enabled():
                    return field
            except (NoSuchElementException, ElementNotInteractableException):
                continue

        # If no exact match found, try fuzzy matching on all input fields
        try:
            all_inputs = driver.find_elements(By.CSS_SELECTOR, "input:not([type='hidden']), textarea, select")
            for input_field in all_inputs:
                # Check various attributes for similarity
                for attr in ['name', 'id', 'placeholder', 'aria-label']:
                    attr_value = input_field.get_attribute(attr)
                    if attr_value and self.strings_similar(label, attr_value):
                        return input_field
                
                # Check nearby labels
                try:
                    nearby_labels = driver.find_elements(By.XPATH, f"//label[ancestor::*[count(.|//{input_field.tag_name}[@id='{input_field.get_attribute('id')}'])=count(//{input_field.tag_name}[@id='{input_field.get_attribute('id')}'])]]")
                    for nearby_label in nearby_labels:
                        if self.strings_similar(label, nearby_label.text):
                            return input_field
                except:
                    continue
                    
        except Exception as e:
            logging.warning(f"Error in fuzzy matching: {e}")

        return None

    def load_excel_data(self) -> bool:
        """Load data from Excel template"""
        try:
            self.data = pd.read_excel(self.excel_path)
            if self.data is None:
                logging.error("Excel file could not be loaded")
                return False
                
            if isinstance(self.data, pd.DataFrame) and self.data.empty:
                logging.error("Excel file is empty")
                return False
                
            row_count = len(self.data) if isinstance(self.data, pd.DataFrame) else 0
            if row_count == 0:
                logging.error("No data rows found in Excel file")
                return False
                
            logging.info(f"Loaded {row_count} rows of data from Excel")
            return True
            
        except Exception as e:
            logging.error(f"Error loading Excel data: {e}")
            return False

    def safe_send_keys(self, field: WebElement, value: str) -> None:
        """Safely send keys to a field with multiple attempts"""
        max_attempts = 3
        for attempt in range(max_attempts):
            try:
                # Try clicking the field first
                field.click()
                time.sleep(random.uniform(0.1, 0.3))
                
                # Clear the field
                field.clear()
                time.sleep(random.uniform(0.1, 0.3))
                
                # Try different input methods
                if attempt == 0:
                    # Normal typing
                    for char in value:
                        field.send_keys(char)
                        time.sleep(random.uniform(0.05, 0.15))
                elif attempt == 1:
                    # JavaScript injection
                    if self.driver:
                        self.driver.execute_script("arguments[0].value = arguments[1]", field, value)
                else:
                    # Action chains
                    if self.actions:
                        self.actions.move_to_element(field).click().send_keys(value).perform()
                
                # Verify the input
                actual_value = field.get_attribute('value')
                if actual_value == value:
                    return
                    
            except Exception as e:
                logging.warning(f"Attempt {attempt + 1} failed: {e}")
                time.sleep(0.5)
                
        raise ValueError(f"Failed to input value after {max_attempts} attempts")

    def fill_field(self, field: WebElement, value: Any, field_name: str) -> None:
        """Fill a form field with value using human-like interaction"""
        try:
            # Convert value to string, handling None and nan
            str_value = str(value) if pd.notna(value) and value is not None else ""
            if not str_value:
                return
                
            field_type = field.get_attribute('type')
            
            if field_type == 'select-one':
                # Handle dropdown
                select_handled = False
                options = field.find_elements(By.TAG_NAME, 'option')
                
                for option in options:
                    if self.strings_similar(option.text, str_value):
                        option.click()
                        select_handled = True
                        break
                        
                if not select_handled and options:
                    # If no match found but options exist, try to set value directly
                    if self.driver:
                        self.driver.execute_script("arguments[0].value = arguments[1]", field, str_value)
                        
            elif field_type == 'checkbox':
                # Handle checkbox
                current_state = field.is_selected()
                desired_state = str_value.lower() in ('true', 'yes', '1', 'on')
                if current_state != desired_state:
                    field.click()
                    
            elif field_type == 'radio':
                # Handle radio buttons
                if str_value.lower() in ('true', 'yes', '1', 'on'):
                    field.click()
                    
            else:
                # Handle text input fields
                self.safe_send_keys(field, str_value)
                
            logging.info(f"Filled field '{field_name}' with value: {str_value}")
            
        except Exception as e:
            logging.error(f"Error filling field '{field_name}': {e}")

    def submit_form(self) -> bool:
        """Submit the form with multiple strategies"""
        if not self.driver or not self.wait or not self.actions:
            return False

        driver = self.driver  # Local reference
        actions = self.actions  # Local reference
        wait = self.wait  # Local reference

        submit_strategies = [
            # Standard submit buttons
            (By.CSS_SELECTOR, "button[type='submit'], input[type='submit']"),
            # Common submit button classes
            (By.CSS_SELECTOR, ".submit-button, .submitButton, .submit"),
            # Buttons containing 'submit' text
            (By.XPATH, "//button[contains(translate(text(), 'SUBMIT', 'submit'), 'submit')]"),
            # Submit images
            (By.CSS_SELECTOR, "input[type='image'][name*='submit']"),
            # Form submit via JavaScript
            (None, None)  # Special case for JavaScript submit
        ]

        for by, selector in submit_strategies:
            try:
                if by and selector:
                    submit_button = wait.until(EC.element_to_be_clickable((by, selector)))
                    actions.move_to_element(submit_button).click().perform()
                else:
                    # Try JavaScript form submit
                    driver.execute_script("document.forms[0].submit()")
                        
                time.sleep(random.uniform(2.0, 3.0))
                return True
                
            except Exception as e:
                logging.warning(f"Submit strategy failed: {e}")
                continue

        return False

    def fill_form(self) -> None:
        """Fill form with data from Excel"""
        if not self.driver or not self.wait or self.data is None:
            return
            
        try:
            # Navigate to form URL
            self.driver.get(self.url)
            time.sleep(2)  # Wait for page load
            
            # Process each row in the Excel file
            df = cast(pd.DataFrame, self.data)
            total_rows = len(df)
            
            for idx in range(total_rows):
                row = df.iloc[idx]
                logging.info(f"Processing form submission {idx + 1}")
                
                # Fill each field
                for field_name in df.columns:
                    value = row[field_name]
                    if pd.isna(value):
                        continue
                        
                    # Remove asterisk from required field labels
                    if isinstance(field_name, str):
                        clean_field_name = field_name.replace(' *', '')
                    else:
                        clean_field_name = str(field_name)
                    
                    # Find and fill the field
                    field = self.find_field_by_multiple_strategies(clean_field_name)
                    if field:
                        self.fill_field(field, value, clean_field_name)
                        time.sleep(random.uniform(0.3, 0.7))
                    else:
                        logging.warning(f"Could not find field: {clean_field_name}")
                
                # Submit the form
                if self.submit_form():
                    logging.info(f"Submitted form {idx + 1}")
                    
                    # If there are more forms to fill, navigate back to the form
                    if idx < total_rows - 1:
                        self.driver.get(self.url)
                        time.sleep(2)
                else:
                    logging.error(f"Failed to submit form {idx + 1}")
                
        except Exception as e:
            logging.error(f"Error in form filling process: {e}")
        finally:
            if self.driver:
                self.driver.quit()

def main() -> None:
    """Main function to run the form filler"""
    try:
        url = input("Enter the form URL: ")
        excel_path = input("Enter the path to Excel template: ")
        
        if not os.path.exists(excel_path):
            print("\nError: Excel file not found!")
            return
            
        print("\nStarting form filling process...")
        filler = FormFiller(url, excel_path)
        
        if filler.load_excel_data():
            filler.setup_driver()
            filler.fill_form()
            print("\nForm filling process completed! Check form_automation.log for details.")
        else:
            print("\nError: Could not load Excel data. Please check the file and try again.")
        
    except Exception as e:
        logging.error(f"Error in main: {e}")
        print("\nAn error occurred. Please check form_automation.log for details.")

if __name__ == "__main__":
    main()

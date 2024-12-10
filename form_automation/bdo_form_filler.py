"""
BDO Form Automation Script with enhanced element detection and interaction
"""

import os
import time
import random
import logging
import subprocess
import threading
from dataclasses import dataclass
from typing import Optional, List, Dict, Any, Union, Tuple
from urllib.parse import urlparse

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, ElementNotInteractableException,
    StaleElementReferenceException
)
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def kill_chrome() -> None:
    """Kill Chrome processes"""
    try:
        subprocess.run(['taskkill', '/F', '/IM', 'chrome.exe'], 
                      stdout=subprocess.DEVNULL, 
                      stderr=subprocess.DEVNULL,
                      check=False)
        time.sleep(2)
    except Exception as e:
        logging.error(f"Error killing Chrome: {e}")

@dataclass
class ElementLocator:
    """Class to represent a single element locator with multiple strategies"""
    name: str
    strategies: List[tuple[str, str]]
    
    def __post_init__(self):
        # Ensure strategies are valid
        valid_by = [By.ID, By.NAME, By.CLASS_NAME, By.CSS_SELECTOR, By.XPATH]
        for by, _ in self.strategies:
            if by not in valid_by:
                raise ValueError(f"Invalid locator strategy: {by}")

class FormLocators:
    """Class to store form element locators with multiple strategies"""
    FIRST_NAME = ElementLocator("First Name", [
        (By.NAME, "firstName"),
        (By.ID, "firstName"),
        (By.CSS_SELECTOR, "input[name='firstName']"),
        (By.XPATH, "//input[@placeholder='First Name']")
    ])
    
    LAST_NAME = ElementLocator("Last Name", [
        (By.NAME, "lastName"),
        (By.ID, "lastName"),
        (By.CSS_SELECTOR, "input[name='lastName']"),
        (By.XPATH, "//input[@placeholder='Last Name']")
    ])
    
    MOBILE_NUMBER = ElementLocator("Mobile Number", [
        (By.NAME, "mobileNumber"),
        (By.ID, "mobileNumber"),
        (By.CSS_SELECTOR, "input[name='mobileNumber']"),
        (By.XPATH, "//input[contains(@placeholder, 'Mobile')]")
    ])
    
    EMAIL = ElementLocator("Email", [
        (By.NAME, "email"),
        (By.ID, "email"),
        (By.CSS_SELECTOR, "input[type='email']"),
        (By.XPATH, "//input[@type='email']")
    ])
    
    COUNTRY_CODE = ElementLocator("Country Code", [
        (By.ID, "codrp"),
        (By.NAME, "countryCode"),
        (By.CSS_SELECTOR, "select#codrp"),
        (By.XPATH, "//select[contains(@id, 'codrp')]")
    ])
    
    NEXT_BUTTON = ElementLocator("Next Button", [
        (By.CSS_SELECTOR, "button.btn.btn-primary.button.next.submit-btn"),
        (By.XPATH, "//button[contains(@class, 'next')]"),
        (By.XPATH, "//button[contains(text(), 'Next')]"),
        (By.CSS_SELECTOR, ".btn-primary.next")
    ])
    
    OK_BUTTON = ElementLocator("OK Button", [
        (By.XPATH, "//button[contains(text(), 'OK')]"),
        (By.XPATH, "//button[contains(text(), 'Ok')]"),
        (By.CSS_SELECTOR, "button.ok-button"),
        (By.CSS_SELECTOR, ".modal-footer button")
    ])

class BDOFormFiller:
    """Enhanced BDO form automation with robust element handling"""
    
    def __init__(self) -> None:
        """Initialize the form filler with proper type handling"""
        self.url = "https://www.apply.bdo.com.ph/newntb/"
        self.driver: Optional[webdriver.Chrome] = None
        self.wait: Optional[WebDriverWait] = None
        self.locators = FormLocators()
        
        try:
            options = webdriver.ChromeOptions()
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-gpu')
            options.add_argument('--disable-blink-features=AutomationControlled')
            options.add_argument('--disable-extensions')
            options.add_argument('--disable-notifications')
            options.add_argument('--disable-popup-blocking')
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)
            
            service = Service()
            self.driver = webdriver.Chrome(service=service, options=options)
            self.wait = WebDriverWait(self.driver, 20)  # Increased timeout
            self.driver.maximize_window()
            
        except Exception as e:
            logging.error(f"Error initializing driver: {e}")
            if self.driver:
                self.driver.quit()
            raise

    def setup_logging(self) -> None:
        """Configure detailed logging"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('bdo_form_automation.log'),
                logging.StreamHandler()
            ]
        )

    def setup_driver(self) -> bool:
        """Set up Chrome WebDriver with proper initialization"""
        try:
            # Kill any existing Chrome processes first
            kill_chrome()
            time.sleep(2)
            
            options = webdriver.ChromeOptions()
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-gpu')
            options.add_argument('--disable-blink-features=AutomationControlled')
            options.add_argument('--start-maximized')
            
            # Use default profile to maintain login state
            options.add_argument(f"--user-data-dir=C:\\Users\\{os.getenv('USERNAME')}\\AppData\\Local\\Google\\Chrome\\User Data")
            options.add_argument("--profile-directory=Default")
            
            # Disable automation flags
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)
            
            service = Service()
            driver = webdriver.Chrome(service=service, options=options)
            
            if not isinstance(driver, webdriver.Chrome):
                logging.error("Failed to initialize Chrome WebDriver")
                return False
                
            self.driver = driver
            self.wait = WebDriverWait(self.driver, 20)
            
            # Navigate to URL
            self.driver.get(self.url)
            time.sleep(2)
            
            return True
            
        except Exception as e:
            logging.error(f"Error setting up WebDriver: {e}")
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
            return False
            
    def wait_for_element(self, locator: ElementLocator) -> Optional[WebElement]:
        """Wait for element with proper type checking"""
        if not isinstance(self.driver, webdriver.Chrome) or not self.wait:
            return None
            
        try:
            for strategy in locator.strategies:
                try:
                    element = self.wait.until(
                        EC.presence_of_element_located((strategy[0], strategy[1]))
                    )
                    if element and element.is_displayed():
                        return element
                except:
                    continue
            return None
        except Exception as e:
            logging.error(f"Error finding element {locator.name}: {e}")
            return None
            
    def ensure_element_interactable(self, element: WebElement) -> bool:
        """Ensure element is truly interactable with proper driver checks"""
        if not isinstance(self.driver, webdriver.Chrome):
            logging.error("WebDriver not properly initialized")
            return False
            
        try:
            # Scroll element into center view
            self.driver.execute_script(
                "arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});",
                element
            )
            time.sleep(1)
            
            # Check if element is truly visible and interactable
            if not element.is_displayed() or not element.is_enabled():
                return False
                
            # Try to move mouse to element
            actions = ActionChains(self.driver)
            actions.move_to_element(element).perform()
            time.sleep(0.5)
            
            return True
        except Exception as e:
            logging.error(f"Error ensuring element interactability: {e}")
            return False
            
    def fill_input(self, locator: ElementLocator, text: str) -> bool:
        """Enhanced input filling with proper type checking"""
        if not isinstance(self.driver, webdriver.Chrome):
            logging.error("WebDriver not properly initialized")
            return False
            
        element = self.wait_for_element(locator)
        if not element:
            return False
            
        try:
            if not self.ensure_element_interactable(element):
                return False
                
            # Clear existing text
            try:
                element.clear()
                time.sleep(0.5)
            except:
                element.send_keys(Keys.CONTROL + "a")
                element.send_keys(Keys.DELETE)
                time.sleep(0.5)
                
            # Type text with random delays
            for char in text:
                element.send_keys(char)
                time.sleep(random.uniform(0.1, 0.3))
                
            # Verify input
            actual_value = element.get_attribute('value')
            if actual_value != text:
                self.driver.execute_script(
                    "arguments[0].value = arguments[1];",
                    element, text
                )
                
            return True
        except Exception as e:
            logging.error(f"Error filling input {locator.name}: {e}")
            return False
            
    def select_country_code(self, code: str = "+44") -> bool:
        """Enhanced country code selection with proper type checking"""
        if not isinstance(self.driver, webdriver.Chrome):
            logging.error("WebDriver not properly initialized")
            return False
            
        element = self.wait_for_element(self.locators.COUNTRY_CODE)
        if not element:
            return False
            
        try:
            if not self.ensure_element_interactable(element):
                return False
                
            # Try multiple selection methods
            methods = [
                lambda: Select(element).select_by_value(code),
                lambda: (element.click(), time.sleep(0.5), element.send_keys(code), element.send_keys(Keys.ENTER)),
                lambda: self.driver.execute_script(f"arguments[0].value = '{code}';", element)
            ]
            
            for method in methods:
                try:
                    method()
                    time.sleep(1)
                    # Verify selection
                    value = element.get_attribute('value')
                    if value and code in value:
                        return True
                except:
                    continue
                    
            return False
        except Exception as e:
            logging.error(f"Error selecting country code: {e}")
            return False
            
    def click_button(self, locator: ElementLocator) -> bool:
        """Enhanced button clicking with multiple strategies"""
        try:
            element = self.wait_for_element(locator)
            if not element:
                return False
                
            # Type guard for driver
            assert self.driver is not None

            if not self.ensure_element_interactable(element):
                return False

            # Try multiple click methods
            methods = [
                lambda: element.click(),
                lambda: self.driver.execute_script("arguments[0].click();", element),
                lambda: ActionChains(self.driver).move_to_element(element).click().perform(),
                lambda: element.send_keys(Keys.RETURN)
            ]

            for method in methods:
                try:
                    method()
                    time.sleep(1)
                    return True
                except:
                    continue

            return False
        except Exception as e:
            logging.error(f"Error clicking button {locator.name}: {e}")
            return False

    def handle_popup(self) -> bool:
        """Handle the confirmation popup"""
        try:
            # Wait for popup OK button
            ok_button = self.wait.until(
                EC.presence_of_element_located((By.ID, "personalbtn"))
            )
            if ok_button and ok_button.is_displayed():
                try:
                    ok_button.click()
                except:
                    self.driver.execute_script("arguments[0].click();", ok_button)
                time.sleep(0.5)
                return True
            return False
        except Exception as e:
            logging.error(f"Error handling popup: {e}")
            return False

    def fill_form(self, phone_number: str) -> bool:
        """Fast form filling with exact selectors and proper delays"""
        if not isinstance(self.driver, webdriver.Chrome):
            return False

        try:
            # Fill form using JavaScript with exact IDs and proper sequence
            fill_script = """
                try {
                    // Set country code first
                    const countryCode = document.getElementById('codrp');
                    if (!countryCode) return false;
                    countryCode.value = '+92';
                    countryCode.dispatchEvent(new Event('change', {bubbles: true}));
                    countryCode.dispatchEvent(new Event('input', {bubbles: true}));

                    // Fill other fields
                    const fillField = (id, value) => {
                        const el = document.getElementById(id);
                        if (!el) return false;
                        el.value = value;
                        el.dispatchEvent(new Event('change', {bubbles: true}));
                        el.dispatchEvent(new Event('input', {bubbles: true}));
                        return true;
                    };

                    const data = arguments[0];
                    if (!fillField('firstname', data.firstName)) return false;
                    if (!fillField('lastname', data.lastName)) return false;
                    if (!fillField('posttext', data.phone)) return false;
                    if (!fillField('emailaddress1', data.email)) return false;

                    return true;
                } catch (e) {
                    console.error(e);
                    return false;
                }
            """

            # Prepare form data
            form_data = {
                'firstName': phone_number[:20],
                'lastName': phone_number[:20],
                'phone': phone_number[-11:],
                'email': f"{phone_number}@gmail.com"[:40]
            }

            # Fill the form
            if not self.driver.execute_script(fill_script, form_data):
                logging.error("Failed to fill form fields")
                return False

            # Wait for 2 seconds after filling
            time.sleep(2)

            # Find and click the next button using multiple methods
            try:
                # Try using exact input name first
                next_button = self.driver.find_element(By.NAME, "ctl00$ContentContainer$WebFormControl_50052ed5c76aeb11a812002248167989$NextButton")
                if next_button:
                    self.driver.execute_script("arguments[0].click();", next_button)
                    time.sleep(2)
            except:
                try:
                    # Try using ID if name fails
                    next_button = self.driver.find_element(By.ID, "NextButton")
                    if next_button:
                        self.driver.execute_script("arguments[0].click();", next_button)
                        time.sleep(2)
                except:
                    # Try using JavaScript if both fail
                    click_script = """
                        try {
                            const nextBtn = document.querySelector('input[type="button"][value="Next"]');
                            if (nextBtn) {
                                nextBtn.click();
                                return true;
                            }
                            return false;
                        } catch (e) {
                            return false;
                        }
                    """
                    if not self.driver.execute_script(click_script):
                        logging.error("Failed to click next button")
                        return False

            # Wait for popup and click OK
            try:
                # Wait up to 5 seconds for popup
                popup_btn = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "personalbtn"))
                )
                if popup_btn:
                    self.driver.execute_script("arguments[0].click();", popup_btn)
                    time.sleep(2)
                    return True
            except TimeoutException:
                logging.error("Popup button not found")
                return False

        except Exception as e:
            logging.error(f"Error in form filling: {e}")
            return False

        return False

    def process_single_number(self, phone_number: str) -> bool:
        """Process single number with proper cleanup"""
        success = False
        try:
            if self.setup_driver():
                success = self.fill_form(phone_number)
                if success:
                    time.sleep(1)  # Brief delay after success
                    return True
                time.sleep(2)  # Longer delay after failure
        except Exception as e:
            logging.error(f"Error processing number {phone_number}: {e}")
        finally:
            if hasattr(self, 'driver') and self.driver:
                try:
                    if success:
                        # If successful, quit immediately without retries
                        self.driver.quit()
                    else:
                        # Only retry connection if failed
                        try:
                            self.driver.quit()
                        except:
                            pass
                except:
                    pass
        return False

def process_batch(phone_numbers: List[str], max_tabs: int = 20) -> Tuple[List[str], List[str]]:
    """Process multiple numbers simultaneously with parallel form filling"""
    successful, failed = [], []
    total = len(phone_numbers)
    
    print(f"\nProcessing {total} numbers...")
    
    # Process first batch of 20 numbers, then remaining in smaller batches
    remaining_numbers = phone_numbers.copy()
    while remaining_numbers:
        form_filler = None
        try:
            # Initialize browser for this batch
            form_filler = BDOFormFiller()
            if not form_filler.setup_driver() or not form_filler.driver:
                raise Exception("Failed to initialize browser")
            
            # Determine batch size - first batch 20, then smaller batches
            current_batch_size = 20 if len(successful) + len(failed) == 0 else max_tabs
            batch = remaining_numbers[:current_batch_size]
            remaining_numbers = remaining_numbers[current_batch_size:]
            
            # Type guard for driver
            driver = form_filler.driver
            if not isinstance(driver, webdriver.Chrome):
                raise Exception("Driver not properly initialized")

            # Create all tabs at once using JavaScript
            tabs_script = "".join([
                "window.open('');" for _ in range(len(batch) - 1)
            ])
            driver.execute_script(tabs_script)
            
            # Get all handles
            handles = driver.window_handles
            active_tabs = {}  # handle -> phone_number
            
            # Navigate to form URL in all tabs simultaneously
            for i, phone in enumerate(batch):
                active_tabs[handles[i]] = phone
            
            # Open URLs in all tabs simultaneously using JavaScript
            for handle, phone in active_tabs.items():
                driver.switch_to.window(handle)
                driver.get(form_filler.url)
            
            # Wait briefly for all tabs to load
            time.sleep(1)
            
            # Fill all forms simultaneously
            for handle, phone in active_tabs.items():
                driver.switch_to.window(handle)
                # Fill form using fast JavaScript
                driver.execute_script("""
                    try {
                        // Set country code
                        const countryCode = document.getElementById('codrp');
                        if (countryCode) {
                            countryCode.value = '+92';
                            countryCode.dispatchEvent(new Event('change'));
                        }
                        
                        // Fill all fields at once
                        const data = arguments[0];
                        const fields = {
                            'firstname': data.firstName,
                            'lastname': data.lastName,
                            'posttext': data.phone,
                            'emailaddress1': data.email
                        };
                        
                        Object.entries(fields).forEach(([id, value]) => {
                            const el = document.getElementById(id);
                            if (el) {
                                el.value = value;
                                el.dispatchEvent(new Event('change', {bubbles: true}));
                                el.dispatchEvent(new Event('input', {bubbles: true}));
                            }
                        });
                    } catch (e) {
                        console.error(e);
                    }
                """, {
                    'firstName': phone[:20],
                    'lastName': phone[:20],
                    'phone': phone[-11:],
                    'email': f"{phone}@gmail.com"[:40]
                })
            
            # Submit all forms simultaneously
            for handle, phone in active_tabs.items():
                driver.switch_to.window(handle)
                try:
                    # Click next button
                    next_button = driver.execute_script("""
                        const btn = document.querySelector('input[type="button"][value="Next"], button.next, button.submit-btn');
                        if (btn) {
                            btn.click();
                            return true;
                        }
                        return false;
                    """)
                    
                    if next_button:
                        try:
                            # Handle popup quickly
                            popup_btn = WebDriverWait(driver, 3).until(
                                EC.element_to_be_clickable((By.ID, "personalbtn"))
                            )
                            if popup_btn:
                                driver.execute_script("arguments[0].click();", popup_btn)
                                successful.append(phone)
                                print(f"✓ Successfully processed: {phone}")
                                continue
                        except:
                            pass
                    failed.append(phone)
                    print(f"✗ Failed to process: {phone}")
                except Exception as e:
                    print(f"✗ Error completing form for {phone}: {e}")
                    failed.append(phone)
            
            # Show batch progress
            processed = total - len(remaining_numbers)
            success_rate = (len(successful) / processed) * 100 if processed > 0 else 0
            print(f"\nBatch Progress: {processed}/{total} ({processed/total*100:.1f}%)")
            print(f"Success rate: {success_rate:.1f}%")
            
        except Exception as e:
            print(f"Error in batch processing: {e}")
            logging.error(f"Error in batch processing: {e}")
            # Add remaining numbers in batch to failed
            for phone in batch:
                if phone not in successful and phone not in failed:
                    failed.append(phone)
                    
        finally:
            # Clean up browser
            if form_filler and form_filler.driver:
                try:
                    form_filler.driver.quit()
                except:
                    pass
            kill_chrome()
            time.sleep(1)
    
    return successful, failed

def main():
    """Main function with improved stability"""
    try:
        # Initial cleanup
        kill_chrome()
        time.sleep(2)
        
        # Read numbers
        df = pd.read_excel('phone_numbers.xlsx')
        phone_numbers = df['phone_number'].astype(str).tolist()
        total_numbers = len(phone_numbers)
        
        print(f"Found {total_numbers} numbers to process")
        
        # Process numbers
        successful_numbers, failed_numbers = process_batch(phone_numbers)
        
        # Save results
        results_df = pd.DataFrame({
            'phone_number': successful_numbers + failed_numbers,
            'status': ['success'] * len(successful_numbers) + ['failed'] * len(failed_numbers),
            'timestamp': [time.strftime("%Y-%m-%d %H:%M:%S")] * (len(successful_numbers) + len(failed_numbers))
        })
        results_df.to_excel('form_results.xlsx', index=False)
        
        # Show final summary
        print("\n=== Final Results ===")
        print(f"✓ Total Successful: {len(successful_numbers)}")
        print(f"✗ Total Failed: {len(failed_numbers)}")
        print(f"Final Success Rate: {(len(successful_numbers)/total_numbers*100):.1f}%")
        print("Results saved to form_results.xlsx")
        
    except Exception as e:
        print(f"Error in main process: {e}")
        logging.error(f"Error in main process: {e}")
    finally:
        kill_chrome()

if __name__ == "__main__":
    main()
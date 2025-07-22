import pandas as pd
import time
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
import subprocess
import os
import logging
from typing import Dict, List, Optional, Tuple
import json
import re

class UserInteractionReplicator:
    def __init__(self, csv_file_path: str, headless: bool = False):
        """
        Initialize the User Interaction Replicator
        
        Args:
            csv_file_path: Path to the CSV file containing user interaction data
            headless: Whether to run browser in headless mode
        """
        self.csv_file_path = csv_file_path
        self.headless = headless
        self.driver = None
        self.wait = None
        self.current_window_handles = []
        
        # Setup logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('interaction_replication.log'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
        
        # Load and preprocess data
        self.data = self._load_and_preprocess_data()
        
    def _load_and_preprocess_data(self) -> pd.DataFrame:
        """Load CSV data and preprocess it for replication"""
        try:
            df = pd.read_csv(self.csv_file_path)
            
            # Convert time columns to datetime
            time_columns = ['UTCStartTime', 'UTCEndTime', 'StartTime', 'EndTime']
            for col in time_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
            
            # Sort by start time to maintain chronological order
            df = df.sort_values('StartTime').reset_index(drop=True)
            
            self.logger.info(f"Loaded {len(df)} interaction records from CSV")
            return df
            
        except Exception as e:
            self.logger.error(f"Error loading CSV file: {e}")
            raise
    
    def _setup_webdriver(self) -> webdriver.Chrome:
        """Setup Chrome WebDriver with appropriate options"""
        options = Options()
        if self.headless:
            options.add_argument('--headless')
        
        # Add useful options for automation
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        
        driver = webdriver.Chrome(options=options)
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        return driver
    
    def _launch_application(self, process_name: str, exe_name: str) -> bool:
        """
        Launch the specified application
        
        Args:
            process_name: Name of the process
            exe_name: Name of the executable
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Handle web browsers
            if any(browser in process_name.lower() for browser in ['chrome', 'firefox', 'edge', 'safari']):
                if not self.driver:
                    self.driver = self._setup_webdriver()
                    self.wait = WebDriverWait(self.driver, 10)
                return True
            
            # Handle other applications
            else:
                # Try to launch the application
                try:
                    subprocess.Popen([exe_name])
                    time.sleep(2)  # Wait for application to start
                    return True
                except FileNotFoundError:
                    self.logger.warning(f"Could not launch {exe_name}. Application may not be installed.")
                    return False
                    
        except Exception as e:
            self.logger.error(f"Error launching application {process_name}: {e}")
            return False
    
    def _handle_web_interaction(self, row: pd.Series) -> bool:
        """
        Handle web-based interactions
        
        Args:
            row: DataFrame row containing interaction data
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Navigate to URL if provided
            if pd.notna(row['URL']) and row['URL']:
                self.driver.get(row['URL'])
                time.sleep(2)
            
            # Handle different field types and events
            field_name = row.get('FieldName', '')
            field_type = row.get('FieldType', '')
            event = row.get('Event', '')
            sentence = row.get('Sentence', '')
            
            # Find element by various strategies
            element = self._find_web_element(field_name, row)
            
            if element:
                return self._execute_web_action(element, event, sentence, field_type)
            else:
                self.logger.warning(f"Could not find element for field: {field_name}")
                return False
                
        except Exception as e:
            self.logger.error(f"Error in web interaction: {e}")
            return False
    
    def _find_web_element(self, field_name: str, row: pd.Series):
        """
        Find web element using various strategies
        
        Args:
            field_name: Name of the field to find
            row: DataFrame row with additional context
            
        Returns:
            WebElement or None
        """
        if not field_name:
            return None
        
        # Strategy 1: Find by name
        try:
            return self.wait.until(EC.presence_of_element_located((By.NAME, field_name)))
        except:
            pass
        
        # Strategy 2: Find by ID
        try:
            return self.wait.until(EC.presence_of_element_located((By.ID, field_name)))
        except:
            pass
        
        # Strategy 3: Find by CSS selector (if field_name looks like one)
        if '.' in field_name or '#' in field_name:
            try:
                return self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, field_name)))
            except:
                pass
        
        # Strategy 4: Find by partial text content
        try:
            return self.driver.find_element(By.XPATH, f"//*[contains(text(), '{field_name}')]")
        except:
            pass
        
        # Strategy 5: Find by placeholder
        try:
            return self.driver.find_element(By.XPATH, f"//input[@placeholder='{field_name}']")
        except:
            pass
        
        return None
    
    def _execute_web_action(self, element, event: str, sentence: str, field_type: str) -> bool:
        """
        Execute the specified action on the web element
        
        Args:
            element: WebElement to interact with
            event: Type of event (click, type, etc.)
            sentence: Text to input if applicable
            field_type: Type of field (input, button, etc.)
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            event = event.lower() if event else ''
            
            if 'click' in event or event == 'leftmouseclick':
                element.click()
                
            elif 'type' in event or 'keypress' in event or sentence:
                if sentence and pd.notna(sentence):
                    element.clear()
                    element.send_keys(str(sentence))
                    
            elif 'select' in event and field_type.lower() == 'combobox':
                if sentence and pd.notna(sentence):
                    select = Select(element)
                    try:
                        select.select_by_visible_text(str(sentence))
                    except:
                        select.select_by_value(str(sentence))
                        
            elif 'doubleclick' in event:
                ActionChains(self.driver).double_click(element).perform()
                
            elif 'rightclick' in event:
                ActionChains(self.driver).context_click(element).perform()
                
            elif 'hover' in event:
                ActionChains(self.driver).move_to_element(element).perform()
            
            # Add small delay between actions
            time.sleep(0.5)
            return True
            
        except Exception as e:
            self.logger.error(f"Error executing web action {event}: {e}")
            return False
    
    def _handle_desktop_interaction(self, row: pd.Series) -> bool:
        """
        Handle desktop application interactions (placeholder)
        
        Args:
            row: DataFrame row containing interaction data
            
        Returns:
            bool: True if successful, False otherwise
        """
        # This would require additional libraries like pyautogui, win32gui, etc.
        # For now, we'll just log the interaction
        self.logger.info(f"Desktop interaction: {row.get('Event', '')} on {row.get('WindowName', '')}")
        
        # You could implement desktop automation here using:
        # - pyautogui for mouse/keyboard simulation
        # - win32gui for window management
        # - pygetwindow for window focus
        
        return True
    
    def _calculate_delay(self, current_row: pd.Series, next_row: pd.Series = None) -> float:
        """
        Calculate delay between interactions based on original timing
        
        Args:
            current_row: Current interaction row
            next_row: Next interaction row (if exists)
            
        Returns:
            float: Delay in seconds
        """
        if next_row is not None and pd.notna(current_row['EndTime']) and pd.notna(next_row['StartTime']):
            delay = (next_row['StartTime'] - current_row['EndTime']).total_seconds()
            # Cap delay at reasonable maximum (e.g., 5 seconds)
            return min(max(delay, 0.1), 5.0)
        
        # Default delay
        return 1.0
    
    def replicate_interactions(self, start_index: int = 0, end_index: int = None, 
                             speed_multiplier: float = 1.0) -> Dict:
        """
        Main method to replicate user interactions
        
        Args:
            start_index: Index to start replication from
            end_index: Index to end replication at (None for all)
            speed_multiplier: Multiplier for delays (1.0 = original speed, 0.5 = half speed)
            
        Returns:
            Dict: Summary of replication results
        """
        if end_index is None:
            end_index = len(self.data)
        
        results = {
            'total_interactions': 0,
            'successful_interactions': 0,
            'failed_interactions': 0,
            'applications_launched': set(),
            'errors': []
        }
        
        current_process = None
        
        try:
            for i in range(start_index, min(end_index, len(self.data))):
                row = self.data.iloc[i]
                results['total_interactions'] += 1
                
                process_name = row.get('ProcessName', '')
                exe_name = row.get('ExeName', '')
                
                self.logger.info(f"Processing interaction {i+1}/{end_index}: {row.get('Event', '')} on {process_name}")
                
                # Launch application if it's different from current
                if process_name != current_process:
                    if self._launch_application(process_name, exe_name):
                        current_process = process_name
                        results['applications_launched'].add(process_name)
                    else:
                        self.logger.error(f"Failed to launch {process_name}")
                        results['failed_interactions'] += 1
                        continue
                
                # Determine interaction type and handle accordingly
                success = False
                if any(browser in process_name.lower() for browser in ['chrome', 'firefox', 'edge', 'safari']):
                    success = self._handle_web_interaction(row)
                else:
                    success = self._handle_desktop_interaction(row)
                
                if success:
                    results['successful_interactions'] += 1
                else:
                    results['failed_interactions'] += 1
                
                # Calculate and apply delay
                if i < len(self.data) - 1:
                    next_row = self.data.iloc[i + 1]
                    delay = self._calculate_delay(row, next_row) * speed_multiplier
                    time.sleep(delay)
                
        except KeyboardInterrupt:
            self.logger.info("Replication interrupted by user")
        except Exception as e:
            self.logger.error(f"Error during replication: {e}")
            results['errors'].append(str(e))
        finally:
            self._cleanup()
        
        return results
    
    def _cleanup(self):
        """Clean up resources"""
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
    
    def preview_interactions(self, num_rows: int = 10) -> pd.DataFrame:
        """
        Preview the first few interactions from the CSV
        
        Args:
            num_rows: Number of rows to preview
            
        Returns:
            DataFrame: Preview of interactions
        """
        preview_columns = ['ProcessName', 'Event', 'FieldName', 'FieldType', 'Sentence', 'URL', 'StartTime']
        available_columns = [col for col in preview_columns if col in self.data.columns]
        
        return self.data[available_columns].head(num_rows)
    
    def filter_interactions(self, process_name: str = None, application: str = None, 
                          event_type: str = None) -> pd.DataFrame:
        """
        Filter interactions by various criteria
        
        Args:
            process_name: Filter by process name
            application: Filter by application
            event_type: Filter by event type
            
        Returns:
            DataFrame: Filtered interactions
        """
        filtered_data = self.data.copy()
        
        if process_name:
            filtered_data = filtered_data[filtered_data['ProcessName'].str.contains(process_name, case=False, na=False)]
        
        if application:
            filtered_data = filtered_data[filtered_data['Application'].str.contains(application, case=False, na=False)]
        
        if event_type:
            filtered_data = filtered_data[filtered_data['Event'].str.contains(event_type, case=False, na=False)]
        
        return filtered_data

# Example usage
if __name__ == "__main__":
    # Initialize the replicator
    replicator = UserInteractionReplicator("user_interactions.csv")
    
    # Preview interactions
    print("Preview of interactions:")
    print(replicator.preview_interactions())
    
    # Filter for web browser interactions only
    web_interactions = replicator.filter_interactions(process_name="chrome")
    print(f"\nFound {len(web_interactions)} web browser interactions")
    
    # Replicate interactions
    print("\nStarting interaction replication...")
    results = replicator.replicate_interactions(
        start_index=0,
        end_index=10,  # Limit to first 10 interactions for testing
        speed_multiplier=0.5  # Run at half speed
    )
    
    print("\nReplication Results:")
    print(f"Total interactions: {results['total_interactions']}")
    print(f"Successful: {results['successful_interactions']}")
    print(f"Failed: {results['failed_interactions']}")
    print(f"Applications launched: {results['applications_launched']}")
    
    if results['errors']:
        print(f"Errors encountered: {results['errors']}")

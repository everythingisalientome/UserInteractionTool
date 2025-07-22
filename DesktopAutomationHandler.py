import pyautogui
import pygetwindow as gw
import psutil
import time
import cv2
import numpy as np
from PIL import Image
import logging
from typing import Optional, Tuple, Dict, List
import win32gui
import win32con
import win32process

class DesktopAutomationHandler:
    """
    Enhanced desktop automation handler for non-web applications
    Supports Windows desktop applications, Office applications, etc.
    """
    
    def __init__(self, config: Dict = None):
        self.config = config or {}
        self.logger = logging.getLogger(__name__)
        
        # Configure pyautogui
        pyautogui.FAILSAFE = self.config.get('failsafe', True)
        pyautogui.PAUSE = self.config.get('pause', 0.1)
        
        # Window management
        self.active_windows = {}
        self.current_window = None
        
    def find_application_window(self, process_name: str, window_name: str = None) -> Optional[object]:
        """
        Find application window by process name and optionally window name
        
        Args:
            process_name: Name of the process
            window_name: Optional window title to match
            
        Returns:
            Window object or None
        """
        try:
            # Get all windows
            windows = gw.getAllWindows()
            
            for window in windows:
                if window.title and process_name.lower() in window.title.lower():
                    if window_name is None or window_name.lower() in window.title.lower():
                        return window
            
            # Alternative method using win32gui
            def enum_window_callback(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd):
                    window_title = win32gui.GetWindowText(hwnd)
                    if window_title and process_name.lower() in window_title.lower():
                        if window_name is None or window_name.lower() in window_title.lower():
                            windows.append(hwnd)
                return True
            
            windows_list = []
            win32gui.EnumWindows(enum_window_callback, windows_list)
            
            if windows_list:
                hwnd = windows_list[0]
                # Convert hwnd to pygetwindow window object
                title = win32gui.GetWindowText(hwnd)
                for window in gw.getAllWindows():
                    if window.title == title:
                        return window
                        
        except Exception as e:
            self.logger.error(f"Error finding window for {process_name}: {e}")
            
        return None
    
    def activate_window(self, window) -> bool:
        """
        Activate and bring window to foreground
        
        Args:
            window: Window object to activate
            
        Returns:
            bool: True if successful
        """
        try:
            if hasattr(window, 'activate'):
                window.activate()
            else:
                # Use win32gui for activation
                hwnd = self._get_window_handle(window)
                if hwnd:
                    win32gui.SetForegroundWindow(hwnd)
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            
            time.sleep(0.5)  # Wait for window to become active
            self.current_window = window
            return True
            
        except Exception as e:
            self.logger.error(f"Error activating window: {e}")
            return False
    
    def _get_window_handle(self, window) -> Optional[int]:
        """Get window handle from window object"""
        try:
            # Try different methods to get handle
            if hasattr(window, '_hWnd'):
                return window._hWnd
            elif hasattr(window, 'hwnd'):
                return window.hwnd
            else:
                # Find by title
                title = getattr(window, 'title', '')
                if title:
                    return win32gui.FindWindow(None, title)
        except:
            pass
        return None
    
    def handle_field_interaction(self, row) -> bool:
        """
        Handle field-based interactions (text fields, buttons, etc.)
        
        Args:
            row: DataFrame row with interaction data
            
        Returns:
            bool: Success status
        """
        try:
            field_name = row.get('FieldName', '')
            field_type = row.get('FieldType', '').lower()
            event = row.get('Event', '').lower()
            sentence = row.get('Sentence', '')
            window_name = row.get('WindowName', '')
            
            # Find and activate the application window
            process_name = row.get('ProcessName', '')
            window = self.find_application_window(process_name, window_name)
            
            if window and self.activate_window(window):
                return self._execute_desktop_action(field_name, field_type, event, sentence)
            else:
                self.logger.warning(f"Could not find or activate window for {process_name}")
                return False
                
        except Exception as e:
            self.logger.error(f"Error in desktop field interaction: {e}")
            return False
    
    def _execute_desktop_action(self, field_name: str, field_type: str, 
                               event: str, sentence: str) -> bool:
        """
        Execute desktop action based on field type and event
        
        Args:
            field_name: Name of the field
            field_type: Type of field (button, text, etc.)
            event: Type of event
            sentence: Text to input
            
        Returns:
            bool: Success status
        """
        try:
            if 'button' in field_type and 'click' in event:
                return self._click_button(field_name)
                
            elif 'text' in field_type or 'edit' in field_type:
                return self._handle_text_field(field_name, sentence, event)
                
            elif 'combobox' in field_type or 'dropdown' in field_type:
                return self._handle_dropdown(field_name, sentence)
                
            elif 'menu' in field_type:
                return self._handle_menu_click(field_name)
                
            else:
                # Generic click action
                return self._generic_click(field_name)
                
        except Exception as e:
            self.logger.error(f"Error executing desktop action: {e}")
            return False
    
    def _click_button(self, button_name: str) -> bool:
        """Click a button by name or text"""
        try:
            # Method 1: Find button by image template matching
            if self._click_by_image_template(button_name):
                return True
            
            # Method 2: Find button by text using OCR-like approach
            if self._click_by_text_search(button_name):
                return True
                
            # Method 3: Use keyboard shortcut if it's a common button
            shortcut = self._get_button_shortcut(button_name)
            if shortcut:
                pyautogui.hotkey(*shortcut)
                return True
                
            return False
            
        except Exception as e:
            self.logger.error(f"Error clicking button {button_name}: {e}")
            return False
    
    def _handle_text_field(self, field_name: str, text: str, event: str) -> bool:
        """Handle text field interactions"""
        try:
            # Try to find the text field
            if self._find_and_focus_field(field_name):
                if 'clear' in event or 'delete' in event:
                    pyautogui.hotkey('ctrl', 'a')
                    pyautogui.press('delete')
                
                if text and text != '':
                    pyautogui.write(str(text))
                    
                return True
            return False
            
        except Exception as e:
            self.logger.error(f"Error handling text field {field_name}: {e}")
            return False
    
    def _handle_dropdown(self, field_name: str, selection: str) -> bool:
        """Handle dropdown/combobox selections"""
        try:
            if self._find_and_focus_field(field_name):
                # Open dropdown
                pyautogui.press('space')  # or 'enter' or click
                time.sleep(0.5)
                
                # Type selection or use arrow keys
                if selection:
                    pyautogui.write(str(selection))
                    pyautogui.press('enter')
                
                return True
            return False
            
        except Exception as e:
            self.logger.error(f"Error handling dropdown {field_name}: {e}")
            return False
    
    def _find_and_focus_field(self, field_name: str) -> bool:
        """Find and focus on a field"""
        try:
            # Method 1: Use Tab navigation with field name matching
            # This is a simplified approach - you might need more sophisticated logic
            
            # Method 2: Click on field if we can find it by image/text
            if self._click_by_text_search(field_name):
                return True
            
            # Method 3: Use accessibility APIs (requires additional libraries)
            # This would be more reliable but more complex to implement
            
            return False
            
        except Exception as e:
            self.logger.error(f"Error finding field {field_name}: {e}")
            return False
    
    def _click_by_image_template(self, template_name: str) -> bool:
        """Click by finding template image on screen"""
        try:
            # This would require pre-captured template images
            template_path = f"templates/{template_name.lower().replace(' ', '_')}.png"
            
            if os.path.exists(template_path):
                location = pyautogui.locateOnScreen(template_path, confidence=0.8)
                if location:
                    pyautogui.click(location)
                    return True
            
            return False
            
        except Exception as e:
            self.logger.debug(f"Template matching failed for {template_name}: {e}")
            return False
    
    def _click_by_text_search(self, text: str) -> bool:
        """
        Click by finding text on screen using OCR
        This is a simplified version - you'd need actual OCR implementation
        """
        try:
            # Take screenshot
            screenshot = pyautogui.screenshot()
            
            # Convert to OpenCV format
            screenshot_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            # Here you would use OCR to find text
            # For now, this is a placeholder
            # You could use libraries like pytesseract, easyocr, etc.
            
            return False
            
        except Exception as e:
            self.logger.debug(f"Text search failed for {text}: {e}")
            return False
    
    def _get_button_shortcut(self, button_name: str) -> Optional[List[str]]:
        """Get keyboard shortcut for common buttons"""
        shortcuts = {
            'ok': ['enter'],
            'cancel': ['escape'],
            'close': ['alt', 'f4'],
            'save': ['ctrl', 's'],
            'open': ['ctrl', 'o'],
            'copy': ['ctrl', 'c'],
            'paste': ['ctrl', 'v'],
            'cut': ['ctrl', 'x'],
            'undo': ['ctrl', 'z'],
            'redo': ['ctrl', 'y'],
            'find': ['ctrl', 'f'],
            'new': ['ctrl', 'n'],
            'print': ['ctrl', 'p'],
            'refresh': ['f5'],
            'help': ['f1']
        }
        
        button_lower = button_name.lower().strip()
        for key, shortcut in shortcuts.items():
            if key in button_lower or button_lower in key:
                return shortcut
        
        return None
    
    def _generic_click(self, element_name: str) -> bool:
        """Generic click handler for unknown elements"""
        try:
            # Try image template first
            if self._click_by_image_template(element_name):
                return True
            
            # Try text search
            if self._click_by_text_search(element_name):
                return True
            
            # Try common shortcuts
            shortcut = self._get_button_shortcut(element_name)
            if shortcut:
                pyautogui.hotkey(*shortcut)
                return True
            
            return False
            
        except Exception as e:
            self.logger.error(f"Generic click failed for {element_name}: {e}")
            return False
    
    def _handle_menu_click(self, menu_path: str) -> bool:
        """Handle menu navigation (e.g., 'File > Save As')"""
        try:
            menu_items = [item.strip() for item in menu_path.split('>')]
            
            for i, menu_item in enumerate(menu_items):
                if i == 0:
                    # First level menu - use Alt key
                    # Try to find the menu accelerator key
                    alt_key = self._find_menu_accelerator(menu_item)
                    if alt_key:
                        pyautogui.hotkey('alt', alt_key)
                    else:
                        # Click on menu bar
                        self._click_by_text_search(menu_item)
                else:
                    # Submenu items
                    time.sleep(0.5)  # Wait for menu to open
                    self._click_by_text_search(menu_item)
                
                time.sleep(0.3)  # Small delay between menu navigations
            
            return True
            
        except Exception as e:
            self.logger.error(f"Menu navigation failed for {menu_path}: {e}")
            return False
    
    def _find_menu_accelerator(self, menu_name: str) -> Optional[str]:
        """Find menu accelerator key (underlined letter)"""
        # Common menu accelerators
        accelerators = {
            'file': 'f',
            'edit': 'e',
            'view': 'v',
            'insert': 'i',
            'format': 'o',
            'tools': 't',
            'table': 'a',
            'window': 'w',
            'help': 'h'
        }
        
        menu_lower = menu_name.lower().strip()
        return accelerators.get(menu_lower)
    
    def handle_office_interaction(self, row) -> bool:
        """
        Special handler for Microsoft Office applications
        
        Args:
            row: DataFrame row with interaction data
            
        Returns:
            bool: Success status
        """
        try:
            process_name = row.get('ProcessName', '').lower()
            
            if 'excel' in process_name:
                return self._handle_excel_interaction(row)
            elif 'word' in process_name:
                return self._handle_word_interaction(row)
            elif 'powerpoint' in process_name or 'pptx' in process_name:
                return self._handle_powerpoint_interaction(row)
            elif 'outlook' in process_name:
                return self._handle_outlook_interaction(row)
            else:
                return self.handle_field_interaction(row)
                
        except Exception as e:
            self.logger.error(f"Error in Office interaction: {e}")
            return False
    
    def _handle_excel_interaction(self, row) -> bool:
        """Handle Excel-specific interactions"""
        try:
            field_name = row.get('FieldName', '')
            sentence = row.get('Sentence', '')
            event = row.get('Event', '').lower()
            
            # Excel cell reference (e.g., A1, B2)
            if self._is_excel_cell_reference(field_name):
                # Navigate to cell
                pyautogui.hotkey('ctrl', 'g')  # Go to dialog
                time.sleep(0.5)
                pyautogui.write(field_name)
                pyautogui.press('enter')
                
                # Enter data if provided
                if sentence and 'type' in event:
                    pyautogui.write(str(sentence))
                    pyautogui.press('enter')
                
                return True
            
            # Handle Excel ribbon/toolbar clicks
            return self._handle_excel_ribbon_click(field_name)
            
        except Exception as e:
            self.logger.error(f"Excel interaction error: {e}")
            return False
    
    def _is_excel_cell_reference(self, text: str) -> bool:
        """Check if text is an Excel cell reference"""
        import re
        pattern = r'^[A-Z]+[0-9]+$'
        return bool(re.match(pattern, text.upper())) if text else False
    
    def _handle_excel_ribbon_click(self, element_name: str) -> bool:
        """Handle Excel ribbon/toolbar clicks"""
        # Common Excel shortcuts
        excel_shortcuts = {
            'bold': ['ctrl', 'b'],
            'italic': ['ctrl', 'i'],
            'underline': ['ctrl', 'u'],
            'save': ['ctrl', 's'],
            'copy': ['ctrl', 'c'],
            'paste': ['ctrl', 'v'],
            'cut': ['ctrl', 'x'],
            'undo': ['ctrl', 'z'],
            'redo': ['ctrl', 'y']
        }
        
        element_lower = element_name.lower()
        for key, shortcut in excel_shortcuts.items():
            if key in element_lower:
                pyautogui.hotkey(*shortcut)
                return True
        
        # Try generic click
        return self._generic_click(element_name)
    
    def _handle_word_interaction(self, row) -> bool:
        """Handle Word-specific interactions"""
        try:
            field_name = row.get('FieldName', '')
            sentence = row.get('Sentence', '')
            event = row.get('Event', '').lower()
            
            # Handle text typing
            if 'type' in event and sentence:
                pyautogui.write(str(sentence))
                return True
            
            # Handle Word-specific elements
            return self._handle_word_element_click(field_name)
            
        except Exception as e:
            self.logger.error(f"Word interaction error: {e}")
            return False
    
    def _handle_word_element_click(self, element_name: str) -> bool:
        """Handle Word element clicks"""
        word_shortcuts = {
            'bold': ['ctrl', 'b'],
            'italic': ['ctrl', 'i'],
            'underline': ['ctrl', 'u'],
            'save': ['ctrl', 's'],
            'print': ['ctrl', 'p'],
            'find': ['ctrl', 'f'],
            'replace': ['ctrl', 'h']
        }
        
        element_lower = element_name.lower()
        for key, shortcut in word_shortcuts.items():
            if key in element_lower:
                pyautogui.hotkey(*shortcut)
                return True
        
        return self._generic_click(element_name)
    
    def _handle_powerpoint_interaction(self, row) -> bool:
        """Handle PowerPoint-specific interactions"""
        # Similar to Word but with PowerPoint-specific shortcuts
        try:
            field_name = row.get('FieldName', '')
            sentence = row.get('Sentence', '')
            event = row.get('Event', '').lower()
            
            if 'type' in event and sentence:
                pyautogui.write(str(sentence))
                return True
            
            ppt_shortcuts = {
                'new slide': ['ctrl', 'm'],
                'slide show': ['f5'],
                'slide sorter': ['ctrl', 'shift', 's']
            }
            
            field_lower = field_name.lower()
            for key, shortcut in ppt_shortcuts.items():
                if key in field_lower:
                    pyautogui.hotkey(*shortcut)
                    return True
            
            return self._generic_click(field_name)
            
        except Exception as e:
            self.logger.error(f"PowerPoint interaction error: {e}")
            return False
    
    def _handle_outlook_interaction(self, row) -> bool:
        """Handle Outlook-specific interactions"""
        try:
            field_name = row.get('FieldName', '')
            sentence = row.get('Sentence', '')
            event = row.get('Event', '').lower()
            
            outlook_shortcuts = {
                'new mail': ['ctrl', 'n'],
                'send': ['ctrl', 'enter'],
                'reply': ['ctrl', 'r'],
                'reply all': ['ctrl', 'shift', 'r'],
                'forward': ['ctrl', 'f']
            }
            
            field_lower = field_name.lower()
            for key, shortcut in outlook_shortcuts.items():
                if key in field_lower:
                    pyautogui.hotkey(*shortcut)
                    return True
            
            # Handle email field typing
            if sentence and ('to' in field_lower or 'subject' in field_lower or 'body' in field_lower):
                pyautogui.write(str(sentence))
                return True
            
            return self._generic_click(field_name)
            
        except Exception as e:
            self.logger.error(f"Outlook interaction error: {e}")
            return False
    
    def take_screenshot(self, filename: str = None) -> str:
        """Take screenshot for debugging purposes"""
        try:
            if filename is None:
                filename = f"screenshot_{int(time.time())}.png"
            
            screenshot = pyautogui.screenshot()
            screenshot.save(filename)
            return filename
            
        except Exception as e:
            self.logger.error(f"Error taking screenshot: {e}")
            return ""
    
    def get_active_window_info(self) -> Dict:
        """Get information about the currently active window"""
        try:
            active_window = gw.getActiveWindow()
            if active_window:
                return {
                    'title': active_window.title,
                    'left': active_window.left,
                    'top': active_window.top,
                    'width': active_window.width,
                    'height': active_window.height,
                    'isMaximized': active_window.isMaximized,
                    'isMinimized': active_window.isMinimized
                }
        except Exception as e:
            self.logger.error(f"Error getting active window info: {e}")
        
        return {}
    
    def cleanup(self):
        """Clean up resources"""
        try:
            # Reset pyautogui settings
            pyautogui.FAILSAFE = True
            pyautogui.PAUSE = 0.1
        except:
            pass

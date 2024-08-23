from reference_images import *
import PIL.Image
import pyautogui
import pytesseract
import win32gui
import re
import pandas as pd
import win32con
from AppOpener import open as open_app
import time
import PIL
import pydirectinput
import numpy as np
import csv
import pandas as pd
from typing import Optional

class WindowMgr:
    """
    Encapsulates some calls to the win32gui for window management
    """

    def __init__ (self):
        self._handle = None

    def find_window(self, class_name:str, window_name:Optional[str] = None) -> None:
        """
        This method finds a window by its class_name
        """
        self._handle = win32gui.FindWindow(class_name, window_name)

    def get_window_rect(self) -> tuple:
        """
        This method gets the tuple representing the x, y, width, and height of 
        the active window.
        """
        rect = win32gui.GetWindowRect(self._handle)
        x = rect[0]
        y = rect[1]
        width = rect[2] - x
        height = rect[3] - y

        return (x , y, width, height)

    def _window_enum_callback(self, hwnd, wildcard:str) -> None:
        """
        This method passes to win32gui.EnumWindows() to check all the opened windows.
        """
        if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) is not None:
            self._handle = hwnd

    def find_window_wildcard(self, wildcard:str) -> None:
        """
        This method finds a window whose title matches the wildcard regular expression.
        """
        self._handle = None
        win32gui.EnumWindows(self._window_enum_callback, wildcard)

    def set_foreground(self) -> None:
        """
        This method puts the active window in the foreground. 
        """
        win32gui.SetForegroundWindow(self._handle)

    def maximize_window(self) -> None:
        """
        This method maximizes the active window. 
        """
        win32gui.SetWindowPos(
            self._handle, 0, 0, 0, 
            pyautogui.size()[0],
            pyautogui.size()[1],
            win32con.SWP_SHOWWINDOW
        )

class Selector:
    """
    This is a class to simulate that simulates the order selection process. 
    It has the following attributes:
      - An Eye that leverages computer vision to find coordinates on the page
      - A Hand that clicks on a word found by the eye.
      - A WindowMgr to handle window operations like searching, maximizing, and activating. 
    """

    def __init__(self):
        self.eye = Eye()
        self.hand = Hand()
        self.wmgr = WindowMgr()

    def open_order_entry(self) -> None:
        """
        This method opens Citrix Workspace and then Total Information Management System (TIMS)
        using computer vision to navigate the screen. When TIMS opens, it automatically opens
        the Order Entry app. The method then sleeps while the app opens. 
        """

        open_app("Citrix Workspace")
        time.sleep(10)

        self.wmgr.find_window_wildcard("^Citrix Workspace$")
        self.wmgr.set_foreground()
        self.wmgr.maximize_window()

        self.wait_until_seen("Apps")

        categories_coords = pyautogui.locateCenterOnScreen(CATEGORIES_TAB, confidence=0.7)
        pydirectinput.click(categories_coords.x, categories_coords.y)

        time.sleep(0.5)

        uncategorized_coords = pyautogui.locateCenterOnScreen(UNCATEGORIZED_TAB, confidence=0.7)
        pydirectinput.click(uncategorized_coords.x, uncategorized_coords.y)

        time.sleep(0.5)

        tims_logo_coords = pyautogui.locateCenterOnScreen(TIMS_LOGO, confidence=0.7)
        pydirectinput.click(tims_logo_coords.x, tims_logo_coords.y)

        time.sleep(40)

        self.set_selection_mode()

    def set_selection_mode(self) -> None:
        """
        This method sets the mode of the Order Entry app using keyboard shortcuts
        built into the Order Entry app. It enters my initials and initializes 
        selection mode. 
        """

        self.wmgr.find_window_wildcard(".*Order Entry.*")
        
        self.wmgr.set_foreground()
        self.wmgr.maximize_window()
        self.wait_until_seen("Address", self.wmgr.get_window_rect())

        pydirectinput.press('alt', 1, 0.5)
        pydirectinput.press('alt', 1, 0.5)
        print("pressing m")
        pydirectinput.press('m', 1, 0.5)
        print("pressing s")
        pydirectinput.press('s', 1, 0.5)
        pydirectinput.press('shift', 1, 0.5)
        pydirectinput.press('shift', 1, 0.5)

        pydirectinput.press("enter", 3, 0.5)
        pyautogui.write("LB", interval=0.5)
        pydirectinput.press("enter", 2, 0.5)

    def get_order_line_items(self, order_number:str, order_line_items:"pd.DataFrame") -> None:
        """
        This method types in an order number, opens the order and handles any pop up windows or
        program slowdowns that may occur, ultimately presenting the Selector class with 
        a view of all the line items on the order. 
        #
        Parameters:
          - order_number (String): the specific order the program is working on
          - order_line_items (DataFrame): each individual order's worth of items 
        """

        pydirectinput.write(order_number)

        pydirectinput.press("enter", 2, 1.5)

        if self.check_if_already_selected():
            pydirectinput.press("enter", 1, 1.5)

        self.wait_until_seen("Physician")

        pydirectinput.press("enter", 1, 1.5)

        self.wait_until_seen("Whole")

        pydirectinput.press("enter", 2, 1.5)

        self.wait_until_seen("Additional")

        pydirectinput.press("enter", 1, 1.5)

        self.wait_until_seen("Item", rect=self.wmgr.get_window_rect())

        self.eye.get_screen_grab_data(self.wmgr.get_window_rect())

        self.sort_items(order_line_items)

    def sort_items(self, order_line_items:"pd.DataFrame") -> None:
        """
        This method clicks on the "Item" filter button until the item in first position
        on the list of selectable items is in the first position on the list of orders in
        the Order Entry app. This makes it easier to process the orders accordingly. 
        
        Parameters:
          - order_line_items (DataFrame): each individual order's worth of items. 
        """
        
        # Click the "Item" filter button
        self.hand.click_from_screen(self.eye, "Item")

        # Reset the cursor to the top of the list of items (i.e. the row of headers)
        for i in range(len(order_line_items)+5):
            pydirectinput.press('up')
        # Move down one line to the first item. 
        pydirectinput.press('down')

        # Saving shorthand version of first item for comparison
        first_item_shorthand = order_line_items.iloc[0]["Product Code"][-5:]

        # Searching for the first item in the highlighted box, then continuing
        # to click the "Item" filter button until things align. 
        while not self.eye.check_highlighted_item(self.wmgr, self.eye.find_highlight_on_screen(self.wmgr), first_item_shorthand):
            self.hand.click_from_screen(self.eye, "Item")
            for i in range(len(order_line_items)):
                pydirectinput.press('up')
            pydirectinput.press('down')

    def select_next_item(self) -> None:
        """
        This method moves down one item, does the necessary popup window checks, 
        changes the ship quantity appropriately, and then carries through with selecting
        the item.
        """

        pydirectinput.press("enter", 2, 0.5)
        self.check_date_span()
        self.change_ship_qty()
        pydirectinput.press("enter", 1, 0.5)

    def close_order(self) -> None:
        """
        This method searches for the cancel button, presses it, and then
        confirms the cancellation.
        """

        cancel_btn_coords = pyautogui.locateCenterOnScreen(CANCEL_BTN, confidence=0.5)
        pydirectinput.click(cancel_btn_coords.x, cancel_btn_coords.y)

        yes_btn_coords = pyautogui.locateCenterOnScreen(YES_BTN, confidence=0.5)
        pydirectinput.click(yes_btn_coords.x, yes_btn_coords.y)
    
    def check_date_span(self) -> None:
        """
        This method looks for a specific pop-up window and handles it 
        accordingly if it finds it. 
        """

        tries = 0
        self.eye.get_screen_grab_data()
        while not (self.eye.can_see("Date") & self.eye.can_see("Span")):
            if tries >= 3: break
            self.eye.get_screen_grab_data()
            tries += 1

        if tries == 3: 
            print("date span not found")
            pass
        else: 
            pydirectinput.press("tab", 2, 0.5)
            # pydirectinput.write(ship date)
            pydirectinput.press("enter", 2, 0.5)

    def change_ship_qty(self) -> None:
        """     
        This method updates the Ship Quantity field of the order
        using keyboard shortcuts
        """

        pydirectinput.press('tab', 3, 0.5)
        #pydirectinput.write(ship qty)
        pydirectinput.press('enter', 12, 0.5)

    def wait_until_seen(self, target_word:str, rect:tuple = None, max_tries:int = 5) -> bool:
        """
        This method causes the program to wait while it searches for  
        a given target word to found on the screen or not within a given
        number of tries.
        
        Parameters:
          - target_word: the word to search for
          - rect: the rectangle to search in
          - max_tries: the maximum number of times to search for the word
        
        Return: True if target_word is found, False if not 
        """

        print(f'Beginning search for "{target_word}".')
        tries = 1
        self.eye.get_screen_grab_data(rect=rect)
        while not (self.eye.can_see(target_word)):
            print("searching for " + target_word + " " + str(tries))
            if tries >= max_tries: break
            time.sleep(1)
            self.eye.get_screen_grab_data(rect=rect)
            tries += 1

        if tries == max_tries:
            print(f'"{target_word}" was not found on the screen after {tries} tries.')
            return False
        else:
            print(f'"{target_word}" was found after {tries} tries.')
            return True
        
    def check_if_already_selected(self) -> bool:
        return self.wait_until_seen("Warning", rect=self.wmgr.get_window_rect())

class Eye:
    """
    A class to keep track of what pytesseract is seeing. The constructor 
    initializes the tesseract_cmd variable with the appropriate path.
    It has the following attributes:
      - view: an Image that shows the latest screenshot the Eye has analyzed
      - data: a DataFrame that shows the latest OCR data from view. 
    """

    def __init__(self, view:Optional[tuple] = None, data:Optional["pd.DataFrame"] = None):
        pytesseract.pytesseract.tesseract_cmd = 'C:\\Users\\levi.banks\\AppData\\Local\\Programs\\Tesseract-OCR\\tesseract.exe'
        self.view = view
        self.data = data

    def get_screen_grab_data(self, rect:tuple = None) -> None:
        """
        This method takes a screenshot and does some processing on it to then analyze
        it using pytesseract. The Eye updates its data attribute with the DataFrame created
        in this method. 
        #
        Parameters:
          - rect: the coordinates of the area to be analyzed. 
        """

        screengrab = pyautogui.screenshot(imageFilename=r"C:\Users\levi.banks\OneDrive - Providence St. Joseph Health\Python Projects\stealth_pod_import\view.png", region = rect)
        new_size = tuple(2*x for x in screengrab.size)
        screengrab = screengrab.resize(new_size, PIL.Image.Resampling.LANCZOS)
        screengrab.save(r"C:\Users\levi.banks\OneDrive - Providence St. Joseph Health\Python Projects\stealth_pod_import\view.png")
        self.view = screengrab.copy()

        data = pytesseract.image_to_data(screengrab)

        with open(r'C:\Users\levi.banks\OneDrive - Providence St. Joseph Health\Python Projects\stealth_pod_import\data.tsv', 'w', newline='') as tsvfile:
            tsvfile.truncate(0)
            tsvfile.write(data)

        with open(r'C:\Users\levi.banks\OneDrive - Providence St. Joseph Health\Python Projects\stealth_pod_import\data.tsv', 'rb') as tsvfile:
            data = pd.read_csv(r'C:\Users\levi.banks\OneDrive - Providence St. Joseph Health\Python Projects\stealth_pod_import\data.tsv', delimiter='\t', na_values=-1, encoding='unicode_escape', on_bad_lines="skip", quoting=csv.QUOTE_NONE)

        self.data = data

    def can_see(self, target_word:str) -> bool:
        """
        This method verifies whether a given word is found in the data attribute
        of the Eye object.
        
        Parameters:
          - target_word: the word to search for
        """

        self.get_screen_grab_data()
        word_is_visible = not self.data[self.data['text'] == target_word].empty
        return word_is_visible
    
    def find_highlight_on_screen(self, 
                                 wmgr:"WindowMgr", 
                                 color:tuple = (51, 153, 255), 
                                 color_tolerance:int = 5) -> tuple | None:
        """
        This function looks for the area where a given color is present on the screen, returning
        the coordinates of a rectangle that surrounds that color. 
        #
        Parameters:
          - wmgr: the WindowMgr object that will help get the rectangle of the window of interest
          - color: the color in question, with a default value for this project that matches the 
                   highlight color of the Order Entry app 
          - color_tolerance: the amount of wiggle room allowed when detecting the color
        """
        
        window_rect = wmgr.get_window_rect()

        self.get_screen_grab_data(window_rect)
        
        window_coords = (window_rect[0], window_rect[1])

        # The coordinates to track where the color has been found.
        left = self.view.width
        top = self.view.height
        right = 0
        bottom = 0
        
        # Going through the Eye's view to look for the color pixel by pixel.
        color_found = False
        for x in range(self.view.width):
            for y in range(self.view.height):
                r, g, b = self.view.getpixel((x,y))
                if abs(r - color[0]) <= color_tolerance and \
                   abs(g - color[1]) <= color_tolerance and \
                   abs(b - color[2]) <= color_tolerance:
                    # When the color is found, update the coordinates accordingly.
                    color_found = True
                    left = min(left, x)
                    top = min(top, y)
                    right = max(right, x)
                    bottom = max(bottom, y)
        
        # Creating a bounding box tuple that accounts for where the window is relative
        # to the screen. 
        color_bbox = (
            left + window_coords[0], 
            top + window_coords[1], 
            right + window_coords[0], 
            bottom + window_coords[1]
        )

        # Move to the center of where the color is found, mainly for debugging purposes
        if color_found:
            pyautogui.moveTo(
                np.mean([color_bbox[0], color_bbox[2]]),
                np.mean([color_bbox[1], color_bbox[3]])
            )
            return color_bbox
        else:
            print("The color was not found.")
            return None

    def check_highlighted_item(self, wmgr:"WindowMgr", color_bbox:tuple, item:str) -> bool:
        """
        
        This method uses a color bounding box from the find_highlight_on_screen 
        method to check that highlighted portion of the screen for a particular
        order line item. It returns True or False depending on if the item is in
        the highlighted area.
        
        Parameters:
          - wmgr: the WindowMgr object that will help manage window operations
          - color_bbox: the tuple representing coordinates of where a particular
                        color is on screen
          - item: the order line item in question to be checked against
        
        """

        window_rect = wmgr.get_window_rect()

        # Getting the text data from the specific area of the color_bbox.
        self.get_screen_grab_data(
            rect=(
                color_bbox[0],
                color_bbox[1],
                color_bbox[2] - color_bbox[0],
                color_bbox[3] - color_bbox[1]
            )
        )

        # Getting the text data for the particular item that we are looking for.
        item_data = self.data[self.data['text'].str.contains(item).fillna(False)]

        # Starting with item_coords as None
        item_coords = None

        # Trying to populate item_coords based on the relative position of the 
        # item within the full monitor coordinate system. If it fails, the word
        # wasn't found, so item_coords remains None, and the method returns False.
        try:
            item_coords = (
                item_data['left'].values[0] + window_rect[0],
                item_data['top'].values[0] + window_rect[1],
                item_data['left'].values[0] + item_data['width'].values[0] + window_rect[0],
                item_data['left'].values[0] + item_data['height'].values[0] + window_rect[1]
            )
        except:
            pass

        if item_coords:
            return True
        else: 
            return False

class Hand():
    """
    A class that simulates some clicking actions done by the Selector class.
    """
    def __init__(self):
        pass

    def click_active_window(self, wmgr:"WindowMgr") -> None:
        """
        This method clicks the active window in the top left corner. It's used
        to make the window active in certain instances.

        Parameters:
            - wmgr: WindowMgr object to handle window operations
        """
        window_rect = wmgr.get_window_rect()
        pyautogui.click(window_rect[0] + 50, window_rect[1] + 55)

    def click_from_screen(self, eye: "Eye", target_word: str, xadj:int = 0, yadj:int = 0) -> None:
        """
        This method clicks a particular word on the screen, taking into account any
        adjustments provided by the user. 

        Parameters:
            - eye: Eye object to check visibility of target_word
            - target_word: the word to be found
            - xadj: the amount to adjust from the "origin" coordinates of target_word, x-axis
            - yadj: the amount to adjust from the "origin" coordinates of target_word, y-axis
        """
        tries = 0
        while not eye.can_see(target_word):
            if tries >= 3: break
            eye.get_screen_grab_data()
            tries += 1

        if tries == 3:
            print("I couldn't find that word.")
            return None
        else:
            pyautogui.click(
               (eye.data[eye.data['text'] == target_word]['left'].values[0]) / 2 + xadj,
               (eye.data[eye.data['text'] == target_word]['top'].values[0]) / 2 + yadj
            )
from datetime import datetime, date, timedelta
import os
import fitz
import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium_stealth import stealth
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException, NoSuchWindowException, ElementClickInterceptedException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time
from tkinter import filedialog
import tkinter as tk
import win32com.client
import constants
from selenium.webdriver.common.action_chains import ActionChains
from pathlib import Path
from zipfile import ZipFile
import pyautogui
import pytesseract
from selection import Selector
from typing import Callable, Optional

class Window:
    """
    A class that defines generic methods for constructing windows using 
    the Tkinter package. 

    Attributes:
        - root: the Tk root of each Window object
        - entries: a dictionary to keep track of the entry objects/variables 
                   that are added to the window
        - check_vars: a dictionary like entries, but it is for checkboxes and their
                      variables
        - current_row: the row that the last element was added to. 
    """
    def __init__(self, root:Optional["tk.Tk"] = None, geometry:str = "500x500", title:str = ""):
        if root is None:
            self.root = tk.Tk()
        else:
            self.root = root
        self.root.geometry(geometry)
        self.root.attributes('-topmost', )
        self.root.title(title)
        self.entries = {}
        self.check_vars = {}
        self.current_row = 0

    def add_label(self, text:str, font:tuple = ('calibre', 12, 'bold'), 
                  justify:str = 'center', row:int | None = None, column:int = 0, 
                  padx:tuple = (0,0), pady:tuple = (0,0)) -> "tk.Label":
        """
        This method adds a label object to the window's grid, taking in parameters
        to pass to the tk.Label constructor and the grid method.

        Parameters:
            - text: what the label will say
            - font: how the font should look (<font_family>, <font_size>, <emphasis>)
            - justify: how the text should be aligned
            - row: the grid row to place the label in
            - column: the grid column to place the label in
            - padx: the amount of horizontal padding to include on either side of the label. 
            - pady: the amount of vertical padding to include on either side of the label.
        """
        if row is None:
            row = self.current_row
            self.current_row += 1

        label = tk.Label(self.root, text = text, font = font, justify= justify)
        label.grid(row=row, column=column, padx=padx, pady=pady)
        return label

    def add_entry(self, name:str, text:str, font:tuple = ('calibre', 10, 'normal'), 
                  justify:str = 'center', show:str | None = None, 
                  row:int | None = None, column:int = 0, 
                  padx:tuple = (0,0), pady:tuple = (0,0), ipadx:int = 0) -> "tk.Entry":
        """
        This method adds a label and an entry object to the window's grid, taking in parameters
        to pass to the Tkinter constructors and the grid method.

        Parameters:
            - name: the key for the Entry object in the entries dictionary. 
            - text: what the label will say
            - font: how the font should look (<font_family>, <font_size>, <emphasis>)
            - justify: how the text should be aligned
            - show: the character that the Entry object should show when text is typed into it
                    (None means it will just show normal text)
            - row: the grid row to place the label and entry in
            - column: the grid column to place the label and entry in
            - padx: the amount of horizontal padding to include on either side of the label. 
            - pady: the amount of vertical padding to include on either side of the label.
            - ipadx: the amount of internal horizontal padding to include on either side of
                     the entered text. 
        """
        if row is None:
            row = self.current_row
            self.current_row += 1

        var = tk.StringVar()
        label = tk.Label(self.root, text=text, font=font, justify=justify)
        entry = tk.Entry(self.root, textvariable=var, font=font, show=show, justify=justify)
        label.grid(row=row, column=column, padx=padx, pady=pady)
        entry.grid(row=row+1, column=column, padx=padx, pady=pady, ipadx=ipadx)
        self.entries[name] = (entry, var)
        return entry

    def add_checkbutton(self, name:str, text:str, font:tuple = ('calibre', 12, 'bold'), 
                        justify:str = 'center', default:bool = True, 
                        row:int | None = None, column:int = 0, 
                        padx:tuple = (0,0), pady:tuple = (0,0)) -> "tk.Checkbutton":
        """
        This method adds a checkbutton object to the window's grid, taking in parameters
        to pass to the Tkinter constructor and the grid method.

        Parameters:
            - name: the key for the Checkbutton object in the check_vars dictionary 
            - text: the text to accompany the Checkbutton
            - font: how the font should look (<font_family>, <font_size>, <emphasis>)
            - justify: how the text should be aligned
            - default: the starting state of the Checkbutton (default is True/checked)
            - row: the grid row to place the Checkbutton in
            - column: the grid column to place the Checkbutton in
            - padx: the amount of horizontal padding to include on either side of the Checkbutton. 
            - pady: the amount of vertical padding to include on either side of the Checkbutton. 
        """
        if row is None:
            row = self.current_row
            self.current_row += 1

        check_var = tk.BooleanVar(value=default)
        check = tk.Checkbutton(
            self.root,
            text=text,
            font=font,
            justify=justify,
            variable=check_var
        )
        check.grid(row=row, column=column, padx=padx, pady=pady)
        self.check_vars[name] = (check, check_var)
        return check

    def add_button(self, text:str, command:Callable, 
                   row:int | None = None, column:int = 0, 
                   padx:tuple = (0,0), pady:tuple = (0,0)):
        """
        This method adds a button object to the window's grid, taking in parameters
        to pass to the Tkinter constructors and the grid method.

        Parameters:
            - text: what the button will say
            - command: the function that the button should carry out when pressed
            - row: the grid row to place the button in
            - column: the grid column to place the button in
            - padx: the amount of horizontal padding to include on either side of the button. 
            - pady: the amount of vertical padding to include on either side of the button. 
        """
        if row is None:
            row = self.current_row
            self.current_row += 1

        button = tk.Button(self.root, text=text, command=command)
        button.grid(row=row, column=column, padx=padx, pady=pady)
        return button

    def close_window(self) -> None:
        """
        A simple method to close the window, used for the submit buttons mainly
        """
        self.root.destroy()

    def display(self) -> None:
        """
        A simple method to display a given window after its elements have been set up
        """
        self.root.mainloop()

    def get_values(self) -> dict:
        """
        This method returns a dictionary of dictionaries that contain the names and values
        for each Entry object and Checkbutton object. 
        """
        return {
            'entry_vals': {k: v[1].get() for k, v in self.entries.items()},
            'checkbtn_vals': {k: v[1].get() for k, v in self.check_vars.items()}
        }
    
    def get_downloads_folder(self) -> str:
        """
        This method creates a special file dialog window to have the user select and
        submit their downloads folder. This is returned as a string version of the path
        to the folder. 
        """
        self.root.withdraw()
        downloads_folder = filedialog.askdirectory(title='Please select your Downloads folder.')
        self.root.destroy()
        os.chdir(downloads_folder)
        return downloads_folder

def display_intro() -> None:
    """
    This function leverages the Window class created above to display some
    introductory text that explains what the program will do for the user. 
    """

    intro_text = """
    This program will automatically download the Open Orders Details Report 
    and then use the credentials you enter to download all relevant Proof of 
    Delivery documents (PODs) from the Cardinal website for PAP resupply orders. 
    It then reads and formats all of this data down into a digestible spreadsheet 
    of all delivered orders from the specified Cardinal account. 

    The program will ask you to specify your downloads folder and then will ask
    you to enter your Cardinal 4477CPAP account information. 

    Please click Continue to start the program. 
    """

    intro_window = Window(geometry="675x300", title="Automated 4477CPAP Order Selection")
    intro_window.add_label(intro_text)
    intro_window.add_button("Continue", intro_window.close_window)
    intro_window.display()

def display_settings() -> dict:
    """
    This function displays the list of steps in the process and allows the user to choose
    which of the steps they would like to have executed, allowing the user to skip tedious
    repeats of the whole process if only one part of the process fails. 
    """

    settings_text = """
    This program goes through a series of steps to download all 
    necessary documents for this procedure. Sometimes there are issues
    that cause the code to stop working in the middle of the procedure.

    Please check the boxes for the steps you need the program to 
    go through for your case. Please only select later steps by themselves
    if the code has completed the previous steps successfully.
    """

    settings_window = Window(geometry="600x400", title="Settings")

    settings_window.add_label(settings_text)
    settings_window.add_checkbutton("open orders report", "Download Open Orders Details Report")
    settings_window.add_checkbutton("cardinal PODs", "Download Cardinal POD documents")
    settings_window.add_checkbutton("find selectables", "Find all selectable items using previous documents")
    settings_window.add_checkbutton("select items", "Select items from 'Selectable Items.xlsx'")

    settings_window.add_button("Submit", settings_window.close_window)

    settings_window.display()

    return settings_window.get_values()['checkbtn_vals']

def get_downloads_folder() -> str:
    """
    This function creates a basic Window and uses it get the path to the user's download
    folder. 
    """
    file_dialog = Window(title='Please select your Downloads folder.')
    return file_dialog.get_downloads_folder()

def get_credentials(first_try:bool) -> list:
    """
    This function creates a log-in window for the user to enter/re-enter their log-in
    info for the site that this code is trying to reach. It then returns the log-in
    information as a list for later use. 

    Parameters:
        - first_try: alters window text slightly if set to false
    """
    credentials_window = Window(geometry="575x300", title="Login")

    row = 0
    if first_try:
        credentials_window.add_label(
            text='Please enter your Cardinal 4477CPAP Account information:',
            row=row, column=1, padx=(50,0), pady=(20,10)
        )
    else:
        credentials_window.add_label(
            text='There was an error, please try again:',
            row=row, column=1, padx=(100,0), pady=(20,10)
        )

    row += 2
    credentials_window.add_entry(
        "username_entry", "Username", 
        row=row, column=1, padx=(50,0), ipadx=50
    )
    row += 2
    credentials_window.add_entry(
        "password_entry", "Password", show='*', 
        row=row, column=1, padx=(50,0), ipadx=50
    )

    def toggle_password() -> None:
        """
        This sub-function toggles whether the password Entry object is showing asterisks
        or actual letters. 
        """
        password_entry = credentials_window.entries['password_entry'][0]
        if password_entry.cget('show') == '':
            password_entry.config(show='*')
            show_pw_btn.config(text='Show password')
        else:
            password_entry.config(show='')
            show_pw_btn.config(text='Hide password')

    row += 2
    show_pw_btn = credentials_window.add_button(
        "Show password", command=toggle_password, 
        row=row, column = 1, padx=(50,0), pady=(0, 5)
    )

    row += 2
    credentials_window.add_entry(
        "account_number_entry", "Account Number", 
        row=row, column=1, padx=(50,0), ipadx=50
    )

    row += 2
    credentials_window.add_button(
        "Submit", command=credentials_window.close_window, 
        row=row, column=1, padx=(50,0)
    )

    credentials_window.display()

    values = credentials_window.get_values()
    return [
        values['entry_vals']['username_entry'],
        values['entry_vals']['password_entry'],
        values['entry_vals']['account_number_entry']
    ]

def engage_stealth_mode() -> "WebDriver":
    """
    This function creates a WebDriver object that is configured such that it is
    "stealth mode", meaning that it can more easily bypass web automation roadblocks. 
    """

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    driver = webdriver.Chrome(options=options)

    stealth(driver,
            languages=['en-US', 'en'],
            vendor='Google Inc.',
            platform='Win32',
            webgl_vendor='Intel Inc.',
            renderer="Intel Iris OpenGL Engine",
            fix_hairline=True)
    
    return driver

def wait_for_element(driver:"WebDriver", identifier:str, by:str, wait_time:int = 600) -> Optional["WebElement"]:
    """
    This function causes the WebDriver to wait until a particular element is visible on the 
    page, found using either its XPATH or its ID. 

    Parameters:
        - driver: the WebDriver controlling the browser
        - identifier: the string identifying the element in question
        - by: search by 'xpath' or 'id'
        - wait_time: the time in seconds that the driver will wait before throwing an error.
    """
    
    wait = WebDriverWait(driver, wait_time)
    if by == 'xpath':
        return wait.until(EC.visibility_of_element_located((By.XPATH, identifier)))
    if by == 'id':
        return wait.until(EC.visibility_of_element_located((By.ID, identifier)))

def select_branches(driver:"WebDriver", dropdown_xpath:str, branches:list) -> None:
    """
    This function goes through a list of checkbuttons and selects only the relevant ones
    based on the content of branches. 

    Parameters:
        - driver: the driver controlling the web browser
        - dropdown_xpath: the xpath string for the "branches" dropdown
        - branches: a list of branches that are desired
    """
    dropdown_menu = driver.find_element(By.XPATH, dropdown_xpath)
    for element in dropdown_menu.find_elements(By.XPATH, './/*'):
        label = element.text
        already_clicked = False
        for sub_element in element.find_elements(By.XPATH, './/*'):
            if label in branches and not already_clicked:
                sub_element.click()
                already_clicked = sub_element.is_selected()

def wait_for_element_text_change(driver:"WebDriver", element_xpath:str, 
                                 old_text:str, timeout:int = 300) -> bool:
    """
    This function waits until the old_text is no longer present in the element 
    at the xpath element_xpath. If it changes within the timeout, return True, 
    else return false. 
    """

    try:
        WebDriverWait(driver, timeout).until_not(
            EC.text_to_be_present_in_element((By.XPATH, element_xpath), old_text)
        )
        return True
    except TimeoutException:
        return False  

def get_open_orders_report(driver:"WebDriver") -> None:
    """
    This function accesses the report server website to download the Open Orders Details Report.
    It chooses the appropriate branches and insurance values, waits for the report to load, 
    then downloads the report. 

    Parameters:
        - driver: the WebDriver controlling the browser
    """

    driver.get(constants.OODR_VIEWER)

    branch_dd_btn = wait_for_element(driver, '//*[@id="ReportViewerControl_ctl04_ctl03_ctl01"]', 'xpath')
    branch_dd_btn.click()
    wait_for_element(driver, '//*[@id="ReportViewerControl_ctl04_ctl03_divDropDown_ctl00"]', 'xpath')
    select_branches(driver, '//*[@id="ReportViewerControl_ctl04_ctl03_divDropDown"]', constants.DESIRED_BRANCHES)

    ins_dd_btn = wait_for_element(driver, '//*[@id="ReportViewerControl_ctl04_ctl05_ctl01"]', 'xpath')
    ins_dd_btn.click()
    select_all_btn = wait_for_element(driver, '//*[@id="ReportViewerControl_ctl04_ctl05_divDropDown_ctl00"]', 'xpath')
    select_all_btn.click()

    view_report_btn = driver.find_element(By.ID, "ReportViewerControl_ctl04_ctl00")
    view_report_btn.click()

    wait_for_element(driver, '//*[@id="ReportViewerControl_ctl05_ctl00_TotalPages"]', 'xpath')
    wait_for_element_text_change(driver, '//*[@id="ReportViewerControl_ctl05_ctl00_TotalPages"]', '0', 600)

    download_dd_btn = wait_for_element(driver, "ReportViewerControl_ctl05_ctl04_ctl00_ButtonImg", 'id')
    download_dd_btn.click()

    download_dd_excel = wait_for_element(driver, '//*[@id="ReportViewerControl_ctl05_ctl04_ctl00_Menu"]/div[2]/a', 'xpath')
    download_dd_excel.click()

def keep_download_check() -> None:
    """
    This function creates a Selector object to look for the "insecure download" pop-up 
    that Chrome shows after downloading the report. It then clicks the "Keep" button to
    save the download appropriately.
    """
    selector = Selector()

    while not selector.eye.can_see("Keep"):
        time.sleep(1)

    selector.hand.click_from_screen(selector.eye, "Keep")
    
def get_sorted_files(folder:str) -> list[str]:
    """
    This function takes a path to a folder and returns a list of file paths in that
    folder sorted based on their "created at" time, newest to oldest. 

    Parameters:
        - folder: the string identifying the path to the folder in question
    """

    sorted_files = sorted(os.listdir(folder), key=os.path.getctime, reverse=True)
    return sorted_files
                
def get_open_orders_from_downloads(sorted_files:list[str]) -> "pd.DataFrame":
    """
    This function takes a list of sorted files and finds the most recent one with 
    "Open Orders" in the title. It then opens the file in Excel to do some manipulation
    before saving the raw report back into a DataFrame.

    Parameters: 
        - sorted_files: the list of files in question, sorted from newest to oldest. 
    """

    open_orders_df = None
    open_orders_path = [f for f in sorted_files if "Open Orders" in f][0]

    if os.path.exists(open_orders_path):
        xl = win32com.client.Dispatch('Excel.Application')
        xl.Visible = True
        workbook = xl.Workbooks.Open(os.path.abspath(open_orders_path))
        with_header_footer = workbook.ActiveSheet
        if len(workbook.Worksheets) == 1:
            with_header_footer.Range('A5').Select()
            with_header_footer.Range(with_header_footer.Cells(5, 1), \
                                     with_header_footer.Cells(with_header_footer.UsedRange.Rows.Count-1, \
                                                              with_header_footer.UsedRange.Columns.Count)).Copy()
            raw_report = workbook.Worksheets.Add()
            raw_report.Name = 'Raw Report'
            raw_report.Range('A1').Select()
            raw_report.Paste()
        else:
            raw_report = workbook.Worksheets['Raw Report']
        workbook.Save()
        workbook.Close(False)
        open_orders_df = pd.read_excel(open_orders_path, parse_dates=True, sheet_name='Raw Report', engine='openpyxl')

        xl.Quit()
        del xl

    return open_orders_df  

def get_code_from_inbox() -> str:
    """
    This function creates a connection to Microsoft Outlook and waits
    for a two-factor authentication code email to enter the inbox. It then 
    uses a regular expression to search the email for the code and then returns
    the code. 
    """

    # Establishing a connection with Outlook and finding the inbox
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace("MAPI")
    inbox = None
    for folder in outlook.Folders:
        for subfolder in folder.Folders:
            subfolder_name = subfolder.Name
            if subfolder_name == 'Inbox':
                inbox = subfolder

    # Waiting until we receive the Two-factor authentication email
    current_num_messages = len(inbox.Items)
    while len(inbox.Items) == current_num_messages:
        time.sleep(0.5) 

    # Searching through the messages to find the latest one with 
    # a verification code from Cardinal
    message_with_code = None
    for i, message in enumerate(inbox.Items):
        subject = message.Subject
        if ('One-time verification code' in subject):
            message_with_code = message

    # Using a regular expression to search for a string of six numbers
    code = re.search(r'\d{6,6}', message_with_code.Body).group(0)
    message_with_code.Delete()
    return code 

def get_PODs(driver:"WebDriver", df:"pd.DataFrame") -> None:
    """
    This function goes to the report download page of the vendor's website and 
    searches for/downloads all PODs from all relevant date ranges (calculated by get_date_ranges).
    It can handle some error pages, but it requires some manual input occasionally if the error page
    shows up well after the tab has already been open. 

    Parameters:
        - driver: the WebDriver controlling the web browser
        - df: the DataFrame being used to calculate the relevant date ranges
    """

    minute_wait = WebDriverWait(driver, 300)
    seconds_wait = WebDriverWait(driver, 5)
    actions = ActionChains(driver)

    date_ranges = get_date_ranges(df)

    for date_range in date_ranges:
        min_date = date_range[0]
        max_date = date_range[1]

        driver.get(constants.POD_SEARCH)

        set_date_range(driver, min_date, max_date)
        click_POD_search(driver)

        minute_wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="divContent"]/div/div[1]/font')))
        page_list = get_page_list(driver)

        for i in page_list:
            set_page(driver, i)
            click_all_PODs(driver)
            old_handles = set(driver.window_handles)
            download_selected_PODs(driver)
            new_handles = set(driver.window_handles)-old_handles
            minute_wait.until(lambda d: len(new_handles) > 0)
            new_handle = new_handles.pop()

            while error_page(driver, new_handle):
                new_handle = restart_download(driver, minute_wait, actions, new_handle)
            
            while len(driver.window_handles) > 3:
                time.sleep(1)

        driver.switch_to.window(driver.window_handles[0])
        try:
            close_popup = seconds_wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Close"]')))
            actions.move_to_element(close_popup).perform()
            close_popup.click()
        except TimeoutException:
            pass

        minute_wait.until(EC.number_of_windows_to_be(1))

def get_date_ranges(df:"pd.DataFrame") -> list[tuple["date"]]:
    """
    This function iterates through all the dates in the Create Date column of the df
    parameter and finds all date ranges that are separated by 3 days or more. For instance, 
    if the column had dates ranging from 12/12/24 to 12/24/24 with no gaps, the return would
    be the list [(12/12/24, 12/24/24)], but if there were any gaps in the dates of 3 days 
    or more, the function would return a list of multiple tuples, one for each separate range.

    Parameters:
        - df: the DataFrame being analyzed
    """
    dates = [date for date in sorted(set(df['Create Date'].dt.date.tolist()))]
    date_ranges = []
    current_range = None

    for i in range(len(dates)):
        if current_range is None:
            current_range = [dates[i], dates[i]]
        else:
            if dates[i] - current_range[-1] <= timedelta(days=3):
                current_range[-1] = dates[i]
            else:
                date_ranges.append(tuple(current_range))
                current_range = [dates[i], dates[i]]

    if current_range is not None:
        date_ranges.append(tuple(current_range))

    return date_ranges

def set_date_range(driver:"WebDriver", min_date:date, max_date:date) -> None:
    """
    This function gets the date values for a given date range and inputs them 
    into the appropriate fields on the vendor's website.

    Parameters:
        - driver: the WebDriver controlling the web browser
        - min_date: the earliest date to look for reports from
        - max_date: the latest date to look for reports from
    """

    start_date = datetime.strftime(min_date, '%m/%d/%Y')
    end_date = datetime.strftime(max_date, '%m/%d/%Y')
    driver.execute_script('document.getElementById("txtStartDate").value = "' + start_date + '"')
    driver.execute_script('document.getElementById("txtEndDate").value = "' + end_date + '"')

def click_POD_search(driver:"WebDriver") -> None:
    """
    This function clicks the search button after setting the date range.

    Parameters:
        - driver: the WebDriver controlling the web browser
    """

    search_button = driver.find_element(By.XPATH, '//*[@id="searchFormTable"]/tbody/tr[9]/td/input')
    search_button.click()

def get_page_list(driver:"WebDriver") -> list[int]:
    """
    This function looks at the paginate buttons at the top of the POD
    search results and returns a list of the page numbers. 

    Parameters:
        - driver: the WebDriver controlling the web browser
    """

    # Finding the first and last page buttons
    first_page_btn = driver.find_element(By.XPATH, '//*[@id="podResults_first"]')
    last_page_btn = driver.find_element(By.XPATH, '//*[@id="podResults_last"]')

    # Finding the HTML element that holds all the page buttons
    page_btns_span = driver.find_element(By.XPATH, '//*[@id="podResults_paginate"]/span')
    page_btns = page_btns_span.find_elements(By.XPATH, '//*[@id="podResults_paginate"]/span/a')

    # Finding the button and innerHTML value for page one
    page_one_btn = page_btns[0]
    page_one = int(page_one_btn.get_attribute('innerHTML'))

    # Clicking the last page button and updating page_btns to reflect the new page
    # numbers shown
    last_page_btn.click()
    page_btns = page_btns_span.find_elements(By.XPATH, '//*[@id="podResults_paginate"]/span/a')
    
    # Finding the button and innerHTML value for page n
    page_n_btn = page_btns[-1]
    page_n = int(page_n_btn.get_attribute('innerHTML'))

    # Navigating back to the first page
    first_page_btn.click()

    # Creating a list out of the page_one and page_n values and returning it.
    page_list = list(range(page_one,
                           page_n + 1))
    return page_list

def set_page(driver:"WebDriver", page_number:int) -> None:
    """
    This function iteratively clicks through all the paginate buttons until it reaches
    the target page, identified by page_number.

    Parameters:
        - driver: the WebDriver controlling the web browser
        - page_number: the target page number to be reached
    """
    wait = WebDriverWait(driver, 2)
    actions = ActionChains(driver)
    driver.switch_to.window(driver.window_handles[0])

    try:
        close_popup = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Close"]')))
        actions.move_to_element(close_popup).perform()
        close_popup.click()
    except TimeoutException:
        pass

    select_all_checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="chkAll"]')))
    select_all_checkbox.click()
    select_all_checkbox.click()

    while True:
        page_btns_span = driver.find_element(By.XPATH, '//*[@id="podResults_paginate"]/span')
        page_btns = page_btns_span.find_elements(By.XPATH, '//*[@id="podResults_paginate"]/span/a')

        target_page_btn = None
        for btn in page_btns:
            try:
                if int(btn.get_attribute('innerHTML')) == page_number:
                    target_page_btn = btn
                    break
            except ValueError:
                continue

        if target_page_btn:
            target_page_btn.click()
            return
        else:
            next_page_btn = driver.find_element(By.XPATH, '//*[@id="podResults_next"]')
            next_page_btn.click()
            wait.until(EC.staleness_of(page_btns_span))

def click_all_PODs(driver:"WebDriver") -> None:
    """
    This function clicks all the checkboxes on a particular page of documents. 

    Parameters:
        - driver: the WebDriver controlling the web browser
    """

    # Find all the checkboxes on the page
    POD_checkboxes = driver.find_elements(By.XPATH, '//input[@type="checkbox"]')        
    for checkbox in POD_checkboxes:

        # Go down the list and check off each POD 
        # (avoiding the Select All checkboxes)
        if 'chkAll' not in checkbox.get_attribute('name'):
            checkbox.click()

def download_selected_PODs(driver:"WebDriver") -> None:
    """
    This function downloads the PODs selected by the click_all_PODs function. 

    Parameters:
        - driver: the WebDriver controlling the web browser
    """
     
    wait = WebDriverWait(driver, 60)
    actions = ActionChains(driver)

    # Locating the download button, moving to it, and clicking it.
    download_btn = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="save"]')))
    driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.CONTROL + Keys.HOME)
    time.sleep(1)
    download_btn.click()

    # Clicking the confirm download button that appears in a popup window
    confirm_download_btn = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="yes"]')))
    confirm_download_btn.click()

def error_page(driver:"WebDriver", page_handle:str) -> bool:
    """
    This function checks a particular page for error text using computer vision (pytesseract). 
    It returns a boolean based on whether it finds an error message or not.

    Parameters:
        - driver: the WebDriver controlling the web browser
        - page_handle: the string identifying the page in question
    """

    time.sleep(2)

    driver.switch_to.window(page_handle)

    pytesseract.pytesseract.tesseract_cmd = r'C:\Users\levi.banks\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

    screengrab = pyautogui.screenshot()

    text = pytesseract.image_to_string(screengrab)

    error_messages = [
        "We are sorry, the (Proof of Delivery) documents",
        "An error occurred while processing your request"
    ]

    for error_message in error_messages:
        if error_message in text:
            return True
        
    return False

def restart_download(driver:"WebDriver", wait:"WebDriverWait", actions:"ActionChains", handle:str) -> str:
    """
    This function triggers when an error message is found. It closes the problem tab, 
    then goes back to the main download page to start the download over again. It returns 
    the most recently created handle after the download has been restarted. 

    Parameters:
        - driver: the WebDriver controlling the web browser
        - wait: a WebDriverWait object to handle waiting for web elements
        - actions: an ActionChains object, used to chain together multiple 
                   WebDriver steps when necessary
        - handle: the string identifying the page in question
    """

    driver.switch_to.window(handle)
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    close_popup = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Close"]')))
    actions.move_to_element(close_popup).perform()
    close_popup.click()
    old_handles = set(driver.window_handles)
    download_selected_PODs(driver)
    new_handles = set(driver.window_handles)-old_handles
    wait.until(lambda d: len(new_handles) > 0)
    return new_handles.pop()

def unzip_current_zips(dl_folder:str) -> "Path":
    """
    This function creates a new folder to extract all the recently downloaded
    POD PDFs into and returns the path to this new folder.

    Parameters:
        - dl_folder: the file path of the user's downloads folder
    """

    # Creating destination folder
    extracted_zips = Path(dl_folder + f'/{datetime.now().date()} PODs')

    # Looping through all the zip files returned from get_current_zips
    for zip in get_current_zips(dl_folder):
        with ZipFile(zip, 'r') as zip_temp:
            # Extract all files into extracted_zips
            zip_temp.extractall(extracted_zips)

    return extracted_zips

def get_current_zips(dl_folder:str) -> list[str]:
    """
    This function looks through the files in the downloads folder 
    and returns a list of all .zip files that were downloaded today.

    Parameters:
        - dl_folder: the path to the user's download folder
    """

    today = datetime.now().date()
    current_zips = []
    for file in os.listdir(dl_folder):

        # Getting the creation time as a datetime object for each file 
        # in Downloads
        filetime = datetime.fromtimestamp(
            os.path.getctime(dl_folder + '/' + file)
        )

        # Appending .zip files from today
        if filetime.date() == today and ".zip" in file:
            current_zips.append(file)

    return current_zips

def read_pods(folder_path:str) -> "pd.DataFrame":
    """
    This function allows the user to select a folder containing PDF versions 
    of proof of delivery documents and convert them all into a tidy DataFrame.

    Parameters: 
        - folder_path: the path to the folder containing all the files in question. 
    """

    # Setting up the returned variable with the the columns of interest
    delivered_orders_df = pd.DataFrame(columns=["Order Number", "Ship Date", "Delivery Date", "Customer Name", 'Item Number', 'Quantity'])
    
    # Going through every file in the directory to look for PDF files to convert
    for file_num, file_name in enumerate(os.listdir(folder_path)):
        # Printing progress bar
        os.system('cls')
        print("Processing: " + file_name)
        percent_processed = int(((file_num + 1) / len(os.listdir(folder_path))) * 100)
        num_pipes = '|' * percent_processed
        num_spaces = ' ' * (100 - percent_processed)
        print(f"Processed {percent_processed}% of PODs [{num_pipes}{num_spaces}]")
        
        # Clearing these variables at the start of each loop
        order_data_df = None # The variable that holds the top, non-tabular portion of the data from each PDF
        item_data_dfs = [] # The variable that holds the tables from the PDF, to be stored as DataFrames

        # Only executing code for .pdf files
        if file_name.endswith('.pdf'):
            doc_path = os.path.join(folder_path, file_name)

            # If there are any errors in opening or reading the file, pass onto the next one. 
            try:

                # Using .pdf reading library fitz to open each document
                with fitz.open(doc_path) as doc:
                    for page_num, page in enumerate(doc):
                        
                        # Noting the first page because this is where one chunk of non-tabular information to be processed always is.
                        if page_num == 0:
                            first_page = page

                        # Extracting the relevant field from the tabular portion of the PDF and storing it as a DataFrame. 
                        item_data_dfs.append(page.find_tables()[0].to_pandas())

                        # Promoting headers if the headers are set to non-data
                        if ('Order' in item_data_dfs[page_num].columns[0] and 'Details' in item_data_dfs[page_num].columns[1]):
                            new_header = item_data_dfs[page_num].iloc[0]
                            item_data_dfs[page_num] = item_data_dfs[page_num][1:]
                            item_data_dfs[page_num].columns = new_header

                        # Renaming columns for this item data table
                        item_data_dfs[page_num].columns = [
                            'Item Number', 
                            'Manufacturer Item Number', 
                            'Manufacturer', 
                            'Item Description', 
                            'Quantity'
                        ]

                        # Saving just the Item Number column for this item data table
                        item_data_dfs[page_num] = item_data_dfs[page_num].drop(columns = ['Manufacturer Item Number', 'Manufacturer', 'Item Description'])

                    # Trying to do the following, unless it runs into formatting errors, at which point it will
                    # just pass onto the next document. 
                    try:

                        # Extracting and formatting the relevant fields from the upper portion of the PDF (non-tabulated section).
                        # Storing in DataFrame.
                        order_data_df = first_page.get_text(sort = True)
                        order_data_df = order_data_df.split('\n')
                        for item in order_data_df:
                            if ': ' not in item: 
                                order_data_df.remove(item)
                        order_data_df = order_data_df[3:8]
                        for j, item in enumerate(order_data_df):
                            order_data_df[j] = item.split(": ")
                    
                        # Creating DataFrame based on dict list comprehension from list of lists created in previous line.
                        order_data_df = pd.DataFrame.from_dict({sub[0]: [sub[1]] for sub in order_data_df})

                        # Removing irrelevant columns and renaming the remaining ones
                        order_data_df.pop('Package Weight')
                        order_data_df.columns = ['Order Number', 'Ship Date', 'Delivery Date', 'Customer Name']
                        
                        # If there was only one page in the document, and item_data_dfs exists:
                        if page_num == 0 and item_data_dfs:

                            # Concatenate n copies of order_data_df, where n is equal to the total number of item entries in item_data_dfs
                            order_data_df = pd.concat([order_data_df] * (len(item_data_dfs[0])), ignore_index=True)
                            order_data_df = order_data_df.reset_index(drop=True)

                        # Else if there were two pages (and therefore two tables in item_data_dfs):
                        elif page_num == 1: 

                            # Concatenate n copies of order_data_df, where n is equal to the total number of item entries in item_data_dfs
                            order_data_df = pd.concat([order_data_df] * (len(item_data_dfs[0]) + len(item_data_dfs[1])), ignore_index=True)
                            order_data_df = order_data_df.reset_index(drop=True)
                    except:
                        pass
            except: 
                pass 

        # Collect together all the DataFrame(s) into one variable, item_data_df  
        item_data_df = None
        if len(item_data_dfs) > 1:
            item_data_df = pd.concat(item_data_dfs)
            item_data_df = item_data_df.reset_index(drop=True)
        elif len(item_data_dfs) == 1: 
            item_data_df = item_data_dfs[0]
            item_data_df = item_data_df.reset_index(drop=True)

        # Checking to see if both variables got assigned
        try:
            if (item_data_df is not None and order_data_df is not None) and\
            (len(item_data_df.columns) == 2 and len(order_data_df.columns) == 4): 

                #If so, concatenate them together horizontally and reorder the columns
                POD_df = pd.concat([order_data_df, item_data_df], axis=1)
                POD_df = POD_df[["Order Number", "Ship Date", "Delivery Date", "Customer Name", 'Item Number', 'Quantity']]
                delivered_orders_df = pd.concat([delivered_orders_df, POD_df], ignore_index=True)
                delivered_orders_df = delivered_orders_df[[
                    'Customer Name',
                    'Order Number',
                    'Item Number',
                    'Quantity',
                    'Ship Date',
                    'Delivery Date'
                ]]
        except:
            pass
        
    return delivered_orders_df
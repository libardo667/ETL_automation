from constants import *
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.remote.webdriver import WebDriver
from utils import get_code_from_inbox

def cardinal_login(driver:"WebDriver", credentials:list) -> bool:
    """
    This function takes the user's input as given from display_login() and inputs it into the
    Cardinal website's log-in page. It also deals with the two-factor authentication
    page that pops up by retrieving the code from the user's email inbox. It returns False if 
    the login failed and True if there are no issues.

    Parameters:
        - driver: the WebDriver controlling the web browser
        - credentials: the credentials of the user provided by the get_credentials function
                       in utils.py
    """

    # Navigating to the Cardinal website
    driver.get(CARDINAL_LOGIN)

    # Entering log in information and submitting
    username_field = driver.find_element(By.NAME, 'username')
    username_field.send_keys(credentials[0])
    password_field = driver.find_element(By.NAME, 'password')
    password_field.send_keys(credentials[1])
    account_number_field = driver.find_element(By.NAME, 'accountnumber')
    account_number_field.send_keys(credentials[2])
    cursor = driver.switch_to.active_element
    cursor.send_keys(Keys.TAB)
    cursor.send_keys(Keys.ENTER)

    wait = WebDriverWait(driver, 7)
    
    # After entering credentials, waits to see if a 'wrong credentials' error message
    # appears. If the user puts in the wrong credentials, return false so that a loop 
    # in the main script will prompt the user for credentials again and attempt multiple log-ins. 
    try:
        wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="form8"]/div[2]/input')))
        tfa_button = driver.find_element(By.XPATH, '//*[@id="form8"]/div[2]/input')
    except Exception:
        return False

    # Clicking the tfa_button once it appears
    tfa_button.click()

    # Getting the TFA code from my inbox, entering it, and submitting it
    tfa_code = get_code_from_inbox()
    tfa_code_entry = driver.switch_to.active_element
    tfa_code_entry.send_keys(tfa_code)
    cursor = driver.switch_to.active_element
    cursor.send_keys(Keys.TAB)
    cursor.send_keys(Keys.ENTER)

    # Sometimes Cardinal will tell you that the account is still in use when you try to log in.
    # This searches for the hyperlink that releases your account, clicks on it, 
    # if it finds it, and starts cardinal_login() again.
    try:
        wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="spnOverRide"]')))
        error_button = driver.find_element(By.XPATH, '//*[@id="spnOverRide"]')
    except:
        error_button = None
    if error_button != None:
        error_button.click()
        cardinal_login(driver, credentials)

    return True

def cardinal_log_out(driver:"WebDriver") -> None:
    """
    This function logs out of Cardinal, making future log ins faster.

    Parameters:
        - driver: the WebDriver controlling the web browser 
    """
    
    driver.switch_to.window(driver.window_handles[0])

    # Clicking the log out button
    log_out = driver.find_element(By.XPATH, '//*[@id="divMainContent"]/div/div[1]/div[1]/div/div[2]/div[1]/a')
    log_out.click()
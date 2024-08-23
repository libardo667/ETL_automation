from cardinal_login_logout import *
from constants import *
from utils import *
from reports import *
from selection import Selector

def main() -> None:
    """
    This function initializes the selectable_items variable, gets user
    input on some settings, and then defines selectable_items with 
    run_selectables_script(settings). If there are any selectable_items and
    the user chooses the "select items" setting, it will then call 
    run_selection_script(selectable_items, settings).
    """
    selectable_items = None

    display_intro()

    settings = display_settings()

    selectable_items = run_selectables_script(settings)

    if selectable_items is not None:
        run_selection_script(selectable_items, settings)

def run_selectables_script(settings:dict) -> "pd.DataFrame":
    """
    This function describes the set of steps needed to create a list of selectable items
    based on Cardinal Proof of Delivery (POD) documents and the Open Orders Details report.

    Parameters:
        - settings: the dictionary of settings chosen by the user.
    """

    # Ask the user for their downloads folder
    downloads_folder = get_downloads_folder()

    if settings['cardinal PODs']:
        # Ask the user for their Cardinal credentials 
        credentials = get_credentials(first_try=True)

    if settings['open orders report'] or settings['cardinal PODs']:
        # Engage a Selenium WebDriver in a mode that allows web automation 
        driver = engage_stealth_mode()

    if settings['open orders report']:
        # Go to the Open Orders Details Report viewer page and get the report,
        # downloading it as an Excel file
        get_open_orders_report(driver)

        # Checking to make sure the user has fully downloaded the file
        keep_download_check()

    if settings['cardinal PODs'] or settings['find selectables']:
        # Getting the list of sorted files from newest to oldest and 
        current_sorted_files = get_sorted_files(downloads_folder)
        open_orders_df = get_open_orders_from_downloads(current_sorted_files)
        pap_pin_df, headgear_orders = format_open_orders_df(open_orders_df)

    if settings['cardinal PODs']:
        # Log into Cardinal and repeat the credentialing process if needed
        successful_login = cardinal_login(driver, credentials)
        while successful_login == False:
            credentials = get_credentials(first_try=False)
            successful_login = cardinal_login(driver, credentials)

        # Get all POD PDF documents from the last 7 days
        get_PODs(driver, pap_pin_df)

        # Log out of Cardinal when done
        cardinal_log_out(driver)

    if settings['find selectables'] and settings['cardinal PODs']:
        # Process the POD PDFs into one DataFrame
        extracted_zips = unzip_current_zips(downloads_folder)
        delivered_orders_df = format_delivered_orders_df(read_pods(extracted_zips))
        delivered_orders_df.to_excel(downloads_folder + r'\Delivered Items.xlsx', index=False)

    if settings['find selectables'] and not settings['cardinal PODs']:
        delivered_orders_df = pd.read_excel(downloads_folder + r'\Delivered Items.xlsx').reset_index()
        delivered_orders_df = delivered_orders_df.astype({
            "Order Number": "str",
            "Product Code Main": 'str'
        })    

    # Format the two DataFrames and then merge them together to form a 
    # DataFrame containing all selectable items.
    selectable_items = get_selectable_items(delivered_orders_df, pap_pin_df, headgear_orders)

    selectable_items.to_excel(downloads_folder + r'\Selectable Items.xlsx') 

    if not settings['find selectables']:
        selectable_items = pd.read_excel(downloads_folder + r"\Selectable Items.xlsx", index_col=[0, 1])

    return selectable_items

def run_selection_script(selectable_items:"pd.DataFrame", settings:dict) -> None:
    """
    This function leverages selection.py and the Selector class to automate
    the order selection process. In its current, incomplete form, it only tries
    to open each order and then immediately close it. 

    Parameters:
        - selectable_items: the list of all selectable line items for the program to address
        - settings: the dictionary of settings chosen by the user.
    """
    if settings['select items']:
        selector = Selector()

        selector.open_order_entry()

        for order_number, new_df in selectable_items.groupby(level=0):

            selector.get_order_line_items(str(order_number), new_df.droplevel(0).reset_index())
            selector.close_order()

            #print(order_number, ": \n", new_df.droplevel(0).reset_index())

# Starting the script    
if __name__ == "__main__":
    main()
    
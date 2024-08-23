import pandas as pd
from datetime import datetime, date
import constants

def format_open_orders_df(df:"pd.DataFrame") -> tuple["pd.DataFrame"]:
    """
    This function formats the Open Orders Details report to prepare it for merging
    with the Proof of Delivery DataFrame. It formats the DataFrame and then filters
    out headgear orders as a separate DataFrame before returning the two of them 
    as in a tuple.

    Parameters:
        - df: the DataFrame being manipulated
    """ 

    # Values were stored as floats, converting to integers to truncate decimal portion
    df['Order '] = df['Order '].astype('Int64')
    df['Invy Loc'] = df['Invy Loc'].astype('Int64')

    # Selecting just the relevant columns in a nice order
    df = df[[
        'CusNo',
        'Patient Name',
        'Order ',
        'Product Category',
        'Product Code',
        'Invy Loc', 
        'Initials',
        'Line Selection Status', 
        'Create Date'
    ]]

    # Renaming a couple columns
    df = df.rename(columns={'CusNo': 'Customer Number', 'Order ': 'Order Number'})
    
    # Type casting the columns accordingly
    df = df.astype({
        'Customer Number': 'str',
        'Patient Name': 'str',
        'Order Number': 'str',
        'Product Category': 'str',
        'Product Code': 'str',
        'Invy Loc': 'str',
        'Initials': 'str',
        'Line Selection Status': 'str',
        'Create Date': 'datetime64[ns]'
    })

    # Setting up a new column with temporary values
    df["Product Code Main"] = df['Product Code']
    
    # Modifying the product code column to more closely match what is represented
    # in Cardinal
    for key, value in constants.MISMATCHED_NAMES.items():
        df.loc[df["Product Code"] == key, "Product Code Main"] = value

    # Creating a new field to be used for merging (last 5 characters of Product Code)
    df['Product Code Main'] = df['Product Code Main'].str[-5:]

    # Filtering the DataFrame to extract just 104 PAP PIN orders and the line items for
    # headgear from within that subset. Returning these DataFrames
    df, headgear_orders = filter_pap_pin(df)
    return df, headgear_orders

def filter_pap_pin(df:"pd.DataFrame") -> tuple["pd.DataFrame"]:
    """
    This function applies filters to the given DataFrame that filter for 104 PAP PIN orders. 
    It also creates a new DataFrame of just the line items for headgear. It returns both DataFrames.
    
    Parameters:
        - df: the DataFrame being manipulated
    """

    # Setting filter conditions
    headgear_filter = df['Product Code'].isin(['HCS HEADGEAR', 'HCS A7035'])
    pap_pin_filters = (df['Invy Loc'] == '104') & \
                      (df['Product Category'].isin(['CPAP BIPAP ACC', 'RESPIRATORY'])) & \
                      (df['Initials'] == 'PIN') & \
                      (df['Line Selection Status'] == 'No')
    
    # Assigning/re-assigning the DataFrames to their new values and returning
    df_filtered = df[pap_pin_filters]
    headgear_orders = df_filtered.loc[headgear_filter]
    return df_filtered, headgear_orders

def format_delivered_orders_df(df:"pd.DataFrame") -> "pd.DataFrame":
    """
    This function formats the delivered_orders_df returned by the read_pods 
    function in utils.py.

    Parameters: 
        - df: the DataFrame being formatted
    """

    # Selecting just the following columns in this order from the given df
    df = df[[
        'Customer Name',
        'Order Number',
        'Item Number',
        'Quantity',
        'Ship Date',
        'Delivery Date'
    ]]

    # Type casting the fields
    df = df.astype({
        'Customer Name': 'str',
        'Order Number': 'str',
        'Item Number': 'str',
        'Quantity': 'int64',
        'Ship Date': 'datetime64[ns]',
        'Delivery Date': 'datetime64[ns]'
    })

    # Creating a new field to be used for merging (last 5 characters of Item Number)
    df['Product Code Main'] = df['Item Number'].str[-5:]
    df = df.reset_index()

    return df

def get_selectable_items(delivered:"pd.DataFrame", 
                         pap_pin:"pd.DataFrame", 
                         headgear:"pd.DataFrame") -> "pd.DataFrame":
    """
    This function merges delivered_orders_df with pap_pin_df to show a list of 
    line items that have been delivered and that can be selected. It then adds
    the headgear orders back into the mix and returns the selectable_items. 

    Parameters:
        - delivered: the DataFrame containing all the information about every delivered
                     order from the date ranges searched
        - pap_pin: the DataFrame containing all the information about Healthcall orders
                   for PAP supplies
        - headgear: the DataFrame containing the information about headgear line items
    """

    # Left joining pap_pin with delivered on the appropriate columns
    selectable_items = pap_pin.merge(delivered,
                                    how='left',
                                    on=['Order Number', 'Product Code Main'])
    
    # Dropping line items that have no delivery date associated with them
    selectable_items = selectable_items.dropna(subset=['Delivery Date'])

    # Concatenating the headgear orders to the bottom of selectable_items
    selectable_items = pd.concat([selectable_items, headgear])

    # Selecting the relevant columns in a nice order
    selectable_items = selectable_items[[
        'Ship Date', 
        'Delivery Date',
        'Quantity',
        'Customer Number',
        'Patient Name',
        'Order Number',
        'Product Code'
    ]]

    # Dropping duplicates based on Order Number and Product Code
    selectable_items = selectable_items.drop_duplicates(subset=['Order Number', 'Product Code'])

    # Sorting by ascending (Order Number, Product Code)
    selectable_items = selectable_items.sort_values(['Order Number', 'Product Code'], 
                                                    ascending=[True, True])

    # Filling in the quantity, ship date and delivery date for headgear items 
    selectable_items.loc[
        selectable_items['Product Code'].isin(['HCS HEADGEAR', 'HCS A7035']), 
        'Quantity'
    ] = selectable_items['Quantity'].fillna(1)

    selectable_items.loc[
        selectable_items['Product Code'].isin(['HCS HEADGEAR', 'HCS A7035']), 
        'Ship Date'
    ] = selectable_items['Ship Date'].bfill(limit=1)

    selectable_items.loc[
        selectable_items['Product Code'].isin(['HCS HEADGEAR', 'HCS A7035']), 
        'Delivery Date'
    ] = selectable_items['Delivery Date'].bfill(limit=1)

    # Setting the index to be a multi-index of Order number and Product Code
    selectable_items = selectable_items.set_index(["Order Number", "Product Code"])
    
    return selectable_items
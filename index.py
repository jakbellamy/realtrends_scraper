import pandas as pd
from datetime import datetime

# Set base URL
base_url = 'https://www1.realtrends.com/best-real-estate-agents-'

# Set the states
states = ['georgia', 'florida', 'north-carolina']

# Set the subsets
data_subsets = ['individuals', 'teams-small',
                'teams-medium', 'teams-large', 'teams-mega']

# Set transaction types
transaction_types = ['sides', 'volume']

# function to convert currency string into float
def convert_currency_to_float(currency_string):
    try:
        # Remove the '$' and ',' from the string and convert the string to a float
        currency_string = currency_string.replace('$', '').replace(',', '')
        currency_float = float(currency_string)
    except:
        currency_float = 0  # If the string is empty, set the float to 0
    return currency_float

# function to pull the data from the website and return a dataframe
def get_data(state, data_subset, transaction_type):
    # Define the URLs. Only add '-by-' if the data_subset is 'individuals' else add '-'
    if data_subset == 'individuals':
        target_url = base_url + state + '/' + data_subset + '-by-' + transaction_type
    else:
        target_url = base_url + state + '/' + data_subset + '-' + transaction_type
    # Read website's table into a dataframe
    df = pd.read_html(target_url)[0]
    df.drop(columns=['Website'], inplace=True)
    # Return the dataframe
    return df

#  function that calls both transaction types and joins them on Team Name
def get_both_transaction_types(state, data_subset):
    # Get the data for both transaction types
    by_sides_df = get_data(state, data_subset, 'sides')
    by_volume_df = get_data(state, data_subset, 'volume')

    # If the data_subset == 'individuals', set a new column Full Name to the value of 'First Name' + 'Last Name'
    if data_subset == 'individuals':
        by_sides_df['Full Name'] = by_sides_df['First Name'] + \
            ' ' + by_sides_df['Last Name']
        by_sides_df.drop(columns=['First Name', 'Last Name'], inplace=True)
        by_volume_df['Full Name'] = by_volume_df['First Name'] + \
            ' ' + by_volume_df['Last Name']
        by_volume_df.drop(columns=['First Name', 'Last Name'], inplace=True)

    # Join the dataframes on Team Name (unless data_subset == 'individuals', then by 'Full Name')
    if data_subset == 'individuals':
        df = by_sides_df.join(by_volume_df.set_index(
            'Full Name')['Volume'], on='Full Name')
    else:
        df = by_sides_df.join(by_volume_df.set_index(
            'Team Name')['Volume'], on='Team Name')

    # Convert the 'Volume' column to a float
    df['Volume'] = df['Volume'].apply(convert_currency_to_float)

    # If the data_subset == 'individuals', move the 'Full Name' column to the front of the dataframe
    if data_subset == 'individuals':
        cols = list(df.columns)
        cols.remove('Full Name')
        cols = ['Full Name'] + cols
        df = df[cols]

    # reset the rank column based on volume (descending) then transactions (descending) then set the index to the Rank column
    # df.sort_values(by=['Volume', 'Transactions'], ascending=False, inplace=True) # This is the original sort. Modified because not every agent/team shows volume??
    df.sort_values(by=['Transactions', 'Volume'],
                   ascending=False, inplace=True)
    df['Rank'] = range(1, len(df) + 1)
    df.set_index('Rank', inplace=True)

    return df

# function that creates an excel workbook with a worksheet for each state and data_subset
def create_excel_workbook(states=states, data_subsets=data_subsets):
    # Create an excel writer object
    file_name = 'real_trends ' + str(datetime.now())[:10] + '.xlsx'
    writer = pd.ExcelWriter(
        '/Users/jakobbellamy/Documents/Reports/real_trends_scrapes/' + file_name, engine='xlsxwriter')

    for state in states:
        for data_subset in data_subsets:
            # Get the data for the state and data_subset
            df = get_both_transaction_types(state, data_subset)
            # Write the dataframe to the excel writer object
            df.to_excel(writer, sheet_name=state + '_' + data_subset)

    # format the excell workbook so that columns are automatically as wide as the data and set the currency format for the Volume column
    workbook = writer.book
    for worksheet in workbook.worksheets():
        # Adjust the width of the columns based on pre-defined column widths
        worksheet.set_column('A:A', 6)
        worksheet.set_column('B:B', 25)
        worksheet.set_column('C:C', 30)
        worksheet.set_column('D:D', 15)
        worksheet.set_column('E:E', 10)
        worksheet.set_column('F:F', 10)
        # Set the Volume column to have a currency format
        worksheet.set_column(
            'G:G', 20, workbook.add_format({'num_format': '$#,##0'}))

    workbook.close()

create_excel_workbook() # Run the function

# Notes
# All Possible States
# states = ['alabama', 'alaska', 'arizona', 'arkansas', 'california', 'colorado', 'connecticut', 'delaware', 'florida', 'georgia', 'hawaii', 'idaho', 'illinois', 'indiana', 'iowa', 'kansas', 'kentucky', 'louisiana', 'maine', 'maryland', 'massachusetts', 'michigan', 'minnesota', 'mississippi', 'missouri', 'montana', 'nebraska', 'nevada', 'new-hampshire', 'new-jersey', 'new-mexico', 'new-york', 'north-carolina', 'north-dakota', 'ohio', 'oklahoma', 'oregon', 'pennsylvania', 'rhode-island', 'south-carolina', 'south-dakota', 'tennessee', 'texas', 'utah', 'vermont', 'virginia', 'washington', 'west-virginia', 'wisconsin', 'wyoming']

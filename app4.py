from pdfminer.high_level import extract_text
import re
import re
from datetime import datetime
import openpyxl
import os
import streamlit as st
import time
import shutil
import tempfile
import pandas as pd
import base64 
import requests
from openpyxl.styles import Alignment, Border, Side,Font, PatternFill, NamedStyle, numbers
import streamlit as st
import base64
from io import BytesIO

def get_platform_type(text):
    #Entra
    if "please do revert for any further clariﬁcations." in text.lower():
        return "Entravision"
    elif "https://www.addynamo.com/payus.php" in text.lower():
        return "Twitter"
    
    elif "eskimi pte limited" in text.lower():
        return "Eskimi"  
    
    else:
        return "Unrecognizable Invoice"


def extract_entravision(text):

    # Define regex patterns
    invoice_no_pattern = r"INVOICE NO\.\s+(\d+)"
    billing_period_pattern = r"BILLING PERIOD\s+([A-Za-z]+-\d{4})"
    rate_pattern = r"The USD/GHS rate used is GHs (\d+\.\d+)/\$"
    
    # Extract global details
    invoice_no_match = re.search(invoice_no_pattern, text)
    billing_period_match = re.search(billing_period_pattern, text)
    rate_match = re.search(rate_pattern, text)
    
    invoice_no = invoice_no_match.group(1) if invoice_no_match else "Invoice number not found"
    billing_period = billing_period_match.group(1) if billing_period_match else "Billing period not found"
    rate = rate_match.group(1) if rate_match else "Rate not found"
  


    # Extract text between 'DESCRIPTION' and 'USD'
    description_match = re.search('DESCRIPTION(.*?)USD', text, re.DOTALL)
    if not description_match:
        return 'description finding'
    
    description_text = description_match.group(1)
    
    # Process each line for item descriptions
    lines = description_text.split('\n')
    items = []
    temp_item = ""
    
    for line in lines:
        line = line.strip()
        if not line:  # Skip empty lines
            continue
        if any(substring in line for substring in ['BRD_CLBB', 'BRD_CLBS', 'BRD_BTM' ,'GinoTomatoMix', 'POMOTomatoMix']) or temp_item == "":
            if temp_item:  # Save the previous item if it exists
                items.append({'item_line': temp_item})
                temp_item = line
            else:
                temp_item = line
        else:
            temp_item = f"{temp_item} {line}" if temp_item else line
    
    if temp_item:  # Add the last item if it exists
        items.append({'item_line': temp_item.strip()})
    
    # Extract USD 
    usd_match = re.search('USD(.*?)Remarks / Payment Instructions:', text, re.DOTALL)
    if not usd_match:
        return []  # Return an empty list if usd_match fails
    
    usd_text = usd_match.group(1)
    usds = [usd.strip() for usd in usd_text.split('\n') if usd.strip()]
    
    # Check if the number of usd matches the number of items
    if len(usds) != len(items):
        return usds , items
    
    # Add usd and brand to the items
    for item, usd in zip(items, usds):
        item['platform'] ='Meta'
        item['usd'] = usd
        
        if 'BRD_CLBB' in item['item_line']:
            item['brand'] = 'Club Beer'
        elif 'BRD_CLBS' in item['item_line']:
            item['brand'] = 'Club Shandy'
        elif 'BRD_BTM' in item['item_line']:
            item['brand'] = 'Beta Malt'

        elif 'GinoTomatoMix' in item['item_line']:
            item['brand'] = 'Gino Tomato Mix'
        elif 'POMOTomatoMix' in item['item_line']:
            item['brand'] = 'Pomo Tomato Mix'

        # Add global details
        item['invoice_no'] = invoice_no
        item['billing_period'] = billing_period
        item['rate'] = 13.5
        item['impressions'] = 0

    
    return items

def extract_twitter(text):
    # Extract Invoice Date and Invoice Number , he exchange rate between GBP and USD using regex
    invoice_date_match = re.search(r"Invoice Date\n(\d{2} \w{3} \d{4})", text)
    invoice_number_match = re.search(r"Invoice Number\n(.*?)\n", text)
    rate_match = re.search(r"1\s*GBP\s*=\s*(\d+\.\d+)\s*USD", text)

    # Initialize variables to hold the extracted data
    extracted_date = invoice_date_match.group(1) if invoice_date_match else None
    extracted_number = invoice_number_match.group(1) if invoice_number_match else None
    extracted_rate = rate_match.group(1) if rate_match else None

    # Find the description section
    description_start_index = text.find("Twitter")
    description_end_index = text.find("*GBP Equivalent")
    description_section = text[description_start_index:description_end_index].strip()

    # Process item lines in the description section
    items = []
    current_item_lines = []
    for line in description_section.split('\n'):
        if line.strip() == "":  # Check for empty line signaling end of current item
            if current_item_lines:
                items.append(" ".join(current_item_lines).strip())
                current_item_lines = []
        else:
            current_item_lines.append(line.strip())
    # Add the last item if not empty
    if current_item_lines:
        items.append(" ".join(current_item_lines).strip())

    # Find USD amounts
    usd_amounts = re.findall(r"(\d{1,3}(?:,\d{3})*\.\d{2,}) Zero Rated", text)
    
    # Match items with USD amounts (assuming the order matches)
    matched_items = []
    for item, amount in zip(items, usd_amounts):

        # Determine the brand based on the item line content
        if '@BetaMaltGhana' in item:
            brand = 'Beta Malt'
        elif '@Chale_Club' in item:
            brand = 'Club Beer'
        elif '@clubshandybosoe' in item:
            brand = 'Shandy'
        elif '@budweiserghana' in item:
            brand = 'Budweiser'                
        else:
            brand = 'Unidentified Brand'

        matched_items.append({
            'item_line': item,
            'platform': 'Twitter',
            'usd': amount,
            'brand': brand,
            'invoice_no': extracted_number,
            'billing_period': extracted_date,
            'rate': 13.5,  # Include the exchange rate in each item
            'impressions': 0
        })

    return matched_items

def extract_eskimi(text):
    # Extract Invoice Number
    invoice_number_match = re.search(r"TAX INVOICE No\. (\w+ \d+)", text)
    invoice_number = invoice_number_match.group(1) if invoice_number_match else None

    # Extract Billing Period
    billing_period_match = re.search(r"Date: (\d{4}-\d{2}-\d{2})", text)
    billing_period = billing_period_match.group(1) if billing_period_match else None

    # Extract Item Line (Service Details)
    service_details_start_index = text.find("Service details:") + len("Service details:")
    channel_index = text.find("Channel /")
    item_line = text[service_details_start_index:channel_index].strip()

    # Determine the brand based on the item line (case insensitive)
    item_line_lower = item_line.lower()  # Convert to lowercase to make matching case insensitive
    if 'beta malt' in item_line_lower:
        brand = 'Beta Malt'
    elif 'club shandy' in item_line_lower:
        brand = 'Club Shandy'
    elif 'club beer' in item_line_lower:
        brand = 'Club Beer'
    else:
        brand = 'Unknown'  # Default value if no known brand is found

    # Find USD and Impressions based on 'CPM' occurrence
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if 'CPM' in line and line.strip() == 'CPM':  # Ensure 'CPM' stands alone on the line
            # Extract Impressions from the 4th line after 'CPM'
            impressions_line = lines[i + 4] if (i + 4) < len(lines) else None
            impressions = re.search(r"(\d{1,3}(?:,\d{3})*)", impressions_line)
            impressions = impressions.group(1).replace(',', '') if impressions else None
            
            # Extract USD from the 6th line after 'CPM' (2 lines after Impressions)
            usd_line = lines[i + 6] if (i + 6) < len(lines) else None
            usd = re.search(r"(\d{1,3}(?:,\d{3})*(?:\.\d+)?|\.\d+)", usd_line)
            usd = usd.group(1).replace(',', '') if usd else None
            break  # Assuming only one 'CPM' per document that needs processing
    # Initialize a list to hold all row dictionaries
    all_rows = []



    all_rows.append({
        'item_line': item_line,
        'platform': 'eskimi',
        'usd': usd,
        'brand': brand,
        'invoice_no': invoice_number,
        'billing_period': billing_period,
        'rate': 13.5,  # Include the exchange rate in each item
        'impressions': impressions
    })    

    return all_rows


#_--------
def format_checker(data_list):
    required_keys = ["item_line", "platform", "usd", "brand", "invoice_no", "billing_period", "rate", "impressions"]
    
    # Iterate in reverse to safely remove elements while iterating
    for i in range(len(data_list) - 1, -1, -1):
        item = data_list[i]
        
        # Check if the item is a dictionary and contains all required keys
        if isinstance(item, dict) and all(key in item for key in required_keys):
            continue  # This item is fine, move to the next one
        else:
            # If the item is not a dictionary or is missing keys, remove it
            del data_list[i]

def convert_numbers_to_float(dlist):
    for item in dlist:
        # Handle 'usd' conversion
        if 'usd' in item and isinstance(item['usd'], str):
            amount_str = item['usd'].replace(',', '')
            try:
                item['usd'] = float(amount_str)
            except ValueError:
                item['usd'] = 0.0  # Handle invalid input gracefully

        # Handle 'impressions' conversion
        if 'impressions' in item and isinstance(item['impressions'], str):
            impressions_str = item['impressions'].replace(',', '')
            try:
                item['impressions'] = int(impressions_str)
            except ValueError:
                item['impressions'] = 0  # Handle invalid input gracefully

        # Handle 'rate' conversion
        if 'rate' in item and isinstance(item['rate'], str):
            rate_str = item['rate'].replace(',', '')
            try:
                item['rate'] = float(rate_str)
            except ValueError:
                item['rate'] = 0.0  # Handle invalid input gracefully

    return dlist

#Clean item_line
def process_item_lines(data):
    for item in data:
        if isinstance(item, list):
            print("Found a list item:", item)
            continue  # Skip processing this item
        # Define item_line at the beginning to ensure it's always available
        item_line = item.get('item_line', '')

        if item.get('platform') == 'Meta':
            # Cut _PARTICIPATION_ and all text after it
            participation_index = item_line.find('_PARTICIPATION_')
            if participation_index != -1:
                item_line = item_line[:participation_index]

            # Modify here to remove the specified text and all text that comes before it
            for substring in ['GHA_BRD_BTM_', 'GHA_BRD_BUD_', 'GHA_BRD_CLBB_', 'GHA_BRD_CLBS_']:
                index = item_line.find(substring)
                if index != -1:
                    item_line = item_line[index + len(substring):]


            # Additional substring removals based on your previous code
            pomo_index = item_line.find('GBFoods_POMOTomatoMix_')
            if pomo_index != -1:
                item_line = item_line[pomo_index + len('GBFoods_POMOTomatoMix_'):]

            gino_index = item_line.find('GBFoods_GinoTomatoMix_')
            if gino_index != -1:
                item_line = item_line[gino_index + len('GBFoods_GinoTomatoMix_'):]

        # This now works even if 'entravision' condition is not met
        if item.get('platform') == 'Twitter':
            for substring in ['- @clubshandybosoe -', '- @Chale_Club -', '- @BetaMaltGhana -']:
                handle_index = item_line.find(substring)
                if handle_index != -1:
                    item_line = item_line[handle_index + len(substring):]

        if item.get('platform') == 'eskimi':
            last_dash_index = item_line.rfind('-')
            if last_dash_index != -1:
                item_line = item_line[:last_dash_index].strip()

            first_dash_index = item_line.find('-')
            if first_dash_index != -1:
                item_line = item_line[first_dash_index + 1:].strip()

        # Update the item_line after processing
        item['item_line'] = item_line

    return data
#Collapse/sum/merge same lines
def collapser(data):
    # Temporary storage for sums
    temp_storage = {}
    # Final list to return
    collapsed_list = []

    for item in data:
        # Create a unique key for each combination
        key = (item['item_line'], item['invoice_no'], item['brand'], item['billing_period'])

        # If the key is already in temp_storage, update the sums
        if key in temp_storage:
            temp_storage[key]['usd'] += item['usd']
            temp_storage[key]['impressions'] += item['impressions']
        else:
            # Otherwise, add the item to temp_storage
            temp_storage[key] = item

    # Reconstruct the collapsed list from temp_storage
    for key, value in temp_storage.items():
        collapsed_list.append(value)

    return collapsed_list
#Tax Generator
def tax_gen(data):
    for item in data:
        platform = item.get('platform')

        # Calculate 'ghc' as 'rate' * 'usd', ensuring both exist and are numeric
        if 'rate' in item and 'usd' in item:
            try:
                item['ghc'] = item['rate'] * item['usd']
            except (TypeError, ValueError):
                # Handle cases where 'rate' or 'usd' cannot be multiplied
                item['ghc'] = 0

        # Assign tax-related values based on platform
        if platform in ['Meta', 'Ghana Web']:
            item.update({
                'vol_discount': 0,
                'agency_comm': 0,
                'GETFL': '2.5%',
                'NHIL': '2.5%',
                'covid': '1%',
                'VAT': '14.99997734%'
            })
        elif platform in ['Eskimi', 'Twitter', None]:  # Includes None for unspecified platforms
            item.update({
                'vol_discount': 0,
                'agency_comm': 0,
                'GETFL': 0,
                'NHIL': 0,
                'covid': 0,
                'VAT': 0
            })

    return data

def excel_recons(data):
    # Load the template workbook and select Sheet1
    workbook = openpyxl.load_workbook('carat-recons-template.xlsx')
    sheet = workbook['Sheet1']
    
    # Populate the sheet with values from the dictionary
    sheet['A12'] = data.get('platform', '')
    sheet['B5'] = data.get('brand', '')
    sheet['B6'] = data.get('item_line', '')
    sheet['B8'] = data.get('platform', '')
    sheet['B28'] = data.get('rate', '')
    sheet['B9'] = data.get('billing_period', '')
    sheet['B10'] = data.get('invoice_no', '')
    sheet['C13'] = data.get('impressions', 0)
    sheet['C16'] = data.get('vol_discount', 0)
    sheet['C18'] = data.get('agency_comm', 0)
    sheet['C20'] = data.get('GETFL', 0)
    sheet['C21'] = data.get('NHIL', 0)
    sheet['C22'] = data.get('covid', 0)
    sheet['C24'] = data.get('VAT', 0)
    sheet['G13'] = data.get('ghc', 0)
    
    # Save the workbook with a new name based on the dictionary values
    filename = f"{data.get('brand', 'Unknown')}_{data.get('item_line', 'Unknown')}_{data.get('platform', 'Unknown')}_Recons.xlsx"
    filename = sanitize_filename(filename)
    workbook.save(filename)
    
    #return save_excel_to_memory(workbook)

#-------
def get_base64_of_image(file_path):
    with open(file_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode()

def create_download_link(excel_mem, file_label):
    """Generate a download link that Streamlit can display, for the given in-memory Excel file."""
    b64 = base64.b64encode(excel_mem.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{file_label}">Download {file_label}</a>'
    return href

def sanitize_filename(filename):
    invalid_chars = "/\\?%*:|\"<>\n"
    for char in invalid_chars:
        filename = filename.replace(char, "_")
    return filename





# Streamlit app
def main():        
    st.set_page_config(layout="wide")

    # Convert your image to base64
    image_path = 'bg-hd.jpg'  # Path to your local image
    base64_image = get_base64_of_image(image_path)

     
    st.sidebar.title("REGEN")
    st.header("Reconcilation Generator")
    


    uploaded_files = st.sidebar.file_uploader("Upload PDF Invoice(s)", type=["pdf"], accept_multiple_files=True)
   
    st.image('logostrip.png')
    if uploaded_files:
        all_results = []  # Store results for all uploaded files
        
        for uploaded_file in uploaded_files:
            # Get the brand name from the uploaded file's name

            # Correctly use the pdf_file_path to extract text from each PDF file
            text = extract_text(uploaded_file)

            # Determine the platform type
            platform_type = get_platform_type(text)
            print(platform_type)
           
            #Extractors
            if platform_type == 'Entravision':
                st.subheader('Entravision support is Depracated') 
                #print(result)
                #all_results.extend(result)

            elif platform_type == 'Twitter':
                result = extract_twitter(text)
                #print(result)
                all_results.extend(result)

            elif platform_type == 'Eskimi':
                result = extract_eskimi(text)
                print(text)
                print(result)
                all_results.extend(result)
              

            elif platform_type == 'Unrecognizable Invoice':
                st.header('Unrecognizable Invoice')
        

        #all_results
        print('------  Here RAW ---------')
        print(all_results)
        print('------       ---------')
        format_checker(all_results)
        all_results = convert_numbers_to_float(all_results)
        all_results = process_item_lines(all_results)
        all_results = collapser(all_results)
        all_results = tax_gen(all_results)
        print('         ')
        print('------  Processed ---------')
        print(all_results)
        print('------       ---------')        


        download_links = {}  # Store download links
        for index, result in enumerate(all_results):
            # Populate and save an Excel file for each dictionary
            excel_recons(result)  # This saves the file directly
            
            sanitized_item_line = sanitize_filename(result.get('item_line', 'Unknown'))
            excel_file_name = f"{result.get('brand', 'Unknown')}_{sanitized_item_line}_{result.get('platform', 'Unknown')}_Recons.xlsx"            
            # Now, read the saved Excel file into memory
            try:
                with open(excel_file_name, "rb") as excel_file:
                    in_memory_excel = BytesIO(excel_file.read())
                
                # Generate a download link for the in-memory Excel file
                download_link = create_download_link(in_memory_excel, excel_file_name)
                download_links[excel_file_name] = download_link
            except FileNotFoundError:
                st.error(f"File not found: {excel_file_name}")

        # Display download buttons for each Excel file
        st.subheader("Download Excel Files:")
        for label, link in download_links.items():
            button_label = f"Download {label} Excel File"
            if st.button(button_label):
                # Display the download link for the clicked file
                st.markdown(link, unsafe_allow_html=True)                       


if __name__ == "__main__":
    main()

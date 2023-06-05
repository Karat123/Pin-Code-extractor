import pandas as pd
import re
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
no=0
# Path to ChromeDriver executable
chromedriver_path = '/path/to/chromedriver'

# Read the Excel file
input_file = 'Blank pincodes NGO datapoint 9000.xlsx'
input_sheet = 'Sheet1'

# Load the input Excel file into a DataFrame
df_input = pd.read_excel(input_file, sheet_name=input_sheet)

# Create a new DataFrame to store the output
df_output = pd.DataFrame(columns=['Address', 'Pin Code'])

# Create a WebDriver instance
driver = webdriver.Chrome(service=Service(chromedriver_path))

# Function to search for pin code, update the output DataFrame, and delete the cell
def search_and_update(row):
    address = row['full_address']
    query = f"{address} pin code"
    driver.get("https://www.google.com/")
    search_input = driver.find_element(By.NAME, "q")
    search_input.send_keys(query)
    search_input.submit()
    try:
        snippet_element = driver.find_element(By.CSS_SELECTOR, "div.g > div > div > div > span")
        snippet_text = snippet_element.text
        pin_code_match = re.search(r"\b\d{6}\b", snippet_text)
        if pin_code_match:
            pin_code = pin_code_match.group()
        else:
            pin_code = ""
        df_output.loc[row.name] = [address, pin_code]
        df_input.drop(row.name, inplace=True)
        # Save the output DataFrame to a new Excel file after each update
        output_file = 'output_data1.xlsx'
        df_output.to_excel(output_file, index=False)
    except:
        pin_code = ""
        df_output.loc[row.name] = [address, pin_code]
        df_input.drop(row.name, inplace=True)
        # Save the output DataFrame to a new Excel file after each update
        output_file = 'output_data1.xlsx'
        df_output.to_excel(output_file, index=False)

# Iterate over the rows of the input DataFrame
for _, row in df_input.head(1500).iterrows():
    search_and_update(row)
    no+=1
    print(no)
    time.sleep(2)  # Delay for 2 seconds


# Close the WebDriver
driver.quit()

# Save the updated input DataFrame to the same input file
df_input.to_excel(input_file, index=False)

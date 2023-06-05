# Pin-Code-extractor
The provided code is a Python script that utilizes Selenium and Pandas libraries to extract pin codes from addresses. Here's a summary of the code:

The required libraries are imported: pandas, re, time, selenium, and the necessary modules from selenium.webdriver.

The path to the ChromeDriver executable is specified using the chromedriver_path variable.

An Excel file (input.xlsx) is read into a Pandas DataFrame (df_input) with a specified sheet (Sheet1).

Another DataFrame (df_output) is created to store the output, containing columns for the address and pin code.

A WebDriver instance is created using the ChromeDriver path.

The search_and_update function is defined, which takes a row from the input DataFrame, searches for the pin code using Google search, updates the output DataFrame with the address and pin code, and removes the processed row from the input DataFrame. The output DataFrame is saved to an Excel file (output_data1.xlsx) after each update.

The script iterates over the first 1500 rows of the input DataFrame using df_input.head(1500).iterrows(). For each row, the search_and_update function is called, and a delay of 2 seconds is added after each iteration using time.sleep(2).

Once the iteration is complete, the WebDriver is closed using driver.quit().

The updated input DataFrame (df_input) is saved back to the same input file (input.xlsx).

To use this code, you need to ensure that you have installed the necessary libraries (pandas, selenium) and have the ChromeDriver executable (chromedriver) available at the specified path. You also need to provide the correct input file path and sheet name in the code.

The script can be run to extract pin codes from addresses in an Excel file and store the results in another Excel file.

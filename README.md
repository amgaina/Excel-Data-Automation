# Excel-Data-Automation

This Python script automates the process of applying a 10% discount to data in an Excel spreadsheet and adding a bar chart to visualize the corrected data. It's particularly useful when dealing with large datasets where manual correction is impractical.

## Problem Statement
An employee of the company mistakenly forgets to put a 10% discount on the price in the Excel spreadsheet. However, millions of data entries have already been saved without the discount. The task is to automate the correction of cell data and visualize the corrected data with a bar chart without the need for manual intervention.

## Solution
The provided Python script utilizes the `openpyxl` library to interact with Excel spreadsheets. It performs the following tasks:
1. Loads the specified Excel file.
2. Iterates through each row in the spreadsheet, applying a 10% discount to the values in column 6.
3. Adds a bar chart to visualize the corrected data in column 6.
4. Saves the modified Excel file with the corrected data and the added chart.

## Usage
1. Ensure you have Python installed on your system.
2. Install the `openpyxl` library using pip:
    ```
    pip install openpyxl
    ```
3. Place your Excel file (e.g., "Book1.xlsx") in the same directory as the script.
4. Run the script by executing the following command in your terminal or command prompt:
    ```
    python script_name.py
    ```
   Replace `script_name.py` with the name of your Python script.
5. The script will automatically apply the 10% discount to the data and add a bar chart to visualize the corrected values in the specified Excel file.

## Dependencies
- Python 3.x
- openpyxl library


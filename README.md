# VBA-challenge

# Data Summary VBA Script

This project contains a VBA script designed to summarize stock data within an Excel workbook. The script adds headers, calculates quarterly changes, formats the data, and identifies key stock performance metrics such as the greatest percentage increase, greatest percentage decrease, and greatest total volume.

## Table of Contents
- [Installation](#installation)
- [Usage](#usage)
- [Code](#code)
- [Contributing](#contributing)
- [License](#license)

## Installation
To use this script, follow these steps:
1. Open the Excel workbook where you want to run the script.
2. Press `ALT + F11` to open the Visual Basic for Applications (VBA) editor.
3. Insert a new module by clicking `Insert > Module`.
4. Copy and paste the script from `Data_Summary()` into the new module.
5. Press `CTRL + S` to save the script.

## Usage
1. Ensure your data is structured with the following columns:
    - Column A: Ticker
    - Column C: Open Price
    - Column F: Close Price
    - Column G: Volume
2. Run the script by pressing `F5` within the VBA editor or by creating a button within your worksheet to trigger the macro.

## Code
The script performs the following actions:
- Iterates through each worksheet in the workbook.
- Adds headers in columns J through M for `Ticker`, `Quarterly Change`, `Percent Change`, and `Total Stock Volume`.
- Calculates the quarterly change and percent change for each stock ticker.
- Formats the results by changing cell colors based on the value of the quarterly change.
- Identifies and records the greatest percentage increase, greatest percentage decrease, and greatest total volume for each stock ticker.
- Adjusts column sizes for better readability.

## Contributing
Contributions are welcome. If you have suggestions or improvements, please open an issue or submit a pull request.

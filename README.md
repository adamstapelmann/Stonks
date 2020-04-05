# Stonks
This provides custom scripting for a Google Sheets spreadsheet to track stock metrics. The script creates a template to input buy and sell information for stocks, and it aggregates and summarizes metrics on a dashboard sheet.

## Setup
1. Create a new Google Sheets project. On your Google Drive, right click and create a blank spreadsheet. Name it however you like.
2. At the bottom of your spredsheet, add two new sheets to your project. Rename the three sheets so that they are "Dashboard", "Bought" and "Sold". Code changes need to be made to support renaming these sheets.
3. In the toolbar of your project, select Tools > Script Editor. Copy paste the code from stonks.js into the Code.gs that is generated for the project. 

## How to use
- Input every purchase of a stock as a new line in the Bought sheet. The same stock can appear here multiple times for different purchases.
- Input every sale of a stock as a new line in the Sold sheet. The same stock can appear here multiple times as well.
- The entire dashboard should be generated, although it might require refreshing the page for information to show up.

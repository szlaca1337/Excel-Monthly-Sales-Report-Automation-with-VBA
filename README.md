
# Monthly Report Automation with VBA

## Project Overview
This project is an Excel-based automation tool designed to streamline the process of consolidating monthly raw data, extracting month names, and updating a master sheet. It includes a dynamic dashboard with slicers and charts for interactive reporting. All processes are automated via VBA, including PivotTable and chart refreshes.

## Features
- **Data Processing**: Extracts month names from sheet names and adds them to the data.
- **Data Consolidation**: Merges data from multiple sheets into a master sheet (`AllData`).
- **Automation**: Refreshes all PivotTables and Charts across all sheets.
- **Interactive Dashboard**: Includes charts with slicers for filtering by month, showing total sales over time, and sales by region and product.
- **Macro Button**: One-click execution of the full workflow via a button on the Dashboard sheet.

## How It Works
1. **AddMonthName**: Adds a "Month" column to sheets with " Raw Data" in their name.
2. **CopyToAllData**: Copies new data from raw sheets to the "AllData" sheet if the month isn't already present.
3. **RefreshAll**: Refreshes all PivotTables and Charts across all sheets.
4. **UpdateFile**: Runs all the above macros in sequence.

## How to Use It
1. **Open the Workbook**: Open the `Monthly_Report_Automation_Before.xlsm` file in Excel.
2. **Enable Macros**: Make sure to enable macros when prompted.
3. **Update File**: Go to the `Dashboard` sheet and click the "Update File" button. This will:
   - Add month names to the raw data sheets.
   - Consolidate data into the `AllData` sheet.
   - Refresh all PivotTables and Charts.

## Technologies Used
- **Excel VBA**: For automation and data processing.
- **Excel PivotTables and Charts**: For interactive reporting.
- **Excel Slicers**: For filtering data in the dashboard.

## Customization
- **Modify Macros**: Open the VBA editor (`Alt + F11`) to customize the macros as needed.
- **Enhance Dashboard**: Add more charts, slicers, or conditional formatting to the `Dashboard` sheet.

## Files Included
- `Monthly_Report_Automation_Before.xlsm`: The main Excel file with embedded macros and dashboard.
- `README.md`: This file, providing an overview and instructions.

## Contact
For any questions or customizations, feel free to reach out via Upwork.

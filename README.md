# Undervalued Stock Scanner

### Disclaimer: This project is intended for educational and informational purposes only. The analysis, models, and methodologies used in this project are not intended to serve as financial or investment advice.The information presented should not be construed as a recommendation to buy, sell, or hold any securities or assets. Investing in the stock market carries inherent risks, and any decisions based on the content of this project are solely the responsibility of the individual. Always consult with a qualified financial advisor before making any investment decisions. 

## Project Overview
The Undervalued Stock Scanner project aims to help identify potentially undervalued stocks by analyzing key financial metrics across various sectors. 
The project focuses on three market cap categories (Large Cap, Mid Cap, and Small Cap) and applies statistical methods to assess stock fundamentals, including quartile and mean calculations.

Raw stock data is downloaded from the Thinkorswim scanner in the form of 11 CSV files, each corresponding to a different sector. 
These files are processed and analyzed using Power Query, advanced Excel features and VBA macros resulting in 11 sector-specific workbooks. 
Each workbook contains detailed analysis of stocks based on fundamental criteria, providing insights into potential investment opportunities.

The project also includes a Power BI dashboard that showcases data from all result Excel workbooks in a more interactive and user-friendly way.

## Prerequisites

- Microsoft Excel
- Microsoft Power BI Desktop
- Optional: Thinkorswim app with active Charles Schwab account for data update/refresh.

## Summary of Metrics and Statistical Methods Used

### Stock Fundamentals

The evaluation is based on a series of criteria and statistical methods outlined in the `Conceptual Framework` document. The criteria include various stock fundamental indicators:
- Price-to-Free Cash Flow (P/FCF) ratio
- Price-to-Book (P/B) ratio
- Return on Equity (ROE)
- Return on Assets (ROA)
- Asset-to-equity (A/E) ratio
- Price-to-Earnings (P/E) ratio

### Statistical Methods

- Quartile Calculations: Quartiles (Q1, Q3) are calculated for each fundamental indicator, and outliers are identified based on the upper and lower bounds.
- Mean Calculations: The mean of each fundamental metric is calculated for stocks within each industry and market cap category (Large Cap, Mid Cap, Small Cap). The means are calculated without outliers to ensure that the results reflect the central tendency of the data.  
- Standard Deviation and Z-score Calculations: Z-scores for individual fundamentals are used to compare a stock’s performance against the industry average or specific peers.

**For a detailed explanation of the logic behind the project, please refer to the "Conceptual Framework.pdf" document.**

## Summary of Data Preparation Workflow

The data processing steps were done in a sample workbook called `Builder.xlsm`. This workbook was then used to automatically apply the processing steps to the raw data and generate result workbooks.
1. Power Query: Used to import, clean, and transform the raw CSV data. It is also used to automate the creation of Excel formulas for some columns. 
2. Excel formulas: Used to compute fundamental metrics, identify outliers, identify underavalued stocks and assign scores to these stocks based on their fundamental metrics.
3. RTD function for real-time stock prices: the RTD (Real-Time Data) function is used to retrieve real-time stock prices directly into the workbook. This feature integrates with the Thinkorswim platform to pull the latest stock prices based on the ticker symbols.  
4. Pivot tables: Summarizes and calculates the mean values of each fundamental metric by industry and market cap (Large Cap, Mid Cap, Small Cap). Outliers are filtered out from pivot tables to ensure accurate mean calculations.
5. Conditional formatting: Highlights the best stocks based on their scores and key metrics, making it easier to identify potential undervalued stocks.
6. VBA subs: Automate data refresh and Excel formula evaluation.

The outcome of the analysis is a set of 11 Excel workbooks (.xlsm). Each workbook is named according to the sector (e.g., Technology.xlsm, Healthcare.xlsm), and the data is analyzed for each industry within the sector. The following sheets are included within each workbook:
- Main_L (hidden): Includes stocks and fundamental metrics in the large-cap market category for each industry within the sector.
- Main_M (hidden): Includes stocks and fundamental metrics in the mid-cap market category for each industry within the sector.
- Main_S (hidden): Includes stocks and fundamental metrics in the small-cap market category for each industry within the sector.
- Quartile Tables (hidden): Includes three quartile tables (one for each market cap category within the sector), calculating the first and third quartiles (Q1, Q3), lower and upper bounds, and identifying outliers for each fundamental metric.
- Mean_L: Includes multiple pivot tables with the mean values of each fundamental metric calculated by industry for large-cap stocks, excluding outliers.
- Mean_M: Includes multiple pivot tables with the mean values of each fundamental metric calculated by industry for mid-cap stocks, excluding outliers.
- Mean_S: Includes multiple pivot tables with the mean values of each fundamental metric calculated by industry for small-cap stocks, excluding outliers.
- Result_L: Evaluation results for large-cap stocks based on preliminary and additional criteria.
- Result_M: Evaluation results for mid-cap stocks based on preliminary and additional criteria.
- Result_S: Evaluation results for small-cap stocks based on preliminary and additional criteria.
- Dynamic Path (hidden): Includes the project folder path which is used as a parameter in Power Query to allow for using the workbooks in different local computers.

## Power BI Dashboard
Microsoft Power BI was used to generate an interactive dashboard that presents data from the result Excel workbooks (in `\Main\Results`).

## Project Folders and Files

``Undervalued Stock Scanner`` contains the following folders and files:

```bash
Undervalued Stock Scanner\
│              
├── Downloaded CSV Files\         # Folder containing 11 raw CSV files and 1 sample CSV file
│   
├── Main\                         # Folder containing workbooks, instructions and code documentation
│   ├── Results\                  # Folder containing all result workbooks (.xlsm files)
│   ├── Builder.xlsm              # Workbook to build the sample results
│   ├── Watch List.xlsx           # User's watch list template to track preferred stocks
│   └── Code Documentation\       # Folder containing documents about all the codes used in workbooks
│ 
├── Power BI\
│   ├── Code Documentation\                              # Folder containing documents about all the codes used in .pbix files
│   ├── Data Model View_Result Excel Workbooks.pbix      # Power BI file containing a sample data model for the workbooks in \Main\Results folder
│   ├── Industry Averages Dashboard.pbix                 # Power BI file with data visualizations about industry-level averages
│   ├── Undervalued Stocks Dashboard.pbix                # Power BI file with data visualizations about undervalued stock valuation
│   └── Using Dashboards.pdf                             # A PDF guide that provides step-by-step instructions on how to use the dashboards and explains how to interpret and analyze the visualizations
│
├── Instructions.pdf              # Instructions for using the workbooks and dashboards
├── Conceptual Framework.pdf      # Detailed explanation of the project's logic and methodology
├── Table Schema.xlsx             # Outlines the structure of tables
└── Common Issues.pdf             # Instructions, troubleshooting guides, and other resources to help users resolve issues they may encounter
```

## Usage

**All the workbooks use dynamic folder path, which means the project can be used across different computers as long as the folder structure remains unchanged.**

### General guide:

1. Download the Project Folder:
   - Fork the git repo or download the project folder as ZIP file to your local machine and extract it in a desired location. This will create the `Undervalued Stock Scanner` folder in your local machine.
2. View the Workbooks:
   - If you just want to view workbooks in the project folder without the intention to update, refresh or edit them, you don’t have to follow the next steps.
3. Ensure Up-to-Date Data:  
   - Log in to Thinkorswim and download the latest data. Replace the 11 CSV files in the **"Downloaded CSV Files"** folder with the updated files.
   - Ensure that the folder structure remains the same to avoid breaking dynamic file paths in Power Query.
4. Refresh Data and Interact with Workbooks:  
   - Open any workbook in the ``Results`` folder. These workbooks contain the processed data for each sector and will be pre-populated with fundamental stock analysis results.
   - To update the data in your workbook, go to any of the result sheets (`Result_L`, `Result_M`, `Result_S`), click the "`Refresh All Data`" button at the top left of the sheet.

**For more detailed guidance on how to use the workbooks, please refer to the "Instructions.pdf" in `Undervalued Stock Scanner\Main` folder.**

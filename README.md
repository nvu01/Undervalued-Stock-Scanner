# Undervalued Stock Scanner

### Disclaimer: This project is intended for educational and informational purposes only. The analysis, models, and methodologies used in this project are not intended to serve as financial or investment advice.The information presented should not be construed as a recommendation to buy, sell, or hold any securities or assets. Investing in the stock market carries inherent risks, and any decisions based on the content of this project are solely the responsibility of the individual. Always consult with a qualified financial advisor before making any investment decisions. 

## Project Overview
The Undervalued Stock Scanner project aims to help identify potentially undervalued stocks by analyzing key financial metrics across various sectors. The project focuses on three market cap categories—Large Cap, Mid Cap, and Small Cap—and applies statistical methods to assess stock fundamentals, including quartile and mean calculations.

Raw stock data is downloaded from the Thinkorswim scanner in the form of 11 CSV files, each corresponding to a different sector. These files are processed and analyzed using advanced Excel features, such as Power Query, pivot tables, formulas, and filters, resulting in 11 sector-specific workbooks. Each workbook contains detailed analysis of stocks based on fundamental criteria, providing insights into potential investment opportunities.

## Prerequisites

- Microsoft Excel
- Microsoft Power BI Desktop
- Optional: Thinkorswim app with active Charles Schwab account for data update/refresh.

## Summary of Metrics and Statistical Methods Used

### Stock Fundamentals

The evaluation is based on a series of criteria and statistical methods outlined in the Conceptual Framework document. The criteria include various stock fundamental indicators:
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

**For a detailed explanation of the logic behind the project, please refer to the "Conceptual Framework" document.**

## Summary of Data Preparation Workflow

The data processing steps were done in a sample workbook called `Builder.xlsm`. This workbook was then used to automatically apply the processing steps to the raw data and generate result workbooks.
1. Power Query: Used to import, clean, and transform the raw CSV data. It is also used to automate the creation of Excel formulas for certain calculations. 
2. Excel Formulas: Used to compute fundamental metrics, identify outliers, identify underavalued stocks and assign scores to these stocks based on their fundamental metrics.
3. RTD Function for real-time stock prices: the RTD (Real-Time Data) function is used to retrieve real-time stock prices directly into the workbook. This feature integrates with the Thinkorswim platform to pull the latest stock prices based on the ticker symbols.  
4. Pivot Tables: Summarizes and calculates the mean values of each fundamental metric by industry and market cap (Large Cap, Mid Cap, Small Cap), excluding outliers. 
5. Outlier Filtering: Filters outliers from pivot tables to ensure accurate mean calculations.
6. Conditional Formatting: Highlights the best stocks based on their scores and key metrics, making it easier to identify potential undervalued stocks.

The outcome of the analysis is a set of 11 Excel workbooks (.xlsm). Each workbook is named according to the sector (e.g., Technology.xlsm, Healthcare.xlsm), and the data is analyzed for each industry within the sector. The following sheets are included within each workbook:
- Main_L: Includes stocks and fundamental metrics in the large-cap market category for each industry within the sector.
- Main_M: Includes stocks and fundamental metrics in the mid-cap market category for each industry within the sector.
- Main_S: Includes stocks and fundamental metrics in the small-cap market category for each industry within the sector.
- Quartile Tables: Includes three quartile tables (one for each market cap category within the sector), calculating the first and third quartiles (Q1, Q3), lower and upper bounds, and identifying outliers for each fundamental metric.
- Mean_L: Includes multiple pivot tables with the mean values of each fundamental metric for large-cap stocks, excluding outliers, calculated by industry.
- Mean_M: Includes multiple pivot tables with the mean values of each fundamental metric for mid-cap stocks, excluding outliers, calculated by industry.
- Mean_S: Includes multiple pivot tables with the mean values of each fundamental metric for small-cap stocks, excluding outliers, calculated by industry.
- Result_L: Evaluation results for large-cap stocks based on the calculated metrics for each industry.
- Result_M: Evaluation results for mid-cap stocks based on the calculated metrics for each industry.
- Result_S: Evaluation results for small-cap stocks based on the calculated metrics for each industry.
- Dynamic Path: Includes the project folder path.

## Power BI Dashboards
Microsoft Power BI was used to create data models for result Excel workbooks (in `\Main\Results`) and generate interactive dashboards. The goal is to provide insights and industry comparisons using the data from these workbooks.

## Project Folders and Files

``Undervalued Stock Scanner`` contains the following folders and files:
```bash
Undervalued Stock Scanner\
│              
├── Downloaded CSV Files\           # Folder containing 11 raw CSV files and 1 sample CSV file
│   
├── Main\                           # Folder containing workbooks, instructions and code documentation
│   ├── Instructions.pdf            # Instructions for using the result workbooks
│   ├── Results\                    # Folder containing all result workbooks (.xlsm files)
│   ├── Builder.xlsm                # Workbook to build the sample results
│   ├── Watch List.xlsx             # User's watch list template to track preferred stocks
│   └── Code Documentation\         # Folder containing documents about all the codes used in workbooks
│
├── Conceptual Framework.pdf        # Detailed explanation of the project's logic and methodology
└── Table Design.pdf                # Outlines the structure of tables in CSV files and other workbooks
```

``Power BI`` contains the following folders and files:
```bash
Power BI\
│              
├── Code Documentation\                               # Folder containing documents about all the codes used in .pbix files
├── Data Model View_Result Excel Workbooks.pbix       # Power BI file containing a sample data model for the workbooks in \Main\Results folder
├── Industry Averages Dashboard.pbix                  # Power BI file with data visualizations about industry-level averages
├── Undervalued Stocks Dashboard.pbix                 # Power BI file with data visualizations about undervalued stock valuation
└── Using Dashboards.pdf                              # A PDF guide that provides step-by-step instructions on how to use the dashboards and explains how to interpret and analyze the visualizations
```

## Usage

**All the workbooks use dynamic folder path, which means the project can be used across different computers as long as the folder structure remains unchanged.**

### General guide:

1. Download the Project Folder:
   - Download/fork the project folder ``Undervalued Stock Scanner`` to your local machine and extract it in a desired location. This will create the ``Undervalued Stock Scanner`` folder.
2. View the Workbooks:
   - If you just want to view workbooks in the project folder without the intention to update, refresh or edit them, you don’t have to follow the next steps.
3. Ensure Up-to-Date Data:  
   - Log in to Thinkorswim and download the latest data. Replace the 11 CSV files in the **"Downloaded CSV Files"** folder with the updated files.
   - Ensure that the folder structure remains the same to avoid breaking dynamic file paths in Power Query.
4. Refresh Data and Interact with Workbooks:  
   - Open any workbook in the ``Results`` folder. These workbooks contain the processed data for each sector and will be pre-populated with fundamental stock analysis results.
   - To update the data in your workbook, go to any of the result sheets (`Result_L`, `Result_M`, `Result_S`), click the "`Refresh All Data`" button at the top left of the sheet.

**For more detailed guidance on how to use the workbooks, please refer to the "Instructions.pdf" in `Undervalued Stock Scanner\Main` folder.**

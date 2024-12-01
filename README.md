# Undervalued Stock Scanner
### Disclaimer: This project is intended for educational and informational purposes only. The analysis, models, and methodologies used in this project are not intended to serve as financial or investment advice.The information presented should not be construed as a recommendation to buy, sell, or hold any securities or assets. Investing in the stock market carries inherent risks, and any decisions based on the content of this project are solely the responsibility of the individual. Always consult with a qualified financial advisor before making any investment decisions. 

## Project Overview
The Undervalued Stock Scanner project aims to help identify potentially undervalued stocks by analyzing key financial metrics across various sectors. The project focuses on three market cap categories—Large Cap, Mid Cap, and Small Cap—and applies statistical methods to assess stock fundamentals, including quartile and mean calculations.

Raw stock data is downloaded from the Thinkorswim scanner in the form of 11 CSV files, each corresponding to a different sector. These files are processed and analyzed using advanced Excel features, such as Power Query, pivot tables, formulas, and filters, resulting in 11 sector-specific workbooks. Each workbook contains detailed analysis of stocks based on fundamental criteria, providing insights into potential investment opportunities.

## Evaluation Criteria & Methodology
### Stock fundamentals
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

**For a detailed explanation of the logic behind the project, please refer to the "Conceptual Framework" document.**

## Summary of Excel Workflow
The data is processed and analyzed using a combination of Excel features:
1. Power Query: Used to import, clean, and transform the raw CSV data. It is also used to automate the creation of Excel formulas for certain calculations. 
2. Excel Formulas: Used to compute fundamental metrics, identify outliers, identify underavalued stocks and assign scores to these stocks based on their fundamental metrics.
3. RTD Function for Real-Time Stock Prices: the RTD (Real-Time Data) function is used to retrieve real-time stock prices directly into the workbook. This feature integrates with the Thinkorswim platform to pull the latest stock prices based on the ticker symbols.
4. Pivot Tables: Summarizes and calculates the mean values of each fundamental metric by industry and market cap (Large Cap, Mid Cap, Small Cap), excluding outliers. 
5. Outlier Filtering: Filters outliers from pivot tables to ensure accurate mean calculations.
6. Conditional Formatting: Highlights the best stocks based on their scores and key metrics, making it easier to identify potential undervalued stocks.

The outcome of the analysis is a set of 11 Excel workbooks. Each workbook is named according to the sector (e.g., Technology.xlsx, Healthcare.xlsx), and the data is analyzed for each industry within the sector. The following sheets are included within each workbook:
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

## Project Versions, Folders and Files
### Project Versions
There are 3 versions of this project: Version 1 (V1), Version 2 (V2), and Version 3 (V3).
- Version 1 (V1): Available for testing and exploration purposes.
- Version 2 (V2): Completed and recommended for users.
- Version 3 (V3): Completed and recommended for users.

### Folders and Files

```bash
Undervalued_Stock_Scanner/
│              
├── Downloaded CSV Files/             # Folder containing raw CSV data files (11 CSV files)
│
├── V1/                               # Version 1 - Testing and exploration
│   ├── Builder V1.xlsx               # Workbook to build the sample results for Version 1
│   └──  Power Query Codes/            # Folder containing .md files with Power Query codes for V1
│   
├── V2/                               # Version 2 - Completed and recommended
│   ├── Builder V2.xlsx               # Workbook to build the sample results for Version 2
│   ├── Watch List.xlsx               # User watch list to track preferred stocks
│   ├── Results/                      # Folder containing result workbooks for Version 2
│   └── Power Query Codes/            # Folder containing .md files with Power Query codes for V2
│
├── V3/                               # Version 3 - Completed and recommended
│   ├── Builder V3.xlsx               # Workbook to build the sample results for Version 3
│   ├── Watch List.xlsx               # User watch list to track preferred stocks
│   ├── Results/                      # Folder containing result workbooks for Version 3
│   └── Power Query Codes/            # Folder containing .md files with Power Query codes for V3
│
├── Conceptual Framework.pdf         # Detailed explanation of the project's logic and methodology
├── Instructions.pdf                 # Instructions for using the result workbooks
├── Table Design.pdf                 # Outlines the structure of tables in CSV files and other workbooks
├── Excel Formulas/                   # Folder containing all Excel formulas used in this project
│
└── (Other project files)
```
## Prerequisites
- Microsoft Excel
- Thinkorswim app with active Charles Schwab account

## Usage
**The project folder uses dynamic file paths in Power Query queries, which means the project can be used across different computers as long as the folder structure remains unchanged.**  

### General guide:
1. Download the Project Folder:
   - Download the project zip file "Undervalued Stock Scanner.zip" to your local machine and extract it to a desired location. This will create the "Undervalued Stock Scanner" folder.
2. Ensure Up-to-Date Data:  
   - Log in to Thinkorswim and download the latest data. Replace the 11 CSV files in the **"Downloaded CSV Files"** folder with the updated files.
   - Ensure that the folder structure remains the same to avoid breaking dynamic file paths in Power Query.
3. Refresh Data and Interact with Workbooks:  
   - Open any workbook from the **"Results"** folder (located in either the **V2** or **V3** folder). These workbooks contain the processed data for each sector and will be pre-populated with fundamental stock analysis results.
   - To update the data in your workbook, you may need to refresh the queries or follow additional steps detailed in the "Instructions" document.

**For more detailed guidance on how to use the workbooks, refer to the "Instructions.pdf" in the project folder.**
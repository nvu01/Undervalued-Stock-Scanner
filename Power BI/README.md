# Power BI Data Modeling and Visualization

This part of Undervalued Stock Scanner project uses Microsoft Power BI to create data models for result Excel workbooks (in `Undervalued Stock Scanner\Main\Results`) and generate interactive dashboards. The goal is to provide insights and industry comparisons using the data from these workbooks.

## Requirements

- Microsoft Power BI Desktop installed on your computer.
- The "Undervalued Stock Scanner" project folder containing Excel workbooks as datasource.

## Important Files

The Power BI folder contains the following key files:

- `Undervalued Stocks Dashboard.pbix`: Power BI file with data visualizations for undervalued stock evaluation.
- `Using Dashboard.pdf`: A PDF guide that provides step-by-step instructions on how to use the dashboards and explains how to interpret and analyze the visualizations.
- `Data Model View_Result Workbooks.pbix`: Power BI file containing a data model view for the Excel workbooks in `\Main\Results` folder. This model organizes the tables and establishes relationships for better analysis.

## How to Use the .pbix Files

For any `.pbix` file, do the following steps when opening the file for the first time in a local machine:
-	Copy the path of the project folder that contains the **Main** folder. (Ex: D:\PROJECTS\Undervalued Stock Scanner)
-	Open a `.pbix` file with Power BI
-	Home tab --> Click `Transform data` drop-down --> `Edit parameters` --> Click on the `FolderPath` parameter --> Paste the current folder path to replace the default folder path --> Click `OK` --> Click `Apply changes` in the popup at the top

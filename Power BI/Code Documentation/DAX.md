## Undervalued Stocks Dashboard

### 1. Absolute Value of Z-Score
These measures in `All_Results` are used to create gradient conditional format in the heatmap matrix.

#### Absolute value of A/E z-score
```
Abs_ZS_AE = ABS(MAX(All_Results[A/E Z-Score]))
```

#### Absolute value of P/B z-score
```
Abs_ZS_PB = ABS(MAX(All_Results[P/B Z-Score]))
```

#### Absolute value of P/E z-score
```
Abs_ZS_PE = ABS(MAX(All_Results[P/E Z-Score]))
```

#### Absolute value of P/FCF z-score
```
Abs_ZS_PFCF = ABS(MAX(All_Results[P/FCF Z-Score]))
```

#### Absolute value of ROA z-score
```
Abs_ZS_ROA = ABS(MAX(All_Results[ROA Z-Score]))
```

#### Absolute value of ROE z-score
```
Abs_ZS_ROE = ABS(MAX(All_Results[ROE Z-Score]))
```

### 2. Ranking Z-Scores
These measures in `All_Results` are used as filter on table visuals to sort and filter the z-scores.

#### Ranking A/E z-score
A/E z-score should be negative and further away from 0
```
Rank_AE_ZS = RANKX(ALLSELECTED(All_Results), [A/E Z-Score], MAX([A/E Z-Score]), ASC)
```

#### Ranking P/B z-score
P/B should be negative and further away from 0
```
Rank_PB_ZS = RANKX(ALLSELECTED(All_Results), [P/B Z-Score], MAX([P/B Z-Score]),ASC)
```

#### Ranking P/E z-score
P/E z-score should be negative and further away from 0
```
Rank_PE_ZS = RANKX(ALLSELECTED(All_Results), [P/E Z-Score], MAX([P/E Z-Score]), ASC)
```

#### Ranking P/FCF z-score
P/FCF z-score should be negative and further away from 0
```
Rank_PFCF_ZS = RANKX(ALLSELECTED(All_Results), [P/FCF Z-Score], MAX([P/FCF Z-Score]),ASC)
```

#### Ranking ROA z-score
ROA z-score should be positive and further away from 0
```
Rank_ROA_ZS = RANKX(ALLSELECTED(All_Results), [ROA Z-Score], MAX([ROA Z-Score]), DESC)
```

#### Ranking ROE z-score
ROE z-score should be positive and further away from 0
```
Rank_ROE_ZS = RANKX(ALLSELECTED(All_Results), [ROE Z-Score], MAX([ROE Z-Score]), DESC)
```

### 3. Criterion Columns
5 new Boolean columns in `All_Results` are used in the horizontal stacked bar chart to indicate whether a stock meet the additional criteria defined in the framework.

#### Criterion 1
```
Criterion 1 = VAR Industry = All_Results[Industry]
VAR MarketCap = All_Results[Market Cap]
VAR MeanValue = CALCULATE(MAX(All_Means[Mean_PB]), FILTER(All_Means, All_Means[Industry] = Industry && All_Means[Market Cap] = MarketCap))
RETURN
    IF(All_Results[P/B] < 0.7 * MeanValue, 1, 0)
```

#### Criterion 2
```
Criterion 2 = VAR Industry = All_Results[Industry]
VAR MarketCap = All_Results[Market Cap]
VAR MeanValue = CALCULATE(MAX(All_Means[Mean_ROE]), FILTER(All_Means, All_Means[Industry] = Industry && All_Means[Market Cap] = MarketCap))
RETURN
    IF(All_Results[ROE] > MeanValue, 1, 0)
```

#### Criterion 3
```
Criterion 3 = VAR Industry = All_Results[Industry]
VAR MarketCap = All_Results[Market Cap]
VAR MeanValue = CALCULATE(MAX(All_Means[Mean_ROA]), FILTER(All_Means, All_Means[Industry] = Industry && All_Means[Market Cap] = MarketCap))
RETURN
    IF(All_Results[ROA] > MeanValue, 1, 0)
```

#### Criterion 4
```
Criterion 4 = VAR Industry = All_Results[Industry]
VAR MarketCap = All_Results[Market Cap]
VAR MeanValue = CALCULATE(MAX(All_Means[Mean_PE]), FILTER(All_Means, All_Means[Industry] = Industry && All_Means[Market Cap] = MarketCap))
RETURN
    IF(All_Results[P/E] < MeanValue, 1, 0)
```

#### Criterion 5
```
Criterion 5 = IF(All_Results[P/E] > 1 && All_Results[P/E] < 25, 1, 0)
```

### 4. Missing Score Column
This new column in `All_Results` is used in the horizontal stacked bar chart to indicate the total score each stock misses.
```
missing_score = 5 - [SCORE]
```

### 5. Dynamic Title

#### Home page
This measure in `All_Results` returns the dynamic page title containing the sector and market cap selected in the slicers.
```
DynamicTitle = 
VAR SectorVal = VALUES(All_Results[Sector])
VAR SectorCount = COUNTROWS(SectorVal)

VAR MarketCapVal = VALUES(All_Results[Market Cap])
VAR MarketCapCount = COUNTROWS(MarketCapVal)

VAR SectorTitle = SWITCH(
        TRUE(),
        SectorCount = 1, SELECTEDVALUE(All_Results[Sector]),
        SectorCount >= 2 && SectorCount < 11, CONCATENATEX(SectorVal, All_Results[Sector], ", "),
        "All Sectors"
    )

VAR MarketCapTitle =
    SWITCH(
        TRUE(),
        MarketCapCount = 1, SELECTEDVALUE(All_Results[Market Cap]) & " Market Cap",
        MarketCapCount = 2, CONCATENATEX(MarketCapVal, All_Results[Market Cap], " & ") & " Caps",
        "All Market Caps"
    )

RETURN
"Undervalued Stocks in " &SectorTitle &" - " &MarketCapTitle
```

#### Industry Averages page
This measure in `All_Means` returns the dynamic page title containing the sector and market cap selected in the slicers.
```
DynamicTitle2 = "Industry Averages for " & SELECTEDVALUE(All_Means[Sector]) & " - " & SELECTEDVALUE(All_Means[Market Cap]) & " Market Cap"
```

### 6. Last Update Date and Time
This measure in `All_Results` returns the date and time of the last data update in result workbooks.
```
LastUpdate = "Last Data Update: " & MAX(All_Results[Date modified])
```
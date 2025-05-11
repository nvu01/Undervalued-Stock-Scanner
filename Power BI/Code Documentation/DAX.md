## Undervalued Stocks Dashboard

### 1. Absolute Value of Z-Score
These measures in `Result_All` are used to create gradient conditional format in the heatmap matrix.

#### Absolute value of A/E z-score
```
Abs_ZS_AE = ABS(MAX('Result_All'[A/E Z-Score]))
```

#### Absolute value of P/B z-score
```
Abs_ZS_PB = ABS(MAX('Result_All'[P/B Z-Score]))
```

#### Absolute value of P/E z-score
```
Abs_ZS_PE = ABS(MAX('Result_All'[P/E Z-Score]))
```

#### Absolute value of P/FCF z-score
```
Abs_ZS_PFCF = ABS(MAX('Result_All'[P/FCF Z-Score]))
```

#### Absolute value of ROA z-score
```
Abs_ZS_ROA = ABS(MAX('Result_All'[ROA Z-Score]))
```

#### Absolute value of ROE z-score
```
Abs_ZS_ROE = ABS(MAX('Result_All'[ROE Z-Score]))
```

### 2. Ranking Z-Scores
These measures in `Result_All` are used as filter on table visuals to sort and filter the z-scores.

#### Ranking A/E z-score
A/E z-score should be negative and further away from 0
```
Rank_AE_ZS = RANKX(ALLSELECTED('Result_All'), [A/E Z-Score], MAX([A/E Z-Score]), ASC)
```

#### Ranking P/B z-score
P/B should be negative and further away from 0
```
Rank_PB_ZS = RANKX(ALLSELECTED('Result_All'), [P/B Z-Score], MAX([P/B Z-Score]),ASC)
```

#### Ranking P/E z-score
P/E z-score should be negative and further away from 0
```
Rank_PE_ZS = RANKX(ALLSELECTED('Result_All'), [P/E Z-Score], MAX([P/E Z-Score]), ASC)
```

#### Ranking P/FCF z-score
P/FCF z-score should be negative and further away from 0
```
Rank_PFCF_ZS = RANKX(ALLSELECTED('Result_All'), [P/FCF Z-Score], MAX([P/FCF Z-Score]),ASC)
```

#### Ranking ROA z-score
ROA z-score should be positive and further away from 0
```
Rank_ROA_ZS = RANKX(ALLSELECTED('Result_All'), [ROA Z-Score], MAX([ROA Z-Score]), DESC)
```

#### Ranking ROE z-score
ROE z-score should be positive and further away from 0
```
Rank_ROE_ZS = RANKX(ALLSELECTED('Result_All'), [ROE Z-Score], MAX([ROE Z-Score]), DESC)
```

### 3. Score Missed
This new column in `Result_All` is used in the horizontal stacked bar chart to indicate the total score each stock misses.
```
missing_score = 5 - [SCORE]
```
### 4. Last Update Date and Time
This measure in `Result_All` returns the date and time of the last data update in result workbooks.
```
LastUpdate = "Last Data Update: " & MAX('Result_All'[Date modified])
```

### 5. Dynamic Title
This measure in `Parameters` returns the dynamic dashboard title containing the name of sector corresponding to the current source workbook.
```
DynamicTitle = "Undervalued Stocks in " & SELECTEDVALUE('Result_All'[Sector])
```

## Industry Averages Dashboards

### 1. Dynamic Title
This measure in `Parameters` returns the dynamic dashboard title containing the name of sector corresponding to the current source workbook.
```
DynamicTitle = "Industry Averages for " & SELECTEDVALUE('Combined Mean Table'[Sector]) & " - " & SELECTEDVALUE('Combined Mean Table'[Market Cap]) & " Market Cap"```

### 2. Last Update Date and Time
This measure in `Last Update` returns the date and time of the last data update in result workbooks.
```
LastUpdate = "Last Data Update: " & MAX('Combined Mean Table'[Date modified])
```
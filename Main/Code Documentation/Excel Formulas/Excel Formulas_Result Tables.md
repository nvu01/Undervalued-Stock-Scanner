### Pull Data From Main Tables Based on Preliminary Criteria (A9):

#### Result_L:
```
=FILTER(FILTER(Main_L, Main_L[Preliminary Criteria]="YES"), COUNTIF($A$6:$S$6,Main_L[#Headers]))
```
#### Result_M:
```
=FILTER(FILTER(Main_M, Main_M[Preliminary Criteria]="YES"), COUNTIF($A$6:$S$6,Main_M[#Headers]))
```
#### Result_S:
```
=FILTER(FILTER(Main_S, Main_S[Preliminary Criteria]="YES"), COUNTIF($A$6:$S$6,Main_S[#Headers]))
```

### Score Columns in Result Tables (T9):

#### Result_L:
```
=SUM(IF($I9<(0.7*VLOOKUP($A9, Mean_PB_L, 2, FALSE)),1,0), IF($J9>VLOOKUP($A9, Mean_ROE_L, 2, FALSE),1,0), IF($K9>VLOOKUP($A9, Mean_ROA_L, 2,FALSE),1,0), IF($M9<VLOOKUP($A9, Mean_PE_L, 2, FALSE),1,0), IF(AND($M9>1,$M9<25),1,0))
```
#### Result_M:
```
=SUM(IF($I9<(0.7*VLOOKUP($A9, Mean_PB_L, 2, FALSE)),1,0), IF($J9>VLOOKUP($A9, Mean_ROE_M, 2, FALSE),1,0), IF($K9>VLOOKUP($A9, Mean_ROA_M, 2,FALSE),1,0), IF($M9<VLOOKUP($A9, Mean_PE_M, 2, FALSE),1,0), IF(AND($M9>1,$M9<25),1,0))
```
#### Result_S:
```
=SUM(IF($I9<(0.7*VLOOKUP($A9, Mean_PB_L, 2, FALSE)),1,0), IF($J9>VLOOKUP($A9, Mean_ROE_S, 2, FALSE),1,0), IF($K9>VLOOKUP($A9, Mean_ROA_S, 2,FALSE),1,0), IF($M9<VLOOKUP($A9, Mean_PE_S, 2, FALSE),1,0), IF(AND($M9>1,$M9<25),1,0))
```

### Summary Cells:

#### Number of stocks (T6):
```
="Number of stocks: "&COUNTA($B$9:$B$1049576)
```
#### Max Score (T2):
```
="Max Score:"&IFERROR(CHAR(10)&MAX($T$9:$T$1048576)&"/5", CHAR(10)&"0/5")
```
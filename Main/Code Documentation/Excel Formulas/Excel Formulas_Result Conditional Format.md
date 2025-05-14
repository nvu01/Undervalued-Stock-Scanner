### RESULT_L CONDITIONAL FORMAT

#### P/B
```
=$I8 < (0.8*VLOOKUP($A8, Mean_PB_L, 2, FALSE))
```
#### ROE
```
=$J8 > VLOOKUP($A8, Mean_ROE_L, 2, FALSE)
```
#### ROA
```
=$K8 > VLOOKUP($A8, Mean_ROA_L, 2, FALSE)
```
#### P/E
```
=AND($M8>1, $J8<25, $M8 < VLOOKUP($A8, Mean_PE_L, 2, FALSE))
=AND($M8>1, $M8<25)
=$M8 < VLOOKUP($A8, Mean_PE_L, 2, FALSE)
```

### RESULT_M CONDITIONAL FORMAT

#### P/B
```
=$I8 < (0.7*VLOOKUP($A8, Mean_PB_M, 2, FALSE))
=$I8 < 1
```
#### ROE
```
=$J8 > VLOOKUP($A8, Mean_ROE_M, 2, FALSE)
```
#### ROA
```
=$K8 > VLOOKUP($A8, Mean_ROA_M, 2, FALSE)
```
#### P/E
```
=AND($M8>1, $M8<25, $M8 < VLOOKUP($A8, Mean_PE_M, 2, FALSE))
=AND($M8>1, $M8<25)
=$M8 < VLOOKUP($A8, Mean_PE_M, 2, FALSE)
```

### RESULT_S CONDITIONAL FORMAT

#### P/B
```
=$I8 < (0.7*VLOOKUP($A8, Mean_PB_S, 2, FALSE))
=$I8 < 1
```
#### ROE
```
=$J8 > VLOOKUP($A8, Mean_ROE_S, 2, FALSE)
```
#### ROA
```
=$K8 > VLOOKUP($A8, Mean_ROA_S, 2, FALSE)
```
#### P/E
```
=AND($M8>1, $M8<25, $M8 < VLOOKUP($M8, Mean_PE_S, 2, FALSE))
=AND($M8>1, $M8<25)
=$M8 < VLOOKUP($A8, Mean_PE_S, 2, FALSE)

```

### Z SCORE COLUMNS CONDITIONAL FORMAT

#### P/FCF Z Score
```
=$N8<0
```
#### P/B Z Score
```
=$O8<0
```
#### ROE Z Score
```
=$P8>0
```
#### ROA Z Score
```
=$Q8>0
```
#### A/E Z Score
```
=$R8<0
```
#### P/E Z Score
```
=$S8<0


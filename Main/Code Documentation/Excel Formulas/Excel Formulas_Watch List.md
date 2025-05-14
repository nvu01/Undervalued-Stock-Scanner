#### Sector Column:
```
=IF([@Symbol]="", "", XLOOKUP([@Symbol], All_Results[Symbol], All_Results[Sector]))
```
#### Industry Column:
```
=XLOOKUP([@Symbol], All_Results[Symbol], All_Results[Industry])
```
#### Market Cap Column:
```
=XLOOKUP([@Symbol], All_Results[Symbol], All_Results[Market Cap (M)])
```
#### Current Price Column:
```
=RTD("tos.rtd", , "LAST", [@Symbol])
```
#### Free CF Column:
```
=XLOOKUP([@Symbol], All_Results[Symbol], All_Results[Free CF])
```
#### BVPS Column:
```
=XLOOKUP([@Symbol], All_Results[Symbol], All_Results[BVPS])
```
#### EPS Column:
```
=XLOOKUP([@Symbol], All_Results[Symbol], All_Results[EPS])
```
#### P/FCF Column:
```
=[@[Current Price]]/[@[Free CF]]
```
#### P/B Column:
```
=[@[Current Price]]/[@BVPS]
```
#### ROE Column:
```
=XLOOKUP([@Symbol], All_Results[Symbol], All_Results[ROE])
```
#### ROA Column:
```
=XLOOKUP([@Symbol], All_Results[Symbol], All_Results[ROA])
```
#### A/E Column:
```
=XLOOKUP([@Symbol], All_Results[Symbol], All_Results[A/E])
```
#### P/E Column:
```
=[@[Current Price]]/[@EPS]
```
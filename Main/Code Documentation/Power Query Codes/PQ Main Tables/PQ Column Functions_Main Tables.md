These column functions returns Excel formulas. Check the files in **Excel Formulas** folder to see specific formulas used for columns in any worksheet.

### Current Price column function in main tables:
```m
let
    Current_Price = () => ("=RTD(""tos.rtd"", , ""LAST"", [@Symbol])")
in
    Current_Price
```
### P/FCF column funtcion in main tables:
```m
let
    PFCF = () => ("=IF(ISNUMBER([@[Free CF]]),IF([@[Free CF]]<>0,[@[Current Price]]/[@[Free CF]],""""),"""")")
in
    PFCF 
 ```
### P/B column function in main tables:
```my
let
    PB = () => ("=IF(ISNUMBER([@BVPS]),IF([@BVPS]<>0,[@[Current Price]]/[@BVPS],""""),"""")")
in
    PB
```
### P/E column function in main tables:
```m
let
    PE = () => ("=IF(ISNUMBER([@EPS]),IF([@EPS]<>0,[@[Current Price]]/[@EPS],""""),"""")")
in
    PE
```



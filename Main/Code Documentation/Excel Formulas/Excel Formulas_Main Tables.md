### Current Price Columns in All Main Tables:
```
=RTD("tos.rtd", , "LAST", [@Symbol])
```
### P/FCF Columns in All Main Tables:
```
=IF(ISNUMBER([@[Free CF]]),IF([@[Free CF]]<>0,[@[Current Price]]/[@[Free CF]],""),"") 
```
### P/B Columns in All Main Tables:
```
=IF(ISNUMBER([@BVPS]),IF([@BVPS]<>0,[@[Current Price]]/[@BVPS],""),"") 
```
### P/E Columns in All Main Tables:
```
=IF(ISNUMBER([@EPS]),IF([@EPS]<>0,[@[Current Price]]/[@EPS],""),"")
```

### Preliminary Criteria Columns:

#### Preliminary Criteria column in Main_L:
```
=IF(AND([@[P/FCF]]>0, [@[P/FCF]]<VLOOKUP([@Industry], Mean_PFCF_L, 2, FALSE), [@[P/B]]>0,  [@[P/B]]<VLOOKUP([@Industry], Mean_PB_L, 2, FALSE), [@ROE]>10, [@ROA]>5, [@[A/E]]>1, [@[A/E]]<VLOOKUP([@Industry], Mean_AE_L, 2, FALSE)), "YES", "NO")
```
#### Preliminary Criteria column in Main_M:
```
=IF(AND([@[P/FCF]]>0, [@[P/FCF]]<VLOOKUP([@Industry], Mean_PFCF_M, 2, FALSE), [@[P/B]]>0,  [@[P/B]]<VLOOKUP([@Industry], Mean_PB_M, 2, FALSE), [@ROE]>10, [@ROA]>5, [@[A/E]]>1, [@[A/E]]<VLOOKUP([@Industry], Mean_AE_M, 2, FALSE)), "YES", "NO")
```
#### Preliminary Criteria column in Main_S:
```
=IF(AND([@[P/FCF]]>0, [@[P/FCF]]<VLOOKUP([@Industry], Mean_PFCF_S, 2, FALSE), [@[P/B]]>0,  [@[P/B]]<VLOOKUP([@Industry], Mean_PB_S, 2, FALSE), [@ROE]>10, [@ROA]>5, [@[A/E]]>1, [@[A/E]]<VLOOKUP([@Industry], Mean_AE_S, 2, FALSE)), "YES", "NO")
```

### Z Score Columns in Main_L:

#### P/FCF Z Score:
```
=([@[P/FCF]]-VLOOKUP([@Industry], Mean_PFCF_L, 2, FALSE))/(VLOOKUP([@Industry], Mean_PFCF_L, 3, FALSE))
```
#### P/B Z Score:
```
=([@[P/B]]-VLOOKUP([@Industry], Mean_PB_L, 2, FALSE))/(VLOOKUP([@Industry], Mean_PB_L, 3, FALSE))
```
#### ROE Z Score:
```
=([@ROE]-VLOOKUP([@Industry], Mean_ROE_L, 2, FALSE))/(VLOOKUP([@Industry], Mean_ROE_L, 3, FALSE))
```
#### ROA Z Score:
```
=([@ROA]-VLOOKUP([@Industry], Mean_ROA_L, 2, FALSE))/(VLOOKUP([@Industry], Mean_ROA_L, 3, FALSE))
```
#### A/E Z Score:
```
=([@[A/E]]-VLOOKUP([@Industry], Mean_AE_L, 2, FALSE))/(VLOOKUP([@Industry], Mean_AE_L, 3, FALSE))
```
#### P/E Z Score:
```
=([@[P/E]]-VLOOKUP([@Industry], Mean_PE_L, 2, FALSE))/(VLOOKUP([@Industry], Mean_PE_L, 3, FALSE))
```

### Z Score Columns in Main_M:

#### P/FCF Z Score:
```
=([@[P/FCF]]-VLOOKUP([@Industry], Mean_PFCF_M, 2, FALSE))/(VLOOKUP([@Industry], Mean_PFCF_M, 3, FALSE))
```
#### P/B Z Score:
```
=([@[P/B]]-VLOOKUP([@Industry], Mean_PB_M, 2, FALSE))/(VLOOKUP([@Industry], Mean_PB_M, 3, FALSE))
```
#### ROE Z Score:
```
=([@ROE]-VLOOKUP([@Industry], Mean_ROE_M, 2, FALSE))/(VLOOKUP([@Industry], Mean_ROE_M, 3, FALSE))
```
#### ROA Z Score:
```
=([@ROA]-VLOOKUP([@Industry], Mean_ROA_M, 2, FALSE))/(VLOOKUP([@Industry], Mean_ROA_M, 3, FALSE))
```
#### A/E Z Score:
```
=([@[A/E]]-VLOOKUP([@Industry], Mean_AE_M, 2, FALSE))/(VLOOKUP([@Industry], Mean_AE_M, 3, FALSE))
```
#### P/E Z Score:
```
=([@[P/E]]-VLOOKUP([@Industry], Mean_PE_M, 2, FALSE))/(VLOOKUP([@Industry], Mean_PE_M, 3, FALSE))
```

### Z Score Columns in Main_S:

#### P/FCF Z Score:
```
=([@[P/FCF]]-VLOOKUP([@Industry], Mean_PFCF_S, 2, FALSE))/(VLOOKUP([@Industry], Mean_PFCF_S, 3, FALSE))
```
#### P/B Z Score:
```
=([@[P/B]]-VLOOKUP([@Industry], Mean_PB_S, 2, FALSE))/(VLOOKUP([@Industry], Mean_PB_S, 3, FALSE))
```
#### ROE Z Score:
```
=([@ROE]-VLOOKUP([@Industry], Mean_ROE_S, 2, FALSE))/(VLOOKUP([@Industry], Mean_ROE_S, 3, FALSE))
```
#### ROA Z Score:
```
=([@ROA]-VLOOKUP([@Industry], Mean_ROA_S, 2, FALSE))/(VLOOKUP([@Industry], Mean_ROA_S, 3, FALSE))
```
#### A/E Z Score:
```
=([@[A/E]]-VLOOKUP([@Industry], Mean_AE_S, 2, FALSE))/(VLOOKUP([@Industry], Mean_AE_S, 3, FALSE))
```
#### P/E Z Score:
```
=([@[P/E]]-VLOOKUP([@Industry], Mean_PE_S, 2, FALSE))/(VLOOKUP([@Industry], Mean_PE_S, 3, FALSE))
```

### Outlier Columns in Main_L:

#### P/FCF Outlier:
```
=IF(OR([@[P/FCF]]<XLOOKUP([@Industry], Quartile_L[Industry], Quartile_L[P/FCF LB]), [@[P/FCF]]>XLOOKUP([@Industry], Quartile_L[Industry], Quartile_L[P/FCF UB])), "Yes", "No")
```
#### P/B Outlier:
```
=IF(OR([@[P/B]]<XLOOKUP([@Industry],Quartile_L[Industry],Quartile_L[P/B LB]),[@[P/B]]>XLOOKUP([@Industry],Quartile_L[Industry],Quartile_L[P/B UB])),"Yes","No")
```
#### ROE Outlier:
```
=IF(OR([@ROE]<XLOOKUP([@Industry],Quartile_L[Industry],Quartile_L[ROE LB]),[@ROE]>XLOOKUP([@Industry],Quartile_L[Industry],Quartile_L[ROE UB])),"Yes","No")
```
#### ROA Outlier:
```
=IF(OR([@ROA]<XLOOKUP([@Industry],Quartile_L[Industry],Quartile_L[ROA LB]),[@ROA]>XLOOKUP([@Industry],Quartile_L[Industry],Quartile_L[ROA UB])),"Yes","No")
```
#### A/E Outlier:
```
=IF(OR([@[A/E]]<XLOOKUP([@Industry],Quartile_L[Industry],Quartile_L[A/E LB]),[@[A/E]]>XLOOKUP([@Industry],Quartile_L[Industry],Quartile_L[A/E UB])),"Yes","No")
```
#### P/E Outlier:
```
=IF(OR([@[P/E]]<XLOOKUP([@Industry],Quartile_L[Industry],Quartile_L[P/E LB]),[@[P/E]]>XLOOKUP([@Industry],Quartile_L[Industry],Quartile_L[P/E UB])),"Yes","No")
```

### Outlier Columns in Main_M:

#### P/FCF Outlier:
```
=IF(OR([@[P/FCF]]<XLOOKUP([@Industry], Quartile_M[Industry], Quartile_M[P/FCF LB]), [@[P/FCF]]>XLOOKUP([@Industry], Quartile_M[Industry], Quartile_M[P/FCF UB])), "Yes", "No")
```
#### P/B Outlier:
```
=IF(OR([@[P/B]]<XLOOKUP([@Industry],Quartile_M[Industry],Quartile_M[P/B LB]),[@[P/B]]>XLOOKUP([@Industry],Quartile_M[Industry],Quartile_M[P/B UB])),"Yes","No")
```
#### ROE Outlier:
```
=IF(OR([@ROE]<XLOOKUP([@Industry],Quartile_M[Industry],Quartile_M[ROE LB]),[@ROE]>XLOOKUP([@Industry],Quartile_M[Industry],Quartile_M[ROE UB])),"Yes","No")
```
#### ROA Outlier:
```
=IF(OR([@ROA]<XLOOKUP([@Industry],Quartile_M[Industry],Quartile_M[ROA LB]),[@ROA]>XLOOKUP([@Industry],Quartile_M[Industry],Quartile_M[ROA UB])),"Yes","No")
```
#### A/E Outlier:
```
=IF(OR([@[A/E]]<XLOOKUP([@Industry],Quartile_M[Industry],Quartile_M[A/E LB]),[@[A/E]]>XLOOKUP([@Industry],Quartile_M[Industry],Quartile_M[A/E UB])),"Yes","No")
```
#### P/E Outlier:
```
=IF(OR([@[P/E]]<XLOOKUP([@Industry],Quartile_M[Industry],Quartile_M[P/E LB]),[@[P/E]]>XLOOKUP([@Industry],Quartile_M[Industry],Quartile_M[P/E UB])),"Yes","No")
```

### Outlier Columns in Main_S:

#### P/FCF Outlier:
```
=IF(OR([@[P/FCF]]<XLOOKUP([@Industry], Quartile_S[Industry], Quartile_S[P/FCF LB]), [@[P/FCF]]>XLOOKUP([@Industry], Quartile_S[Industry], Quartile_S[P/FCF UB])), "Yes", "No")
```
#### P/B Outlier:
```
=IF(OR([@[P/B]]<XLOOKUP([@Industry],Quartile_S[Industry],Quartile_S[P/B LB]),[@[P/B]]>XLOOKUP([@Industry],Quartile_S[Industry],Quartile_S[P/B UB])),"Yes","No")
```
#### ROE Outlier:
```
=IF(OR([@ROE]<XLOOKUP([@Industry],Quartile_S[Industry],Quartile_S[ROE LB]),[@ROE]>XLOOKUP([@Industry],Quartile_S[Industry],Quartile_S[ROE UB])),"Yes","No")
```
#### ROA Outlier:
```
=IF(OR([@ROA]<XLOOKUP([@Industry],Quartile_S[Industry],Quartile_S[ROA LB]),[@ROA]>XLOOKUP([@Industry],Quartile_S[Industry],Quartile_S[ROA UB])),"Yes","No")
```
#### A/E Outlier:
```
=IF(OR([@[A/E]]<XLOOKUP([@Industry],Quartile_S[Industry],Quartile_S[A/E LB]),[@[A/E]]>XLOOKUP([@Industry],Quartile_S[Industry],Quartile_S[A/E UB])),"Yes","No")
```
#### P/E Outlier:
```
=IF(OR([@[P/E]]<XLOOKUP([@Industry],Quartile_S[Industry],Quartile_S[P/E LB]),[@[P/E]]>XLOOKUP([@Industry],Quartile_S[Industry],Quartile_S[P/E UB])),"Yes","No")
```




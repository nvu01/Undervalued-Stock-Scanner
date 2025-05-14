### Q1, Q3 Columns in Quartile_L:

#### P/FCF Q1, P/FCF Q3:
```
=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[P/FCF]),1)
=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[P/FCF]),3)
```
#### P/B Q1, P/B Q3:
```
=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[P/B]),1)
=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[P/B]),3)
```
#### ROE Q1, ROE Q3:
```
=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[ROE]),1)
=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[ROE]),3)
```
#### ROA Q1, ROA Q3:
```
=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[ROA]),1)
=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[ROA]),3)
```
#### A/E Q1, A/E Q3:
```
=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[A/E]),1)
=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[A/E]),3)
```
#### P/E Q1, P/E Q3:
```
=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[P/E]),1)
=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[P/E]),3)
```

### Q1, Q3 Columns in Quartile_M:

#### P/FCF Q1, P/FCF Q3:
```
=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[P/FCF]),1)
=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[P/FCF]),3)
```
#### P/B Q1, P/B Q3:
```
=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[P/B]),1)
=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[P/B]),3)
```
#### ROE Q1, ROE Q3:
```
=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[ROE]),1)
=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[ROE]),3)
```
#### ROA Q1, ROA Q3:
```
=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[ROA]),1)
=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[ROA]),3)
```
#### A/E Q1, A/E Q3:
```
=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[A/E]),1)
=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[A/E]),3)
```
#### P/E Q1, P/E Q3:
```
=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[P/E]),1)
=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[P/E]),3)
```

### Q1, Q3 Columns in Quartile_S:

#### P/FCF Q1, P/FCF Q3:
```
=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[P/FCF]),1)
=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[P/FCF]),3)
```
#### P/B Q1, P/B Q3:
```
=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[P/B]),1)
=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[P/B]),3)
```
#### ROE Q1, ROE Q3:
```
=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[ROE]),1)
=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[ROE]),3)
```
#### ROA Q1, ROA Q3:
```
=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[ROA]),1)
=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[ROA]),3)
```
#### A/E Q1, A/E Q3:
```
=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[A/E]),1)
=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[A/E]),3)
```
#### P/E Q1, P/E Q3:
```
=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[P/E]),1)
=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[P/E]),3)
```

### Lower Bound, Upper Bound Columns in All Quartile Tables:

#### P/FCF LB, P/FCF UB:
```
=[@[P/FCF Q1]]-1.5*([@[P/FCF Q3]]-[@[P/FCF Q1]])
=[@[P/FCF Q3]]+1.5*([@[P/FCF Q3]]-[@[P/FCF Q1]])
```
#### P/B LB, P/B UB:
```
=[@[P/B Q1]]-1.5*([@[P/B Q3]]-[@[P/B Q1]])
=[@[P/B Q3]]+1.5*([@[P/B Q3]]-[@[P/B Q1]])
```
#### ROE LB, ROE UB:
```
=[@[ROE Q1]]-1.5*([@[ROE Q3]]-[@[ROE Q1]])
=[@[ROE Q3]]+1.5*([@[ROE Q3]]-[@[ROE Q1]])
```
#### ROA LB, ROA UB:
```
=[@[ROA Q1]]-1.5*([@[ROA Q3]]-[@[ROA Q1]])
=[@[ROA Q3]]+1.5*([@[ROA Q3]]-[@[ROA Q1]])
```
#### A/E LB, A/E UB:
```
=[@[A/E Q1]]-1.5*([@[A/E Q3]]-[@[A/E Q1]])
=[@[A/E Q3]]+1.5*([@[A/E Q3]]-[@[A/E Q1]])
```
#### P/E LB, P/E UB:
```
=[@[P/E Q1]]-1.5*([@[P/E Q3]]-[@[P/E Q1]])
=[@[P/E Q3]]+1.5*([@[P/E Q3]]-[@[P/E Q1]])
```

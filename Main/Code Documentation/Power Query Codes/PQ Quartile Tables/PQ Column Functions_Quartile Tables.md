These column functions returns Excel formulas. Check the files in **Excel Formulas** folder to see specific formulas used for columns in any worksheet.

### Functions for Lower Bound and Upper Bound Columns:

#### P/FCF LB:
```
let
    PFCF_LB = () => ("=[@[P/FCF Q1]]-1.5*([@[P/FCF Q3]]-[@[P/FCF Q1]])")
in
    PFCF_LB
```
#### P/FCF UB:
```
let
    PFCF_LB = () => ("=[@[P/FCF Q3]]+1.5*([@[P/FCF Q3]]-[@[P/FCF Q1]])")
in
    PFCF_LB
```
#### P/B LB:
```
let
    PB_LB = () => ("=[@[P/B Q1]]-1.5*([@[P/B Q3]]-[@[P/B Q1]])")
in
    PB_LB
```
#### P/B UB:
```
let
    PB_UB = () => ("=[@[P/B Q3]]+1.5*([@[P/B Q3]]-[@[P/B Q1]])")
in
    PB_UB
```
#### ROE LB:
```
let
    ROE_LB = () => ("=[@[ROE Q1]]-1.5*([@[ROE Q3]]-[@[ROE Q1]])")
in
    ROE_LB
```
#### ROE UB:
```
let
    ROE_UB = () => ("=[@[ROE Q3]]+1.5*([@[ROE Q3]]-[@[ROE Q1]])")
in
    ROE_UB
```
#### ROA LB:
```
let
    ROA_LB = () => ("=[@[ROA Q1]]-1.5*([@[ROA Q3]]-[@[ROA Q1]])")
in
    ROA_LB
```
#### ROA UB:
```
let
    ROA_UB = () => ("=[@[ROA Q3]]+1.5*([@[ROA Q3]]-[@[ROA Q1]])")
in
    ROA_UB
```
#### P/E LB:
```
let
    PE_LB = () => ("=[@[P/E Q1]]-1.5*([@[P/E Q3]]-[@[P/E Q1]])")
in
    PE_LB
```
#### P/E UB:
```
let
    PE_UB = () => ("=[@[P/E Q3]]+1.5*([@[P/E Q3]]-[@[P/E Q1]])")
in
    PE_UB
```
#### A/E LB:
```
let
    AE_LB = () => ("=[@[A/E Q1]]-1.5*([@[A/E Q3]]-[@[A/E Q1]])")
in
    AE_LB
```
#### A/E UB:
```
let
    AE_UB = () => ("=[@[A/E Q3]]+1.5*([@[A/E Q3]]-[@[A/E Q1]])")
in
    AE_UB
```

### Functions for Q1, Q3 Columns in Quartile_L:

#### L_P/FCF Q1:
```
let
    PFCF_Q1_L = () => ("=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[P/FCF]),1)")
in
    PFCF_Q1_L
```
#### L_P/FCF Q3:
```
let
    PFCF_Q3_L = () => ("=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[P/FCF]),3)")
in
    PFCF_Q3_L
```
#### L_P/B Q1:
```
let
    PB_Q1_L = () => ("=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[P/B]),1)")
in
    PB_Q1_L
```
#### L_P/B Q3:
```
let
    PB_Q3_L = () => ("=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[P/B]),3)")
in
    PB_Q3_L
```
#### L_ROE Q1:
```
let
    ROE_Q1_L = () => ("=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[ROE]),1)")
in
    ROE_Q1_L
```
#### L_ROE Q3:
```
let
    ROE_Q3_L = () => ("=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[ROE]),3)")
in
    ROE_Q3_L
```
#### L_ROA Q1:
```
let
    ROA_Q1_L = () => ("=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[ROA]),1)")
in
    ROA_Q1_L
```
#### L_ROA Q3:
```
let
    ROA_Q3_L = () => ("=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[ROA]),3)")
in
    ROA_Q3_L
```
#### L_P/E Q1:
```
let
    PE_Q1_L = () => ("=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[P/E]),1)")
in
    PE_Q1_L
```
#### L_P/E Q3:
```
let
    PE_Q3_L = () => ("=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[P/E]),3)")
in
    PE_Q3_L
```
#### L_A/E Q1:
```
let
    AE_Q1_L = () => ("=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[A/E]),1)")
in
    AE_Q1_L
```
#### L_A/E Q3:
```
let
    AE_Q3_L = () => ("=QUARTILE.INC(IF([@Industry]=Main_L[Industry],Main_L[A/E]),3)")
in
    AE_Q3_L
```

### Functions for Q1, Q3 Columns in Quartile_M:

#### M_P/FCF Q1:
```
let
    PFCF_Q1_M = () => ("=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[P/FCF]),1)")
in
    PFCF_Q1_M
```
#### M_P/FCF Q3:
```
let
    PFCF_Q3_M = () => ("=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[P/FCF]),3)")
in
    PFCF_Q3_M
```
#### M_P/B Q1:
```
let
    PB_Q1_M = () => ("=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[P/B]),1)")
in
    PB_Q1_M
```
#### M_P/B Q3:
```
let
    PB_Q3_M = () => ("=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[P/B]),3)")
in
    PB_Q3_M
```
#### M_ROE Q1:
```
let
    ROE_Q1_M = () => ("=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[ROE]),1)")
in
    ROE_Q1_M
```
#### M_ROE Q3:
```
let
    ROE_Q3_M = () => ("=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[ROE]),3)")
in
    ROE_Q3_M
```
#### M_ROA Q1:
```
let
    ROA_Q1_M = () => ("=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[ROA]),1)")
in
    ROA_Q1_M
```
#### M_ROA Q3:
```
let
    ROA_Q3_M = () => ("=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[ROA]),3)")
in
    ROA_Q3_M
```
#### M_P/E Q1:
```
let
    PE_Q1_M = () => ("=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[P/E]),1)")
in
    PE_Q1_M
```
#### M_P/E Q3:
```
let
    PE_Q3_M = () => ("=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[P/E]),3)")
in
    PE_Q3_M
```
#### M_A/E Q1:
```
let
    AE_Q1_M = () => ("=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[A/E]),1)")
in
    AE_Q1_M
```
#### M_A/E Q3:
```
let
    AE_Q3_M = () => ("=QUARTILE.INC(IF([@Industry]=Main_M[Industry],Main_M[A/E]),3)")
in
    AE_Q3_M
```

### Functions for Q1, Q3 Columns in Quartile_S:

#### S_P/FCF Q1:
```
let
    PFCF_Q1_S = () => ("=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[P/FCF]),1)")
in
    PFCF_Q1_S
```
#### S_P/FCF Q3:
```
let
    PFCF_Q3_S = () => ("=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[P/FCF]),3)")
in
    PFCF_Q3_S
```
#### S_P/B Q1:
```
let
    PB_Q1_S = () => ("=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[P/B]),1)")
in
    PB_Q1_S
```
#### S_P/B Q3:
```
let
    PB_Q3_S = () => ("=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[P/B]),3)")
in
    PB_Q3_S
```
#### S_ROE Q1:
```
let
    ROE_Q1_S = () => ("=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[ROE]),1)")
in
    ROE_Q1_S
```
#### S_ROE Q3:
```
let
    ROE_Q3_S = () => ("=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[ROE]),3)")
in
    ROE_Q3_S
```
#### S_ROA Q1:
```
let
    ROA_Q1_S = () => ("=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[ROA]),1)")
in
    ROA_Q1_S
```
#### S_ROA Q3:
```
let
    ROA_Q3_S = () => ("=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[ROA]),3)")
in
    ROA_Q3_S
```
#### S_P/E Q1:
```
let
    PE_Q1_S = () => ("=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[P/E]),1)")
in
    PE_Q1_S
```
#### S_P/E Q3:
```
let
    PE_Q3_S = () => ("=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[P/E]),3)")
in
    PE_Q3_S
```
#### S_A/E Q1:
```
let
    AE_Q1_S = () => ("=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[A/E]),1)")
in
    AE_Q1_S
```
#### S_A/E Q3:
```
let
    AE_Q3_S = () => ("=QUARTILE.INC(IF([@Industry]=Main_S[Industry],Main_S[A/E]),3)")
in
    AE_Q3_S
```
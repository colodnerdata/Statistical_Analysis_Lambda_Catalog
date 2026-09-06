<!-- GENERATED FILE — do not edit. Regenerate: uv run --group docs poe docs-generate -->
# The MODEL SPECIFICATION block (columns A–O)

One row per source-table column, header row 3, spec rows from row 4.
Orange cells are yours to type; everything else computes. The block
sizes itself from `Source_Table`, so retargeting the one name resizes
every band.

| Col | Header | How you use it |
|---|---|---|
| A | Variable | Computed — spills the table's header names. |
| B | Role | Dropdown: Response (y),Predictor (x),Identifier (Row Label),Filter,Omit,Fixed Effects. Exactly one Response; at most one Fixed Effects. |
| C | Include | Dropdown: TRUE,FALSE. |
| D | Type | Dropdown: Continuous,Categorical (Predictor rows only). |
| E | Reference Level | Typed (Categorical rows) — override the reference level; red if absent from the sample. |
| F | Order | Reserved — hidden, width 0, read by nothing in this release. |
| G | Transform | Dropdown: None,Log,Log (drop ≤ 0) (Response and Continuous Predictor rows). |
| H | Sequence | Dropdown: TRUE — mark at most one ordering axis. |
| I | Sequence Period | Typed — Δ override; blank = the computed candidate. |
| J | Period In Use | Computed — the override if typed, else the candidate. |
| K | Levels | Computed — distinct level count in the analysis sample. |
| L | Reference In Use | Computed — the reference actually in effect. |
| M | Interaction Term | Dropdown over variable names — the OTHER operand of an interaction. |
| N | Interaction Operation | Dropdown: Product,Difference,Ratio. |
| O | Design Columns | Computed — design-matrix columns this row contributes. |

## The self-sizing band every column is built on

Each spec column is a sheet-scoped named range over a fixed
16000-row band, trimmed to the table's width at calculation:

```excel
=TAKE(Regression!$B$4:$B$16000,MAX(1,COLUMNS(Source_Data)))
```

and the computed columns are single spills — Period In Use (J):

```excel
=LET(nc,COLUMNS(Source_Data),sq,TAKE(Spec_Sequence,nc),sp,TAKE(Spec_Sequence_Period,nc),cand,IFERROR(Base_Period_Delta_Candidate(),""),MAP(SEQUENCE(nc),LAMBDA(i,IF(INDEX(sq,i)<>TRUE,"",IF(N(INDEX(sp,i))<>0,INDEX(sp,i),cand)))))
```

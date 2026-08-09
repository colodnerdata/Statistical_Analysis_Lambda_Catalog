# Lambda Extraction Audit (Part 6.1)

## Status cells (first batch -- extract to sheet-scoped LAMBDAs)

| Cell | Current location | Formula | Proposed LAMBDA | Scope |
|---|---|---|---|---|
| B2 | spec_layout.py:699 _ROLE_STATUS_FORMULA | Triple-nested IF on response_count and fe_count | Role_Status() | Regression |
| H2 | spec_layout.py:722 _SEQUENCE_STATUS_FORMULA | IF on seq_flag_count > 1 | Sequence_Status() | Regression |
| G2 | spec_layout.py:753 _LOG_DOMAIN_STATUS_FORMULA | LET with MAP over columns, counts non-positive under Log | Log_Domain_Status() | Regression |
| O2 | write_sheet_regression.py:1028 | LET(k,n) with IF(k>max) and IF(OR(k>soft,n*k>soft)) | Design_Width_Status() | Regression |

## Other complex inline formulas (follow-on candidates)

| Cell | Current location | Formula | Notes |
|---|---|---|---|
| AE11 | write_sheet_regression.py:1193 | LET(seq_flags,fe_vars,IF(...)Durbin_Watson_By(...)) | 3-level nested IF + function call |
| AE12 | write_sheet_regression.py:1231 | LET(seq_flags,fe_vars,IF(...)BFN_Panel_DW(...)) | 4-level nested IF + function call |
| AK3 | write_sheet_regression.py:1544 | IF(Zero_Predictors,IF(AND,...)LET(...)Group_PI(...)) | VSTACK inside branch + inline transform dispatch |
| AK19:62 | write_sheet_regression.py per-row | IF(ROW()-offset<=ROWS(means#),INDEX(...),"") | 44 per-row formulas |

## LAMBDAs that are already clean (no action needed)

| Cell | Formula | Why it's fine |
|---|---|---|
| AL3 | Back_Transform_Response(AK3,...) | Single function call |
| AL7:10 | Back_Transform_Response(AK7,...) | Single function call |
| AN3 | Row_Labels() | Single function call |
| AO3 | Y actual | Simple |
| AP3 | Predicted Y | SUMPRODUCT |
| O1 | SUM(TAKE(...))+N(...) | Simple |

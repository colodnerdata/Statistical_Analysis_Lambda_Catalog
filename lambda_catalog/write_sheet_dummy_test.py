"""Build and write the Dummy_Test worksheet — self-checking error-contract tests
for Dummy_Levels and Dummy_Code.

Unlike the MLR test sheets, this sheet does not compare numeric Calc columns
against statsmodels: the contract under test is *typed failure* (a genuine #N/A
on degenerate input, never a descriptive string) plus the treatment-coding
behavior on a small hand-built fixture. Each case is therefore a boolean Pass
formula evaluated by Excel itself (e.g. =ISNA(Dummy_Levels(...))), and the QC
build reads the Pass cells back and reports any that are not TRUE.

The fixture values and every expected count are defined once at module level
and mirrored by pure-Python reference implementations, so the sheet writer,
the QC verifier, and the CI test suite all consume a single source of truth.
"""
from __future__ import annotations

from typing import TYPE_CHECKING

from .make_test_sheet import _ColumnSpec, write_test_table
from .workbook_helpers import reset_column_groups

if TYPE_CHECKING:
    import xlwings as xw


SHEET_NAME = "Dummy_Test"
_TABLE_HEADER_ROW = 1

# Fixture columns. Row 6 (1-based) holds a truly blank Category cell — it must
# be excluded by the blank normalization, not coerced to a numeric-zero level.
# Row 7 is include=FALSE, the excluded row that Dummy_Code must blank per-cell.
CATEGORY_VALUES: list[str] = [
    "Developed",
    "Developing",
    "Developed",
    "Developing",
    "Developing",
    "",
    "Developed",
    "Developing",
]
INCLUDE_VALUES: list[bool] = [True, True, True, True, True, True, False, True]
SINGLE_LEVEL_VALUES: list[str] = ["OnlyLevel"] * len(CATEGORY_VALUES)

_BLANK_CATEGORY_ROW = CATEGORY_VALUES.index("") + 1
_EXCLUDED_INCLUDE_ROW = INCLUDE_VALUES.index(False) + 1


def dummy_levels_reference(
    category: list[str],
    reference: str | None = None,
    include: list[bool] | bool | None = None,
) -> list[str] | None:
    """Pure-Python mirror of the Dummy_Levels LAMBDA.

    Returns the retained level labels (sorted, reference dropped), or None to
    represent the #N/A failure value. A reference of None or "" means "use
    the default" (the first sorted level) — the "" path is how a blank
    reference cell passed straight through requests the default.
    """
    if include is None:
        include = [True] * len(category)
    if isinstance(include, bool):
        include = [include] * len(category)
    active = [value != "" and bool(mask) for value, mask in zip(category, include)]
    levels = sorted({value for value, keep in zip(category, active) if keep})
    if not levels:
        return None
    ref = levels[0] if reference in (None, "") else reference
    if ref not in levels:
        return None
    retained = [level for level in levels if level != ref]
    return retained or None


def dummy_code_reference(
    category: list[str],
    reference: str | None = None,
    include: list[bool] | bool | None = None,
) -> list[list[int | str]] | None:
    """Pure-Python mirror of the Dummy_Code LAMBDA.

    Returns the full-height dummy matrix (excluded rows as ""), or None to
    represent the #N/A failure value.
    """
    retained = dummy_levels_reference(category, reference, include)
    if retained is None:
        return None
    if include is None:
        include = [True] * len(category)
    if isinstance(include, bool):
        include = [include] * len(category)
    return [
        [
            (1 if value == level else 0) if (value != "" and bool(mask)) else ""
            for level in retained
        ]
        for value, mask in zip(category, include)
    ]


def _matrix_count(matrix: list[list[int | str]], target: int | str) -> int:
    return sum(cell == target for row in matrix for cell in row)


# Expected values, derived from the mirrors (the CI suite pins the mirrors to
# independently hand-derived values, so these are not self-certifying).
EXPECTED_DEFAULT_LEVELS = dummy_levels_reference(CATEGORY_VALUES)
EXPECTED_EXPLICIT_LEVELS = dummy_levels_reference(CATEGORY_VALUES, "Developing")
_CODE_DEFAULT = dummy_code_reference(CATEGORY_VALUES)
_CODE_MASKED = dummy_code_reference(CATEGORY_VALUES, "Developed", INCLUDE_VALUES)

_CAT = f"{SHEET_NAME}[Category]"
_INC = f"{SHEET_NAME}[Include]"
_SINGLE = f"{SHEET_NAME}[Single_Level]"
_ALL_BLANK = f"{SHEET_NAME}[All_Blank]"

# (label, boolean Pass formula). Pass formulas re-call the function under test
# rather than referencing a spill anchor, so each check is self-contained.
CHECK_CASES: list[tuple[str, str]] = [
    (
        "levels_default_reference",
        f'=TEXTJOIN("|", FALSE, Dummy_Levels({_CAT}))'
        f'="{"|".join(EXPECTED_DEFAULT_LEVELS)}"',
    ),
    (
        # "" behaves exactly like an omitted reference (use the default) —
        # the Model Construction sheet passes blank Reference cells through.
        "levels_empty_string_reference_defaults",
        f'=TEXTJOIN("|", FALSE, Dummy_Levels({_CAT}, ""))'
        f'="{"|".join(EXPECTED_DEFAULT_LEVELS)}"',
    ),
    (
        "levels_explicit_reference",
        f'=TEXTJOIN("|", FALSE, Dummy_Levels({_CAT}, "Developing"))'
        f'="{"|".join(EXPECTED_EXPLICIT_LEVELS)}"',
    ),
    (
        "levels_invalid_reference_is_na",
        f'=ISNA(Dummy_Levels({_CAT}, "Developped"))',
    ),
    (
        "levels_invalid_reference_not_text",
        f'=NOT(ISTEXT(Dummy_Levels({_CAT}, "Developped")))',
    ),
    (
        "levels_single_level_is_na",
        f"=ISNA(Dummy_Levels({_SINGLE}))",
    ),
    (
        "levels_all_excluded_is_na",
        f'=ISNA(Dummy_Levels({_CAT}, "Developed", FALSE))',
    ),
    (
        "levels_all_blank_is_na",
        f"=ISNA(Dummy_Levels({_ALL_BLANK}))",
    ),
    (
        "code_default_width",
        f"=COLUMNS(Dummy_Code({_CAT}))={len(EXPECTED_DEFAULT_LEVELS)}",
    ),
    (
        "code_default_ones",
        f"=SUMPRODUCT(--(Dummy_Code({_CAT})=1))={_matrix_count(_CODE_DEFAULT, 1)}",
    ),
    (
        "code_default_zeros",
        f"=SUMPRODUCT(--(Dummy_Code({_CAT})=0))={_matrix_count(_CODE_DEFAULT, 0)}",
    ),
    (
        "code_blank_category_row_is_blank",
        f'=INDEX(Dummy_Code({_CAT}), {_BLANK_CATEGORY_ROW}, 1)=""',
    ),
    (
        "code_masked_ones",
        f'=SUMPRODUCT(--(Dummy_Code({_CAT}, "Developed", {_INC})=1))'
        f"={_matrix_count(_CODE_MASKED, 1)}",
    ),
    (
        "code_masked_zeros",
        f'=SUMPRODUCT(--(Dummy_Code({_CAT}, "Developed", {_INC})=0))'
        f"={_matrix_count(_CODE_MASKED, 0)}",
    ),
    (
        "code_excluded_row_is_blank",
        f'=INDEX(Dummy_Code({_CAT}, "Developed", {_INC}), {_EXCLUDED_INCLUDE_ROW}, 1)=""',
    ),
    (
        "code_full_height",
        f'=ROWS(Dummy_Code({_CAT}, "Developed", {_INC}))={len(CATEGORY_VALUES)}',
    ),
    (
        "code_invalid_reference_is_na",
        f'=ISNA(Dummy_Code({_CAT}, "Developped", {_INC}))',
    ),
    (
        "code_single_level_is_na",
        f"=ISNA(Dummy_Code({_SINGLE}))",
    ),
]

_CHECK_HEADER_ROW = 1
_CHECK_START_ROW = _CHECK_HEADER_ROW + 1
_CHECK_LABEL_COL = 6  # F
_CHECK_PASS_COL = 7  # G
ALL_PASS_ROW = _CHECK_START_ROW + len(CHECK_CASES)

_DISPLAY_LABEL_COL = 9  # I
_DISPLAY_SPILL_COL = 10  # J
_DISPLAY_FORMULAS: list[tuple[str, str]] = [
    ("Dummy_Levels(Category)", f"=Dummy_Levels({_CAT})"),
    ('Dummy_Levels(Category, "Developing")', f'=Dummy_Levels({_CAT}, "Developing")'),
    (
        'Dummy_Code(Category, "Developed", Include)',
        f'=Dummy_Code({_CAT}, "Developed", {_INC})',
    ),
]


def write_dummy_test_sheet(workbook: xw.Book) -> None:
    """Create or rewrite the Dummy_Test sheet: fixture table, checks, displays."""
    sheet_names = {sheet.name for sheet in workbook.sheets}
    if SHEET_NAME in sheet_names:
        sheet = workbook.sheets[SHEET_NAME]
    elif len(workbook.sheets) == 1 and not workbook.sheets[0]["A1"].value:
        sheet = workbook.sheets[0]
        sheet.name = SHEET_NAME
    else:
        sheet = workbook.sheets.add(name=SHEET_NAME, after=workbook.sheets[-1])

    for index in range(sheet.api.ListObjects.Count, 0, -1):
        sheet.api.ListObjects(index).Delete()
    sheet.api.Cells.Clear()
    reset_column_groups(sheet)

    columns: list[_ColumnSpec] = [
        ("Category", None, "General"),
        ("Include", None, "General"),
        ("Single_Level", None, "General"),
        ("All_Blank", None, "General"),
    ]
    rows_data = [
        (
            _TABLE_HEADER_ROW + 1 + i,
            {
                # A blank fixture value stays an untouched (truly empty) cell —
                # exactly the input the blank normalization must handle.
                **({"Category": CATEGORY_VALUES[i]} if CATEGORY_VALUES[i] != "" else {}),
                "Include": INCLUDE_VALUES[i],
                "Single_Level": SINGLE_LEVEL_VALUES[i],
            },
        )
        for i in range(len(CATEGORY_VALUES))
    ]
    write_test_table(sheet, SHEET_NAME, columns, _TABLE_HEADER_ROW, rows_data)

    sheet.range((_CHECK_HEADER_ROW, _CHECK_LABEL_COL)).value = "Check"
    sheet.range((_CHECK_HEADER_ROW, _CHECK_PASS_COL)).value = "Pass"
    sheet.range((_CHECK_HEADER_ROW, _CHECK_LABEL_COL)).api.Font.Bold = True
    sheet.range((_CHECK_HEADER_ROW, _CHECK_PASS_COL)).api.Font.Bold = True
    for i, (label, formula) in enumerate(CHECK_CASES):
        sheet.range((_CHECK_START_ROW + i, _CHECK_LABEL_COL)).value = label
        sheet.range((_CHECK_START_ROW + i, _CHECK_PASS_COL)).formula = formula
    sheet.range((ALL_PASS_ROW, _CHECK_LABEL_COL)).value = "ALL PASS"
    sheet.range((ALL_PASS_ROW, _CHECK_LABEL_COL)).api.Font.Bold = True
    pass_first = sheet.range((_CHECK_START_ROW, _CHECK_PASS_COL)).get_address(
        row_absolute=False, column_absolute=False
    )
    pass_last = sheet.range((ALL_PASS_ROW - 1, _CHECK_PASS_COL)).get_address(
        row_absolute=False, column_absolute=False
    )
    sheet.range((ALL_PASS_ROW, _CHECK_PASS_COL)).formula = f"=AND({pass_first}:{pass_last})"
    sheet.range((ALL_PASS_ROW, _CHECK_PASS_COL)).api.Font.Bold = True

    display_row = _CHECK_HEADER_ROW
    sheet.range((display_row, _DISPLAY_LABEL_COL)).value = "Live results"
    sheet.range((display_row, _DISPLAY_LABEL_COL)).api.Font.Bold = True
    for i, (label, formula) in enumerate(_DISPLAY_FORMULAS):
        # Row-height gaps keep the third display's 8-row spill clear of others.
        row = display_row + 1 + i * 2
        sheet.range((row, _DISPLAY_LABEL_COL)).value = label
        sheet.range((row, _DISPLAY_SPILL_COL)).formula = formula

    sheet.range((1, _CHECK_LABEL_COL)).column_width = 34
    sheet.range((1, _DISPLAY_LABEL_COL)).column_width = 38

    sheet.activate()
    sheet.api.Application.ActiveWindow.SplitRow = 1
    sheet.api.Application.ActiveWindow.SplitColumn = 0
    sheet.api.Application.ActiveWindow.FreezePanes = True


def read_dummy_check_failures(workbook: xw.Book) -> list[str]:
    """Read every Pass cell on the Dummy_Test sheet; return QC failure messages.

    A Pass cell that evaluates to anything other than TRUE (FALSE, an error
    value, or blank) produces one failure message in the standard QC format.
    """
    sheet = workbook.sheets[SHEET_NAME]
    failures: list[str] = []
    for i, (label, _) in enumerate(CHECK_CASES):
        value = sheet.range((_CHECK_START_ROW + i, _CHECK_PASS_COL)).value
        if value is not True:
            failures.append(
                f"[Dummy_Test] check={label!r}: expected TRUE, excel_calc={value!r}"
            )
    return failures

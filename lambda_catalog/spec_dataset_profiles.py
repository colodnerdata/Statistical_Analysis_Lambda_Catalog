"""
spec_dataset_profiles.py

Per-dataset spec profiles for the Regression sheet's Source_Table retarget.

SpecDatasetProfile wraps a dataset's variable list, default Role/Include/Type
spec, and Sequence-flagged columns as one unit, so retargeting Source_Table
(the --regression-dataset CLI choice) also retargets the spec block's defaults.

Three profiles are defined: Auto MPG (the default), Life Expectancy (WHO), and
Production Lots (learning-curve panel). The SPEC_DATASET_PROFILES registry maps
the CLI choice string to its profile.

Extracted from write_spec_block.py to separate the dataset data from the
spec-block writer logic.
"""
from __future__ import annotations

from dataclasses import dataclass

from .regression_shared import FEATURE_COLUMNS as _LIFE_EXPECTANCY_FEATURE_COLUMNS
from .spec_layout import (
    _DEFAULT_SEQUENCE_VARIABLES,
    _DEFAULT_SPEC,
    _ROLE_FILTER,
    _ROLE_FIXED_EFFECTS,
    _ROLE_IDENTIFIER,
    _ROLE_OMIT,
    _ROLE_PREDICTOR,
    _ROLE_RESPONSE,
    _VARIABLES,
)
from .write_sheet_csv_dataset import LIFE_EXPECTANCY, MILEAGE, PRODUCTION_LOTS

# ── Per-dataset spec profiles ──────────────────────────────────────────────
# _VARIABLES/_DEFAULT_SPEC/_DEFAULT_SEQUENCE_VARIABLES above are the shipped
# Auto MPG defaults; SpecDatasetProfile wraps a dataset's variable list,
# default Role/Include/Type spec, and Sequence-flagged columns as one unit
# so retargeting Source_Table (the --regression-dataset CLI choice) can
# also retarget the spec block's defaults, instead of leaving every column
# of a newly-targeted dataset to _FALLBACK_SPEC's un-flagged Predictor.
# The profile decides which rows arrive PRE-FILLED, not how many spec rows
# exist — the block sizes itself from COLUMNS(Source_Data). So every column
# of the targeted dataset carries a real Role/Include/Type from the first
# build instead of falling back to _FALLBACK_SPEC, and a dataset the build
# did not target still gets a working (if unfilled) block on retarget.
@dataclass(frozen=True)
class SpecDatasetProfile:
    """One dataset's Source_Table target and shipped spec-block defaults."""

    source_table_ref: str
    variables: tuple[str, ...]
    default_spec: dict[str, tuple[str, bool, str]]
    sequence_variables: frozenset[str] = frozenset()


_AUTO_MPG_PROFILE = SpecDatasetProfile(
    source_table_ref=f"={MILEAGE.table_name}[#All]",
    variables=tuple(_VARIABLES),
    default_spec=_DEFAULT_SPEC,
    sequence_variables=_DEFAULT_SEQUENCE_VARIABLES,
)

# Column order matches the Life Expectancy CSV's normalized header order
# plus the appended "Developed Country after 2013" derived column. Country is
# a text identifier (row labeling only), Year is the natural panel ordering
# axis (Sequence-flagged, Role Omit so it never enters the design matrix
# itself), Status is the one categorical predictor, and "Developed Country
# after 2013" (=AND(Status="Developed", Year>2013)) ships Omit — dormant. A
# Filter role is always-on regardless of Include (Sample_Include applies every
# Filter column unconditionally), so a column with FALSE values would actively
# restrict the sample; Omit keeps the shipped default fitting all rows, one
# Role-dropdown flip from becoming a developed-country-after-2013 filter.
#
# The SHIPPED DEFAULT is the curated four-driver model both presentation decks
# headline (the slide-19 coefficient table): Life expectancy ~ Adult Mortality
# + Alcohol + percentage expenditure + C(Status). Those four predictors ship
# Include=True; every other FEATURE_COLUMNS entry ships Include=False — present
# in the block (so the block still sizes to all 23 source columns) but off,
# ready to toggle on for the EDA / VIF-trim / kitchen-sink beats. This is the
# cold open: k ≈ 5, fast on a low-compute machine, matching the deck. The
# 18-predictor kitchen sink is still a registered case (L05,
# ``_life_full_profile_spec``), just no longer what the workbook opens with —
# and the curated default's own oracle is the registered case ``life_talk_demo``
# (L11). ``_LIFE_EXPECTANCY_VARIABLES`` is unchanged: the block sizes itself
# from COLUMNS(Source_Data), so the 19 dormant rows cost nothing and are one
# toggle away.
_LIFE_EXPECTANCY_VARIABLES: tuple[str, ...] = (
    "Country",
    "Year",
    "Status",
    "Life expectancy",
    *_LIFE_EXPECTANCY_FEATURE_COLUMNS,
    "Developed Country after 2013",
)
# The three curated drivers that ship active.
_LIFE_TALK_DEMO_PREDICTORS: frozenset[str] = frozenset(
    {"Adult Mortality", "Alcohol", "percentage expenditure"}
)
_LIFE_EXPECTANCY_DEFAULT_SPEC: dict[str, tuple[str, bool, str]] = {
    "Country": (_ROLE_IDENTIFIER, False, "Continuous"),
    "Year": (_ROLE_OMIT, False, "Continuous"),
    "Status": (_ROLE_PREDICTOR, True, "Categorical"),
    "Life expectancy": (_ROLE_RESPONSE, False, "Continuous"),
    **{
        column: (
            _ROLE_PREDICTOR,
            column in _LIFE_TALK_DEMO_PREDICTORS,
            "Continuous",
        )
        for column in _LIFE_EXPECTANCY_FEATURE_COLUMNS
    },
    "Developed Country after 2013": (_ROLE_OMIT, False, "Continuous"),
}
_LIFE_EXPECTANCY_SEQUENCE_VARIABLES: frozenset[str] = frozenset({"Year"})
_LIFE_EXPECTANCY_PROFILE = SpecDatasetProfile(
    source_table_ref=f"={LIFE_EXPECTANCY.table_name}[#All]",
    variables=_LIFE_EXPECTANCY_VARIABLES,
    default_spec=_LIFE_EXPECTANCY_DEFAULT_SPEC,
    sequence_variables=_LIFE_EXPECTANCY_SEQUENCE_VARIABLES,
)

# Column order matches the Production Lots CSV's header order — the same
# shape as analyze_regression_spec.py's _production_lots_fixed_effects_spec(),
# the QC-validated Crawford/Wright learning-curve model (ln(unit cost) = a +
# b*ln(cumulative units)), reused here verbatim so the shipped default
# matches a spec the test suite already proves fits correctly: Facility is
# the Fixed Effects panel-grouping column, Fiscal_Year is the Sequence
# axis, log Cum Units is the sole predictor, log Unit Cost is the response.
# The Production Lots sheet ships no appended derived column: Full_Data was
# always TRUE for every row (the dataset has no missing values), so the
# Filter was a no-op and is gone. No shipped dataset column is Role=Filter
# by default now; the Filter role is exercised by the Is_USA fixture (M15)
# in the test-model suite.
_PRODUCTION_LOTS_VARIABLES: tuple[str, ...] = (
    "Lot_ID",
    "Facility",
    "Fiscal_Year",
    "Lot_Quantity",
    "Cumulative_Units",
    "Experience_Stock",
    "Unit_Cost_BY",
    "log Cum Units",
    "log experience",
    "log Unit Cost",
)
_PRODUCTION_LOTS_DEFAULT_SPEC: dict[str, tuple[str, bool, str]] = {
    "Lot_ID": (_ROLE_IDENTIFIER, False, "Continuous"),
    "Facility": (_ROLE_FIXED_EFFECTS, False, "Continuous"),
    "Fiscal_Year": (_ROLE_OMIT, False, "Continuous"),
    "Lot_Quantity": (_ROLE_OMIT, False, "Continuous"),
    "Cumulative_Units": (_ROLE_OMIT, False, "Continuous"),
    "Experience_Stock": (_ROLE_OMIT, False, "Continuous"),
    "Unit_Cost_BY": (_ROLE_OMIT, False, "Continuous"),
    "log Cum Units": (_ROLE_PREDICTOR, True, "Continuous"),
    "log experience": (_ROLE_OMIT, False, "Continuous"),
    "log Unit Cost": (_ROLE_RESPONSE, False, "Continuous"),
}
_PRODUCTION_LOTS_SEQUENCE_VARIABLES: frozenset[str] = frozenset({"Fiscal_Year"})
_PRODUCTION_LOTS_PROFILE = SpecDatasetProfile(
    source_table_ref=f"={PRODUCTION_LOTS.table_name}[#All]",
    variables=_PRODUCTION_LOTS_VARIABLES,
    default_spec=_PRODUCTION_LOTS_DEFAULT_SPEC,
    sequence_variables=_PRODUCTION_LOTS_SEQUENCE_VARIABLES,
)

# The --regression-dataset CLI choice (build_production.py) indexes this
# registry for both the Source_Table retarget and the spec-block defaults —
# adding a new dataset means adding one SpecDatasetProfile and one entry
# here, nothing else.
SPEC_DATASET_PROFILES: dict[str, SpecDatasetProfile] = {
    "auto_mpg": _AUTO_MPG_PROFILE,
    "life_expectancy": _LIFE_EXPECTANCY_PROFILE,
    "production_lots": _PRODUCTION_LOTS_PROFILE,
}

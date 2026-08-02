"""Small xlwings-compatible recorder for testing sheet writers without Excel."""
from __future__ import annotations

from dataclasses import dataclass, field
from types import SimpleNamespace
from typing import Any, cast


class RecordingName:
    def __init__(self, collection: "RecordingNames", name: str, refers_to: str = "") -> None:
        self._collection = collection
        self.Name = name
        self.RefersTo = refers_to

    def Delete(self) -> None:
        self._collection.items.remove(self)


class RecordingNames:
    def __init__(self, scope_prefix: str = "", names: list[str] | None = None) -> None:
        self.scope_prefix = scope_prefix
        self.items = [RecordingName(self, name) for name in names or []]

    @property
    def Count(self) -> int:
        return len(self.items)

    def __call__(self, index: int | str) -> RecordingName:
        if isinstance(index, str):
            for item in self.items:
                if item.Name.lower() == index.lower():
                    return item
            raise OSError(f"Name not found: {index}")
        return self.items[index - 1]

    def Add(self, *, Name: str, RefersTo: str) -> RecordingName:
        qualified_name = f"{self.scope_prefix}{Name}"
        item = RecordingName(self, qualified_name, RefersTo)
        self.items.append(item)
        return item

    def by_short_name(self, name: str) -> RecordingName:
        return next(item for item in self.items if item.Name.split("!", 1)[-1] == name)


@dataclass
class RecordingCondition:
    Type: int
    Formula1: str
    Interior: Any = field(default_factory=lambda: SimpleNamespace(Color=None))
    Font: Any = field(
        default_factory=lambda: SimpleNamespace(
            Color=None,
            Bold=None,
            Strikethrough=None,
        )
    )
    StopIfTrue: bool | None = None


class RecordingFormatConditions:
    def __init__(self) -> None:
        self.items: list[RecordingCondition] = []
        self.color_scales: list[Any] = []

    def Add(self, *, Type: int, Formula1: str) -> RecordingCondition:
        condition = RecordingCondition(Type=Type, Formula1=Formula1)
        self.items.append(condition)
        return condition

    def Delete(self) -> None:
        self.items.clear()
        self.color_scales.clear()

    def AddColorScale(self, count: int) -> Any:
        criteria = {
            index: SimpleNamespace(
                Type=None,
                Value=None,
                FormatColor=SimpleNamespace(Color=None),
            )
            for index in range(1, count + 1)
        }
        scale = SimpleNamespace(ColorScaleCriteria=lambda index: criteria[index])
        self.color_scales.append(scale)
        return scale


class RecordingValidation:
    def __init__(self) -> None:
        self.rules: list[dict[str, Any]] = []
        self.delete_count = 0
        self.IgnoreBlank: bool | None = None

    def Delete(self) -> None:
        self.delete_count += 1
        self.rules.clear()

    def Add(self, *, Type: int, AlertStyle: int, Operator: int, Formula1: str) -> None:
        self.rules.append(
            {
                "Type": Type,
                "AlertStyle": AlertStyle,
                "Operator": Operator,
                "Formula1": Formula1,
            }
        )


@dataclass
class RecordingRangeState:
    value: Any = None
    formula2: str | None = None
    formula: str | None = None
    color: Any = None
    number_format: str | None = None
    column_width: float | None = None


class RecordingRangeApi:
    def __init__(self, state: RecordingRangeState, sheet: "RecordingSheet", address: tuple[Any, ...]) -> None:
        self._state = state
        self._sheet = sheet
        self.address = address
        self._block_formula2 = False
        self._borders: dict[int, Any] = {}
        self.Font = SimpleNamespace(Bold=None, Color=None, Strikethrough=None)
        self.Interior = SimpleNamespace(Color=None)
        self.FormatConditions = RecordingFormatConditions()
        self.Validation = RecordingValidation()

    def Table(self, *, RowInput: Any, ColumnInput: Any) -> None:
        self._sheet.tables.append({
            "range": self.address,
            "row_input": RowInput.address,
            "column_input": ColumnInput.address,
        })

    def Borders(self, edge: int) -> Any:
        return self._borders.setdefault(
            edge,
            SimpleNamespace(LineStyle=None, Weight=None),
        )

    def _block_bounds(self) -> tuple[int, int, int, int] | None:
        """``(r1, c1, r2, c2)`` when this range spans a ``(top_left, bottom_right)`` block."""
        if (
            len(self.address) == 2
            and isinstance(self.address[0], tuple)
            and isinstance(self.address[1], tuple)
        ):
            (r1, c1), (r2, c2) = self.address
            return r1, c1, r2, c2
        return None

    @property
    def Formula2(self) -> Any:
        # Excel returns a scalar formula string for a single cell and a 2D
        # tuple for a multi-cell range; mirror that.  The block form is
        # reconstructed from the constituent cells so the per-cell states stay
        # the single record of what was written.
        if self._block_formula2:
            r1, c1, r2, c2 = cast(tuple[int, int, int, int], self._block_bounds())
            return tuple(
                tuple(self._sheet.cell(r, c).api._state.formula2 for c in range(c1, c2 + 1))
                for r in range(r1, r2 + 1)
            )
        return self._state.formula2

    @Formula2.setter
    def Formula2(self, value: Any) -> None:
        # Distribute a 2D list across the constituent cells, mirroring how
        # Excel expands a multi-cell ``Range.Formula2`` assignment.  Lets a
        # headless test assert ``sheet.cell(r, c).api.Formula2`` per cell after
        # the writer writes a whole grid block in one COM call.  The 2D list is
        # deliberately NOT stored on this range's own ``RecordingRangeState``,
        # whose ``formula2`` field is the per-cell formula string.
        bounds = self._block_bounds()
        if (
            bounds is not None
            and isinstance(value, list)
            and value
            and isinstance(value[0], list)
        ):
            r1, c1, _, _ = bounds
            for i, row_vals in enumerate(value):
                for j, cell_value in enumerate(row_vals):
                    self._sheet.cell(r1 + i, c1 + j).api._state.formula2 = cell_value
            self._block_formula2 = True
            return
        self._block_formula2 = False
        self._state.formula2 = value

    @property
    def Formula(self) -> str | None:
        return self._state.formula

    @Formula.setter
    def Formula(self, value: str) -> None:
        self._state.formula = value


class RecordingRange:
    def __init__(self, sheet: "RecordingSheet", address: tuple[Any, ...]) -> None:
        self._sheet = sheet
        self.address = address
        self.state = RecordingRangeState()
        self.api = RecordingRangeApi(self.state, sheet, address)

    def merge(self) -> None:
        self._sheet.merges.append(self.address)

    @property
    def value(self) -> Any:
        return self.state.value

    @value.setter
    def value(self, value: Any) -> None:
        self.state.value = value

    @property
    def color(self) -> Any:
        return self.state.color

    @color.setter
    def color(self, value: Any) -> None:
        self.state.color = value

    @property
    def number_format(self) -> str | None:
        return self.state.number_format

    @number_format.setter
    def number_format(self, value: str) -> None:
        self.state.number_format = value

    @property
    def column_width(self) -> float | None:
        return self.state.column_width

    @column_width.setter
    def column_width(self, value: float) -> None:
        self.state.column_width = value


class RecordingListColumns:
    def __init__(self, table: "RecordingListObject") -> None:
        self._table = table
        self.items: list[RecordingListColumn] = []

    def Add(self) -> "RecordingListColumn":
        column = RecordingListColumn(self._table)
        self.items.append(column)
        return column

    def __call__(self, name: str) -> "RecordingListColumn":
        for column in self.items:
            if column.Name == name:
                return column
        raise KeyError(name)


@dataclass
class RecordingListColumn:
    table: Any
    Name: str = ""
    DataBodyRange: Any = field(default_factory=lambda: SimpleNamespace(Formula=None))

    def Delete(self) -> None:
        self.table.ListColumns.items.remove(self)


class RecordingListObject:
    def __init__(self, source_range: Any) -> None:
        self.Name: str = ""
        self.TableStyle: str = ""
        self.ShowTableStyleRowStripes: bool = False
        self.ShowTableStyleColumnStripes: bool = False
        self.ListColumns = RecordingListColumns(self)
        self._source_range = source_range
        self._body_range = SimpleNamespace(Formula=None)
        self.HeaderRowRange = source_range

    @property
    def DataBodyRange(self) -> Any:
        return self._body_range


class RecordingListObjects:
    def __init__(self) -> None:
        self.items: list[RecordingListObject] = []

    def Add(
        self,
        *,
        SourceType: int = 0,
        Source: Any = None,
        XlListObjectHasHeaders: int = 0,
    ) -> RecordingListObject:
        del SourceType, XlListObjectHasHeaders
        table = RecordingListObject(Source)
        self.items.append(table)
        return table


class RecordingColumns:
    """Records outline Group/Ungroup calls made via ``sheet.api.Columns(addr)``.

    Several writers fold columns into Excel outline groups
    (``sheet.api.Columns("BO:BO").Group()``); the full regression writer does
    it for every zone, and the §4b materialization zone does it for its two
    bounded blocks. Recording the grouped addresses lets a headless test
    confirm a column was grouped (and, just as importantly, that a gutter was
    NOT).
    """

    def __init__(self, sheet: "RecordingSheet") -> None:
        self._sheet = sheet

    def __call__(self, address: str) -> Any:
        return SimpleNamespace(
            Group=lambda: self._sheet.column_groups.append(address),
            Ungroup=lambda: (
                self._sheet.column_groups.remove(address)
                if address in self._sheet.column_groups
                else None
            ),
        )


class RecordingSheet:
    def __init__(self, name: str = "Univariate", global_names: list[str] | None = None) -> None:
        self.name = name
        self.ranges: dict[tuple[Any, ...], RecordingRange] = {}
        self.merges: list[tuple[Any, ...]] = []
        self.tables: list[dict[str, Any]] = []
        self.list_objects: list[RecordingListObject] = []
        self.column_groups: list[str] = []
        self.api = SimpleNamespace(
            Names=RecordingNames(f"{name}!"),
            ListObjects=RecordingListObjects(),
            Columns=RecordingColumns(self),
            # ``Cells`` is used as a wholesale clear (``Cells.Clear()``) and to
            # drop outline levels (``Cells.ClearOutline()``) before re-grouping;
            # neither is load-bearing for a cell/name assertion, so they no-op.
            Cells=SimpleNamespace(
                Clear=lambda: None,
                ClearOutline=lambda: setattr(self, "column_groups", []),
            ),
        )
        self.book = SimpleNamespace(
            api=SimpleNamespace(Names=RecordingNames(names=global_names))
        )

    def range(self, *addresses: Any) -> RecordingRange:
        key = tuple(addresses)
        if key not in self.ranges:
            self.ranges[key] = RecordingRange(self, key)
        return self.ranges[key]

    def cell(self, row: int, col: int) -> RecordingRange:
        return self.range((row, col))

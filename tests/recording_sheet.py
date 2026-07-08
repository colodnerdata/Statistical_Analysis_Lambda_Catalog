"""Small xlwings-compatible recorder for testing sheet writers without Excel."""
from __future__ import annotations

from dataclasses import dataclass, field
from types import SimpleNamespace
from typing import Any


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
    color: Any = None
    number_format: str | None = None
    column_width: float | None = None


class RecordingRangeApi:
    def __init__(self, state: RecordingRangeState, sheet: "RecordingSheet", address: tuple[Any, ...]) -> None:
        self._state = state
        self._sheet = sheet
        self.address = address
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

    @property
    def Formula2(self) -> str | None:
        return self._state.formula2

    @Formula2.setter
    def Formula2(self, value: str) -> None:
        self._state.formula2 = value


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


class RecordingSheet:
    def __init__(self, name: str = "Univariate", global_names: list[str] | None = None) -> None:
        self.name = name
        self.ranges: dict[tuple[Any, ...], RecordingRange] = {}
        self.merges: list[tuple[Any, ...]] = []
        self.tables: list[dict[str, Any]] = []
        self.api = SimpleNamespace(Names=RecordingNames(f"{name}!"))
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

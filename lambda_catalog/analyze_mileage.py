"""Mileage (Auto MPG) dataset path anchor for the spec-driven Regression oracle."""
from __future__ import annotations

from pathlib import Path

from .write_sheet_csv_dataset import MILEAGE

ROOT_DIR = Path(__file__).resolve().parent.parent
DEFAULT_INPUT_CSV = MILEAGE.default_csv_path
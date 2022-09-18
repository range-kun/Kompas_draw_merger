from collections import namedtuple
from dataclasses import dataclass
from enum import Enum, IntEnum, auto
from typing import Any, NewType, NamedTuple

from pydantic import BaseModel


FilePath = NewType("FilePath", str)
FileName = NewType("FileName", str)
DrawObozn = NewType("DrawObozn", str)
DrawName = NewType("DrawName", str)
DrawExecution = NewType("DrawExecution", str)


class SaveType(Enum):
    AUTO_SAVE_FOLDER = auto()
    MANUALLY_SAVE_FOLDER = auto()


class DrawType(IntEnum):
    ASSEMBLY_DRAW = 5
    SPEC_DRAW = 15
    DETAIL = 20


class ErrorType(Enum):
    FILE_MISSING = auto()
    FILE_NOT_OPENED = auto()


@dataclass
class DrawData:
    draw_obozn: DrawObozn
    draw_name: DrawName | None = None


@dataclass
class SpecSectionData:
    draw_type: DrawType
    draw_names: list[DrawData]


class UserSettings(BaseModel):
    except_folders_list: list[str] | None
    constructor_list: list[str] | None
    checker_list: list[str] | None
    sortament_list: list[str] | None
    watermark_position: list[int]
    watermark_path: FilePath
    add_default_watermark: bool = True
    split_file_by_size: bool = False
    auto_save_folder: bool = False


@dataclass
class Filters:
    date_range: list[int] | None = None
    constructor_list: list[str] | None = None
    checker_list: list[str] | None = None
    sortament_list: list[str] | None = None


@dataclass
class SettingsData:
    filters: Filters = Filters()
    watermark_path: FilePath | None = None
    watermark_position: list[int] | None = None
    split_file_by_size: bool = False
    save_type: SaveType = SaveType.AUTO_SAVE_FOLDER
    except_folders_list: list[str] | None = None


FilterWidgetPositions = namedtuple(
    "FilterWidgetPositions",
    "check_box_position "
    "combobox_position "
    "input_line_position "
    "combobox_radio_button_position "
    "input_radio_button_position "
)


class ThreadKompasAPI(NamedTuple):
    kompas_api7_module: Any
    application_stream: Any
    const: Any


@dataclass
class DoublePathsData:
    draw_obozn: DrawObozn
    spw_paths: list[FilePath]
    cdw_paths: list[FilePath]


@dataclass
class MergerData:
    delete_single_draws_after_merge_checkbox: bool
    specification_path: FilePath | None = None

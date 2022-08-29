from collections import namedtuple
from dataclasses import dataclass
from enum import Enum, IntEnum
from typing import Any, NewType, NamedTuple

from pydantic import BaseModel


FilePath = NewType('FilePath', str)
FileName = NewType('FileName', str)
DrawObozn = NewType('DrawObozn', str)
DrawName = NewType('DrawObozn', str)
DrawExecution = NewType('DrawExecution', str)


class SaveType(Enum):
    AUTO_SAVE_FOLDER = 1
    MANUALLY_SAVE_FOLDER = 2


class DrawType(IntEnum):
    ASSEMBLY_DRAW = 5
    SPEC_DRAW = 15
    DETAIL = 20


class DrawData(BaseModel):
    draw_obozn: DrawObozn
    draw_name: DrawName


class SpecSectionData(BaseModel):
    draw_type: DrawType
    draw_names: list[DrawData]


class UserSettings(BaseModel):
    except_folders_list: list[str] | None
    constructor_list: list[str] | None
    checker_list: list[str] | None
    sortament_list: list[str] | None
    watermark_position: list[int]
    add_default_watermark: bool = True
    split_file_by_size: bool = False
    auto_save_folder: bool = False
    watermark_path: FilePath


class Filters(BaseModel):
    date_range: list[int, int] | None = None
    constructor_list: list[str] | None = None
    checker_list: list[str] | None = None
    sortament_list: list[str] | None = None


class SettingsData(BaseModel):
    filters: Filters | None = None
    watermark_path: FilePath | None = None
    watermark_position: list[int] | None = None
    split_file_by_size: bool = False
    save_type: SaveType = SaveType.AUTO_SAVE_FOLDER
    except_folders_list: list[str] | None = None


FilterWidgetPositions = namedtuple(
    'FilterWidgetPositions',
    'check_box_position '
    'combobox_position '
    'input_line_position '
    'combobox_radio_button_position '
    'input_radio_button_position '
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

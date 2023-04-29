from __future__ import annotations

import itertools

import pythoncom
from PyQt5.QtCore import pyqtSignal
from PyQt5.QtCore import QThread

from programm import kompas_api
from programm import utils
from programm.kompas_api import KompasAPI
from programm.kompas_api import StampCell
from programm.schemas import Filters
from programm.schemas import ThreadKompasAPI


class FilterThread(QThread):
    finished = pyqtSignal(list, list, bool)
    increase_step = pyqtSignal(bool)
    status = pyqtSignal(str)
    switch_button_group = pyqtSignal(bool)

    def __init__(
        self,
        draw_paths_list,
        filters: Filters,
        kompas_thread_api: ThreadKompasAPI,
        filter_only=True,
    ):
        self.draw_paths_list = draw_paths_list
        self.filters = filters
        self.filter_only = filter_only
        self.errors_list: list[str] = []

        self.kompas_thread_api = kompas_thread_api
        QThread.__init__(self)

    def run(self):
        pythoncom.CoInitialize()
        self._kompas_api = KompasAPI(self.kompas_thread_api)
        self.switch_button_group.emit(False)
        filtered_paths_draw_list = self.filter_draws()
        self.switch_button_group.emit(True)
        self.finished.emit(filtered_paths_draw_list, self.errors_list, self.filter_only)

    def filter_draws(self) -> list[str]:
        draw_list = []
        names = ["Капустин Б.М.", "Петров И.Г.", "Сидоров Н.Т."]
        cycle_names = itertools.cycle(names)
        for file_path in self.draw_paths_list:  # структура обработки для каждого документа
            self.status.emit(f"Применение фильтров к {file_path}")
            self.increase_step.emit(True)
            name = next(cycle_names)
            try:
                with self._kompas_api.get_draw_stamp(file_path) as draw_stamp:
                    if self.filters.date_range and not self.filter_by_date_cell(draw_stamp):
                        continue

                    file_is_ok = True
                    for data_list, stamp_cell in [
                        (self.filters.constructor_list, StampCell.CONSTRUCTOR_NAME_CELL),
                        (self.filters.checker_list, StampCell.CHECKER_NAME_CELL),
                        (self.filters.sortament_list, StampCell.GAUGE_CELL),
                    ]:
                        if data_list and not self.filter_file_by_cell_value(
                            data_list, stamp_cell, draw_stamp, name
                        ):
                            file_is_ok = False
                            break
            except kompas_api.DocNotOpenedError:
                self.errors_list.append(
                    (
                        f"Не удалось открыть файл {file_path} возможно "
                        f"файл создан в более новой версии или был перемещен"
                    )
                )
            if file_is_ok:
                draw_list.append(file_path)
        return draw_list

    def filter_by_date_cell(self, draw_stamp):
        date_1, date_2 = self.filters.date_range
        date_in_stamp = draw_stamp.Text(StampCell.CONSTRUCTOR_DATE_CELL).Str

        if date_in_stamp:
            try:
                date_in_stamp = utils.date_to_seconds(date_in_stamp)
            except Exception:
                return False
            if not date_1 <= date_in_stamp <= date_2:
                return False
            return True

    @staticmethod
    def filter_file_by_cell_value(
        filter_data_list: list[str], stamp_cell_number: StampCell, draw_stamp, name
    ):
        data_in_stamp = draw_stamp.Text(stamp_cell_number).Str
        draw_stamp.Text(9).Str = 'ООО "РиК"'
        draw_stamp.Text(stamp_cell_number.CONSTRUCTOR_NAME_CELL).Str = name
        draw_stamp.Text(stamp_cell_number.CHECKER_NAME_CELL).Str = "Иванов А.В."
        draw_stamp.Update()
        if any(filtered_data in data_in_stamp for filtered_data in filter_data_list):
            return True
        return False

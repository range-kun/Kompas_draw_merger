from __future__ import annotations

import os

import pythoncom
import win32com
from PyQt5.QtCore import pyqtSignal
from PyQt5.QtCore import QThread

from programm import kompas_api
from programm import utils
from programm.kompas_api import KompasAPI
from programm.kompas_api import StampCell
from programm.schemas import DoublePathsData
from programm.schemas import DrawErrorsType
from programm.schemas import DrawObozn
from programm.schemas import ErrorType
from programm.schemas import FilePath
from programm.schemas import ThreadKompasAPI


class DataBaseThread(QThread):
    increase_step = pyqtSignal(bool)
    status = pyqtSignal(str)
    finished = pyqtSignal(dict, dict, bool)
    progress_bar = pyqtSignal(int)
    calculate_step = pyqtSignal(int, bool, bool)
    buttons_enable = pyqtSignal(bool)
    errors = pyqtSignal(str)

    def __init__(
        self,
        draw_paths: list[FilePath],
        need_to_merge: bool,
        kompas_thread_api: ThreadKompasAPI,
    ):
        self.draw_paths = draw_paths
        self.need_to_merge = need_to_merge
        self.kompas_thread_api = kompas_thread_api

        self.draws_data_base: dict[DrawObozn, list[FilePath]] = {}
        self.errors_dict: dict[ErrorType:DrawErrorsType] = {
            ErrorType.FILE_ERRORS: [],
            ErrorType.FILE_NAMING: [],
        }
        QThread.__init__(self)

    def run(self):
        pythoncom.CoInitializeEx(0)
        self.shell = win32com.client.gencache.EnsureDispatch(
            "Shell.Application", 0
        )  # подключаемся к windows shell

        self.buttons_enable.emit(False)
        obozn_meta_number = self._get_meta_obozn_number(os.path.dirname(self.draw_paths[0]))
        if not obozn_meta_number:
            self.errors.emit("Ошибка при создания базы чертежей")
            return

        self._kompas_api = KompasAPI(self.kompas_thread_api)
        # нужно создавать именно в run для правильной работы
        self._create_data_base(obozn_meta_number)

        self.progress_bar.emit(0)
        if double_paths_list := self._get_list_of_paths_with_extra_obozn():
            self._proceed_double_paths(double_paths_list)

        self.progress_bar.emit(0)
        self.buttons_enable.emit(True)
        self.finished.emit(self.draws_data_base, self.errors_dict, self.need_to_merge)

    def _get_meta_obozn_number(self, dir_name: str) -> int | None:
        dir_obj = self.shell.NameSpace(dir_name)  # получаем объект папки windows shell
        for number in range(355):
            if dir_obj.GetDetailsOf(None, number) == "Обозначение":
                return number
        return None

    def _create_data_base(self, meta_obozn_number: int):
        for draw_path in self.draw_paths:
            self.status.emit(f"Получение атрибутов {draw_path}")
            self.increase_step.emit(True)
            if draw_obozn := self._fetch_draw_obozn(draw_path, meta_obozn_number):
                if draw_obozn in self.draws_data_base.keys():
                    self.draws_data_base[draw_obozn].append(draw_path)
                else:
                    self.draws_data_base[draw_obozn] = [draw_path]

    def _fetch_draw_obozn(self, draw_path: FilePath, meta_obozn_number: int) -> DrawObozn:
        dir_obj = self.shell.NameSpace(
            os.path.dirname(draw_path)
        )  # получаем объект папки windows shell
        item = dir_obj.ParseName(
            os.path.basename(draw_path)
        )  # указатель на файл (делаем именно объект windows shell)
        draw_obozn = (
            dir_obj.GetDetailsOf(item, meta_obozn_number)
            .replace("$", "")
            .replace("|", "")
            .replace(" ", "")
            .strip()
            .lower()
        )  # читаем обозначение мимо компаса, для увелечения скорости
        return draw_obozn

    def _get_list_of_paths_with_extra_obozn(self) -> list[DoublePathsData]:
        def filter_paths_by_extension(file_extension: str) -> list[FilePath]:
            return list(filter(lambda path: path.endswith(file_extension), paths))

        list_of_double_paths: list[DoublePathsData] = []
        for draw_obozn, paths in self.draws_data_base.items():
            if len(paths) < 2:
                continue
            list_of_double_paths.append(
                DoublePathsData(
                    draw_obozn=draw_obozn,
                    cdw_paths=filter_paths_by_extension("cdw"),
                    spw_paths=filter_paths_by_extension("spw"),
                )
            )
        return list_of_double_paths

    def _proceed_double_paths(self, double_paths_list: list[DoublePathsData]):
        """
        При создании чертежей возможно наличие одинаковых обозначений для деталей
        с разным названием в базе, данный код проверяет наличие таких чертежей.
        """

        def create_output_error_list() -> list[tuple[DrawObozn, FilePath]]:
            # для группировки сообщений при последующей печати
            return [
                (path_data.draw_obozn, path) for path in path_data.cdw_paths + path_data.spw_paths
            ]

        self.status.emit("Открытие Kompas")

        self.calculate_step.emit(len(double_paths_list), False, True)

        temp_path_dict: dict[DrawObozn, list[FilePath]] = {}
        for path_data in double_paths_list:
            self.status.emit(f"Обработка путей для {path_data.draw_obozn}")
            self.increase_step.emit(True)

            correct_paths: list[FilePath] = []
            for draw_paths in [path_data.cdw_paths, path_data.spw_paths]:
                if not draw_paths:
                    continue

                if not self._confirm_same_draw_name_and_obozn(draw_paths):
                    self.errors_dict[ErrorType.FILE_NAMING].extend(create_output_error_list())
                    del self.draws_data_base[path_data.draw_obozn]
                    break
                correct_paths.append(self._get_right_path(draw_paths))
            if correct_paths:
                temp_path_dict[path_data.draw_obozn] = correct_paths

        self.draws_data_base.update(temp_path_dict)

    @staticmethod
    def _confirm_same_draw_name_and_obozn(draw_paths: list[FilePath]) -> bool:
        def clean_file_path(file_path: FilePath):
            return (
                os.path.basename(file_path)
                .split()[0]
                .replace("$", "")
                .replace("|", "")
                .replace(" ", "")
                .replace("-", "")
                .lower()
                .strip()
            )

        file_names = set([clean_file_path(file_path) for file_path in draw_paths])
        if len(file_names) > 1:
            return False
        return True

    def _get_right_path(self, file_paths: list[FilePath]) -> FilePath:
        """
        Сначала сравнивает даты в штампе и выбираем самый поздний.
        Если они равны считывает дату создания файла и выбирает наиболее раннюю версию
        """
        if len(file_paths) < 2:
            return file_paths[0]

        draws_data: list[tuple[FilePath, int, float]] = []
        for path in file_paths:
            try:
                with self._kompas_api.get_draw_stamp(path) as draw_stamp:
                    stamp_time_of_creation = self._get_stamp_time_of_creation(draw_stamp)
            except kompas_api.DocNotOpenedError:
                self.errors_dict[ErrorType.FILE_ERRORS].append(
                    (
                        f"Не удалось открыть файл {path} возможно"
                        f" файл создан в более новой версии или был перемещен"
                    )
                )
                continue
            file_date_of_creation = os.stat(path).st_ctime
            draws_data.append((path, stamp_time_of_creation, file_date_of_creation))
        sorted_paths = sorted(draws_data, key=lambda draw_data: (-draw_data[1], draw_data[2]))
        return sorted_paths[0][0]

    @staticmethod
    def _get_stamp_time_of_creation(draw_stamp) -> int:
        date_in_stamp = draw_stamp.Text(StampCell.CONSTRUCTOR_DATE_CELL).Str
        try:
            date_in_stamp = utils.date_to_seconds(date_in_stamp)
        except Exception:
            date_in_stamp = 0
        return date_in_stamp

from __future__ import annotations

import os
import queue
import time
from pathlib import Path

import pythoncom
from PyQt5.QtCore import pyqtSignal
from PyQt5.QtCore import QThread

from programm import schemas
from programm.errors import ExecutionNotSelectedError
from programm.errors import SpecificationEmptyError
from programm.kompas_api import KompasAPI
from programm.kompas_api import NoExecutionsError
from programm.kompas_api import NotSupportedSpecTypeError
from programm.kompas_api import OboznSearcher
from programm.kompas_api import SpecPathChecker
from programm.schemas import DrawExecution
from programm.schemas import DrawObozn
from programm.schemas import ErrorType
from programm.schemas import EXECUTION_NOT_CHOSEN
from programm.schemas import FilePath
from programm.schemas import SpecSectionData
from programm.schemas import ThreadKompasAPI
from programm.utils import DrawOboznCreation


class SearchPathsThread(QThread):
    # this thread will fill list widget with system paths from spec,
    #  by they obozn_in_specification parameter, using data_base_file
    status = pyqtSignal(str)
    finished = pyqtSignal(dict, list, bool)
    buttons_enable = pyqtSignal(bool)
    errors = pyqtSignal(str)
    choose_spec_execution = pyqtSignal(dict)
    kill_thread = pyqtSignal()

    def __init__(
        self,
        *,
        specification_path: FilePath,
        data_base_file,
        only_one_specification: bool,
        need_to_merge: bool,
        data_queue: queue.Queue,
        kompas_thread_api: ThreadKompasAPI,
        settings_window_data: schemas.SettingsData,
    ):
        self.draw_paths: list[FilePath] = []
        self.errors_dict: dict[ErrorType: schemas.DrawErrorsType] = {
            ErrorType.FILE_ERRORS: [],
            ErrorType.FILE_MISSING: [],
        }
        self.need_to_merge = need_to_merge
        self.specification_path = specification_path
        self.without_sub_assembles = only_one_specification
        self.data_base_file = data_base_file
        self.error_counter = 1
        self.data_queue = data_queue
        self.remove_duplicate_paths = settings_window_data.remove_duplicates

        self.kompas_thread_api = kompas_thread_api
        QThread.__init__(self)

    def run(self):
        self.buttons_enable.emit(False)
        self.status.emit(f"Обработка {os.path.basename(self.specification_path)}")
        pythoncom.CoInitialize()
        self._kompas_api = KompasAPI(self.kompas_thread_api)

        try:
            obozn_in_specification, errors = self._get_obozn_from_base_specification()
        except (
            FileExistsError,
            ExecutionNotSelectedError,
            SpecificationEmptyError,
            NotSupportedSpecTypeError,
        ) as e:
            self.errors.emit(getattr(e, "message", str(e)))
        except NoExecutionsError:
            self.errors.emit(
                f"{os.path.basename(self.specification_path)} - Для групповой спецификации"
                f"не были получены исполнения"
            )
        else:
            self._process_specification(obozn_in_specification)
        if self.remove_duplicate_paths:
            self.draw_paths = list(dict.fromkeys(self.draw_paths))
        self.buttons_enable.emit(True)
        self.finished.emit(self.errors_dict, self.draw_paths, self.need_to_merge)

    def _get_obozn_from_base_specification(self):
        def select_execution(_executions: dict[DrawExecution, int]):
            self.choose_spec_execution.emit(_executions)
            while True:
                time.sleep(0.1)
                try:
                    execution = self.data_queue.get(block=False)
                except queue.Empty:
                    pass
                else:
                    break
            if execution == EXECUTION_NOT_CHOSEN:
                raise ExecutionNotSelectedError(EXECUTION_NOT_CHOSEN)
            return execution

        obozn_searhcer = OboznSearcher(self.specification_path, self._kompas_api)
        column_numbers = None
        if obozn_searhcer.need_to_select_executions():
            column_numbers = select_execution(obozn_searhcer.get_all_spec_executions())

        obozn_in_specification, errors = obozn_searhcer.get_obozn_from_specification(column_numbers)
        self.errors_dict[ErrorType.FILE_ERRORS].extend(errors)

        if not obozn_in_specification:
            raise SpecificationEmptyError(
                f"{os.path.basename(self.specification_path)} - "
                f"Спецификация пуста, как и вся наша жизнь, обновите файл базы чертежей"
            )

        return obozn_in_specification, errors

    def _get_obozn_from_inner_specification(
        self, spw_file_path: FilePath, draw_data: schemas.DrawData
    ) -> tuple[list[SpecSectionData], list[str]] | None:
        try:
            obozn_searcher = OboznSearcher(
                spw_file_path,
                self._kompas_api,
                without_sub_assembles=self.without_sub_assembles,
                spec_obozn=draw_data.draw_obozn,
            )
            response = obozn_searcher.get_obozn_from_specification()
        except (FileExistsError, NotSupportedSpecTypeError) as e:
            self.errors_dict[ErrorType.FILE_ERRORS].append(getattr(e, "message", str(e)))
            return
        except NoExecutionsError:
            self.errors_dict[ErrorType.FILE_ERRORS].append(
                f"{os.path.basename(self.specification_path)} - Для групповой спецификации"
                f"не были получены исполнения, обновите базу чертежей"
            )
            return
        return response

    def _process_specification(self, draws_in_specification: list[SpecSectionData]):
        self.draw_paths.append(self.specification_path)
        self._search_path_recursively(self.specification_path, draws_in_specification)

    # have to put self.obozn_in_specification in because
    # function calls herself with different input data
    def _search_path_recursively(
        self, spec_path: FilePath, obozn_in_specification: list[SpecSectionData]
    ):
        # spec_path берется не None если идет рекурсия
        self.status.emit(f"Обработка {os.path.basename(spec_path)}")
        for section_data in obozn_in_specification:
            if section_data.draw_type in [
                schemas.DrawType.ASSEMBLY_DRAW,
                schemas.DrawType.DETAIL,
            ]:
                self.draw_paths.extend(
                    self._get_cdw_paths_from_specification(section_data, spec_path=spec_path)
                )
            else:  # Specification paths
                self._fill_draw_list_from_specification(section_data, spec_path=spec_path)

    def _get_cdw_paths_from_specification(
        self,
        section_data: SpecSectionData,
        spec_path: FilePath,
    ) -> list[FilePath]:
        draw_paths = []
        for draw_data in section_data.list_draw_data:
            cdw_file_path = self._fetch_draw_path_from_data_base(
                draw_data, file_extension=".cdw", file_path=spec_path
            )
            if not cdw_file_path:
                continue

            if cdw_file_path and cdw_file_path not in draw_paths:
                # одинаковые пути для одной спеки не добавляем
                draw_paths.append(cdw_file_path)
        return draw_paths

    def _fill_draw_list_from_specification(
        self, section_data: SpecSectionData, spec_path: FilePath
    ):
        registered_draws: set[FilePath] = set()
        for draw_data in section_data.list_draw_data:
            spw_file_path = self._fetch_draw_path_from_data_base(
                draw_data, file_extension=".spw", file_path=spec_path
            )
            if not spw_file_path:
                continue

            if Path(spec_path).stem == Path(spw_file_path).stem:
                self.errors_dict[ErrorType.FILE_ERRORS].append(
                    (
                        f"В спецификации {spec_path} имеется"
                        f" рекурсивная вложенность {draw_data.draw_obozn}"
                    )
                )
                continue
            response = self._get_obozn_from_inner_specification(spw_file_path, draw_data)
            if not response:
                continue
            section_spec_data_list, errors = response

            if spw_file_path in registered_draws:
                self._proceed_mirror_spw_file(section_spec_data_list, spw_file_path)
                continue
            else:
                registered_draws.add(spw_file_path)

            if errors:
                self.errors_dict[ErrorType.FILE_ERRORS].extend(errors)
            self.draw_paths.append(spw_file_path)
            self._search_path_recursively(spw_file_path, section_spec_data_list)

    @staticmethod
    def get_correct_draw_path(draw_path: list[FilePath], file_extension: str) -> FilePath:
        """
        Length could be More than one then constructor by mistake give the
        same name to spec file and assembly file.
        For example assembly draw should be XXX-3.06.01.00 СБ but in spec it's XXX-3.06.01.00,
        so it's the same name as spec file
        """
        if len(draw_path) > 1:
            return [file_path for file_path in draw_path if file_path.endswith(file_extension)][0]
        return draw_path[0]

    def _fetch_draw_path_from_data_base(
        self,
        draw_data: schemas.DrawData,
        file_extension: str,
        file_path: FilePath,
    ) -> FilePath | None:
        draw_obozn, draw_name = draw_data.draw_obozn, draw_data.draw_name
        if not draw_name:
            draw_name = schemas.DrawName("")

        try:
            draw_path = self.data_base_file[draw_obozn.lower()]
        except KeyError:
            draw_path = []
            if file_extension == ".spw":
                draw_path = self.try_fetch_spec_path(draw_obozn)
            if file_extension == ".cdw" or not draw_path:
                spec_path = os.path.basename(file_path)
                missing_draw = [
                    draw_obozn.upper(),
                    draw_name.capitalize().replace("\n", " "),
                ]
                missing_draw_info = (spec_path, " - ".join(missing_draw))
                self.errors_dict[ErrorType.FILE_MISSING].append(missing_draw_info)
                return None
        else:
            draw_path = self.get_correct_draw_path(draw_path, file_extension)

        if os.path.exists(draw_path):
            return draw_path
        else:
            if self.error_counter == 1:  # print this message only once
                self.errors_dict[ErrorType.FILE_ERRORS].append(
                    f"Путь {draw_path} является недействительным, обновите базу чертежей"
                )
                self.error_counter += 1
            return None

    def try_fetch_spec_path(self, spec_obozn: DrawObozn) -> FilePath | None:
        def look_for_path_by_obozn(draw_obozn_list: list[DrawObozn]) -> FilePath | None:
            for _draw_obozn in draw_obozn_list:
                if _spec_path := self.data_base_file.get(_draw_obozn):
                    return self.get_correct_draw_path(_spec_path, ".spw")

        def verify_its_correct_spec_path(_spec_path: FilePath, execution: DrawExecution):
            try:
                is_that_correct_spec_path = SpecPathChecker(
                    _spec_path, self._kompas_api, execution
                ).verify_its_correct_spec_path()
            except FileExistsError:
                self.errors_dict[ErrorType.FILE_ERRORS].append(
                    f"Путь {_spec_path} является недействительным, обновите базу чертежей"
                )
            except NoExecutionsError:
                self.errors_dict[ErrorType.FILE_ERRORS].append(
                    (
                        f"{_spec_path} - Для групповой спецификации "
                        f"не были получены исполнения, обновите базу чертежей"
                    )
                )
                return
            else:
                return is_that_correct_spec_path

        draw_obozn_creator = DrawOboznCreation(spec_obozn)
        spec_path = look_for_path_by_obozn(draw_obozn_creator.draw_obozn_list)
        if spec_path is None:
            return None
        if draw_obozn_creator.need_to_verify_path and not verify_its_correct_spec_path(
            spec_path, draw_obozn_creator.execution
        ):
            return None
        return spec_path

    def _proceed_mirror_spw_file(
        self, section_spec_data_list: list[SpecSectionData], spec_path: FilePath
    ):
        unique_paths = self._get_unique_draw_paths(section_spec_data_list, spec_path)
        if not unique_paths:
            return
        cdw_paths, spw_data_list = unique_paths
        self.draw_paths.extend(cdw_paths)
        for file_path, spw_spec_data_list in spw_data_list:
            self.draw_paths.append(file_path)
            self._search_path_recursively(file_path, spw_spec_data_list)

    def _get_unique_draw_paths(
        self,
        section_spec_data_list: list[SpecSectionData],
        spec_path: FilePath,
    ) -> tuple[set, list] | None:
        def set_file_extension():
            if section_spec_data.draw_type == schemas.DrawType.SPEC_DRAW:
                return '.spw'
            return '.cdw'

        unique_cdw_paths = set()
        spw_data_list = []
        for section_spec_data in section_spec_data_list:
            for draw_data in section_spec_data.list_draw_data:
                file_path = self._fetch_draw_path_from_data_base(
                    draw_data, set_file_extension(), spec_path
                )
                if not file_path or file_path in self.draw_paths:
                    continue
                if section_spec_data.draw_type == schemas.DrawType.SPEC_DRAW:
                    response = self._get_obozn_from_inner_specification(
                        file_path,
                        draw_data,
                    )
                    if not response:
                        continue
                    inner_section_spec_data, _ = response
                    spw_data_list.append([file_path, inner_section_spec_data])
                else:
                    unique_cdw_paths.add(file_path)
        if unique_cdw_paths or spw_data_list:
            return unique_cdw_paths, spw_data_list

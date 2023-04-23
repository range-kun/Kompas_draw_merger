# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import queue
import shutil
import sys
from pathlib import Path
from tkinter import filedialog
from typing import Callable

from PyQt5 import QtWidgets

from programm import kompas_api
from programm import schemas
from programm import utils
from programm.errors import FolderNotSelectedError
from programm.kompas_api import CoreKompass
from programm.pyqt_threads import DataBaseThread
from programm.pyqt_threads import FilterThread
from programm.pyqt_threads import MergeThread
from programm.pyqt_threads import SearchPathsThread
from programm.pyqt_windows.main_window import MainWindow
from programm.pyqt_windows.pop_up_windows import Filters
from programm.pyqt_windows.pop_up_windows import RadioButtonsWindow
from programm.schemas import DrawErrorsType
from programm.schemas import DrawExecution
from programm.schemas import DrawObozn
from programm.schemas import ErrorType
from programm.schemas import EXECUTION_NOT_CHOSEN
from programm.schemas import FILE_NOT_CHOSEN_MESSAGE
from programm.schemas import SpecSectionData
from programm.schemas import ThreadKompasAPI
from programm.utils import check_specification
from programm.utils import FilePath

KOMPAS_EXTENSIONS = [".cdw", ".spw"]


class KompasMerger:
    def __init__(self, _kompas_api: CoreKompass):
        self.kompas_api = _kompas_api

        self.mw = MainWindow()
        self._set_mw_callbacks()

        self.search_path: FilePath | None = None
        self.current_progress = 0
        self.progress_step = 0
        self.data_queue: queue.Queue = queue.Queue()
        self.draw_list: list[FilePath] = []

        self.bypassing_folders_inside_previous_status = True
        self.bypassing_sub_assemblies_previous_status = False

        self.merge_thread: MergeThread | None = None
        self.filter_thread: FilterThread | None = None
        self.data_base_thread: DataBaseThread | None = None
        self.search_path_thread: SearchPathsThread | None = None

        self.data_base_file: dict[DrawObozn, list[FilePath]] = {}
        self.specification_path: FilePath | None = None
        self.previous_filters: Filters = Filters()
        self.obozn_in_specification: list[SpecSectionData] = []

    def _set_mw_callbacks(self):
        # setup_look_up_section
        self.mw.choose_folder_button.clicked.connect(self.choose_source_of_draw_folder)
        self.mw.choose_data_base_button.clicked.connect(self.select_data_base_file_path)

        # setup_spec_section
        self.mw.choose_specification_button.clicked.connect(self.choose_specification)
        self.mw.save_data_base_file_button.clicked.connect(self.save_database_to_disk)

        # setup_look_up_parameters_section
        self.mw.search_in_folder_radio_button.clicked.connect(self.choose_search_way)
        self.mw.search_by_spec_radio_button.clicked.connect(self.choose_search_way)

        # setup_upper_list_buttons
        self.mw.clear_draw_list_button.clicked.connect(self.clear_data)
        self.mw.refresh_draw_list_button.clicked.connect(self.refresh_draws_in_list)
        self.mw.save_items_list.clicked.connect(self.copy_files_from_items_list)

        # setup_lower_list_buttons
        self.mw.add_file_to_list_button.clicked.connect(self.add_file_to_list)
        self.mw.add_folder_to_list_button.clicked.connect(self.add_folder_to_list)

        # setup_bottom_section
        self.mw.merge_files_button.clicked.connect(self.check_merge_changes)

    def copy_files_from_items_list(self):
        file_paths = self.mw.list_widget.get_items_text_data()

        dst_path = filedialog.askdirectory(title="Укажите папку для сохранения")
        if not dst_path:
            self.mw.send_error("Папка не была указана")
            return
        for file_path in file_paths:
            try:
                base = os.path.basename(file_path)
                shutil.copyfile(rf"{file_path}", rf"{dst_path}/{base}")
            except PermissionError:
                self.mw.send_error(
                    (
                        "У вас недостаточно прав для копирования "
                        "в указанную папку копирование остановлено"
                    )
                )
                return
            except shutil.SameFileError:
                continue
        self.mw.send_error("Копирование завершено")

    def choose_source_of_draw_folder(self):
        directory_path = filedialog.askdirectory(title="Укажите папку для поиска")
        if not directory_path:
            return

        if self.mw.search_in_folder_radio_button.isChecked():
            self.fill_list_widget_with_paths(search_path=directory_path)
        else:
            self.proceed_database_source_path(directory_path)

    def set_search_path(self, path: FilePath):
        self.search_path = path
        self.mw.source_of_draws_field.setText(path)

    def proceed_folder_draw_list_search(self, folder_path: str):
        if draw_list := self.get_all_draw_paths_in_folder(folder_path):
            self.search_path = FilePath(folder_path)
            self.mw.source_of_draws_field.setText(folder_path)
            return draw_list

    def choose_specification(self):
        spec_path = FilePath(
            filedialog.askopenfilename(
                initialdir="",
                title="Выбор спецификации",
                filetypes=(("spec", "*.spw"),),
            )
        )

        is_spec_path_set = self.set_specification_path(spec_path)
        if not is_spec_path_set:
            return

        self.mw.path_to_spec_field.setText(spec_path)
        if source_path := self.mw.source_of_draws_field.toPlainText():
            if source_path == self.search_path and self.data_base_file:
                self.get_paths_to_specifications()
            else:
                self.proceed_database_source_path(source_path)

    def set_specification_path(self, spec_path: FilePath):
        if not spec_path:
            self.mw.send_error("Файлы спецификации не выбран")
            return False
        try:
            check_specification(spec_path)
        except utils.FileNotSpecError as e:
            self.mw.send_error(getattr(e, "message", str(e)))
            return False
        self.specification_path = spec_path
        return True

    def get_paths_to_specifications(self, need_to_merge=False):
        self.mw.change_list_widget_state(self.mw.list_widget.clear)
        spec_path = self.mw.path_to_spec_field.toPlainText()
        if spec_path != self.specification_path:
            is_spec_path_set = self.set_specification_path(spec_path)
            if not is_spec_path_set:
                return
        self.start_search_paths_thread(need_to_merge)

    def start_search_paths_thread(self, need_to_merge: bool):
        def choose_spec_execution(response: dict[DrawExecution, int]):
            radio_window = RadioButtonsWindow(list(response.keys()))
            radio_window.exec_()
            if not radio_window.radio_state:
                self.data_queue.put(EXECUTION_NOT_CHOSEN)
                return

            column_numbers = response[radio_window.radio_state]
            if column_numbers == kompas_api.WITHOUT_EXECUTION:
                column_numbers = list(response.values())[:-1]
            else:
                column_numbers = [column_numbers]
            self.data_queue.put(column_numbers)

        self.bypassing_sub_assemblies_previous_status = (
            self.mw.bypassing_sub_assemblies_chekcbox.isChecked()
        )

        only_one_specification = not self.mw.bypassing_sub_assemblies_chekcbox.isChecked()
        self.search_path_thread = SearchPathsThread(
            specification_path=self.specification_path,
            data_base_file=self.data_base_file,
            only_one_specification=only_one_specification,
            need_to_merge=need_to_merge,
            data_queue=self.data_queue,
            kompas_thread_api=self.kompas_api.collect_thread_api(ThreadKompasAPI),
            settings_window_data=self.mw.settings_window_data,
        )
        self.search_path_thread.buttons_enable.connect(self.mw.switch_button_group)
        self.search_path_thread.finished.connect(self.handle_search_path_thread_results)
        self.search_path_thread.status.connect(self.mw.status_bar.showMessage)
        self.search_path_thread.errors.connect(self.mw.send_error)
        self.search_path_thread.choose_spec_execution.connect(choose_spec_execution)
        self.search_path_thread.start()

    def handle_search_path_thread_results(
        self,
        errors_info: dict[ErrorType, DrawErrorsType],
        draw_list: list[FilePath],
        need_to_merge: bool,
    ):
        self.mw.status_bar.showMessage("Завершено получение файлов из спецификации")
        self.draw_list = draw_list

        self.print_out_errors(errors_info)
        if not self.draw_list:
            return

        if need_to_merge:
            self.mw.change_list_widget_state(self.mw.list_widget.fill_list, draw_list=draw_list)
            self.check_merge_changes()
            return

        self.calculate_progress_step(len(draw_list), filter_only=True)
        if self.progress_step:
            self.start_filter_thread(self.handle_filter_results, draw_list)
        else:
            self.mw.change_list_widget_state(self.mw.list_widget.fill_list, draw_list=draw_list)

    def print_out_errors(self, errors_info: dict[ErrorType, DrawErrorsType]):
        def save_errors_message_to_txt():
            filename = self.mw.create_save_file_dialog("Сохранить файл", ".", "txt(*.txt)")[0]
            if not filename:
                return

            try:
                with open(filename, "w") as file:
                    file.write(errors_printer.message_for_file)
            except Exception:
                self.mw.send_error("Ошибка записи")

        if not any(errors_info.get(key) for key in errors_info):
            return

        errors_printer = utils.ErrorsPrinter(errors_info)
        title, message_for_window = errors_printer.create_error_message()

        choice = self.mw.create_yes_no_box(
            title,
            message_for_window,
        )
        if choice != QtWidgets.QMessageBox.Yes:
            return
        save_errors_message_to_txt()

    def fill_list_widget_with_paths(self, search_path: str):
        if not search_path:
            self.mw.send_error("Укажите папку с чертежами")
            return

        draw_list = self.proceed_folder_draw_list_search(search_path)
        if not draw_list:
            return

        self.mw.change_list_widget_state(self.mw.list_widget.clear)
        self.calculate_progress_step(len(draw_list), filter_only=True)
        if self.progress_step:
            self.start_filter_thread(self.handle_filter_results, draw_list)
        else:
            self.mw.change_list_widget_state(self.mw.list_widget.fill_list, draw_list=draw_list)

    def refresh_draws_in_list(self, need_to_merge=False):
        search_path = self.mw.source_of_draws_field.toPlainText()
        specification_path = self.mw.path_to_spec_field.toPlainText()

        if self.mw.search_in_folder_radio_button.isChecked():
            self.fill_list_widget_with_paths(search_path)
        else:
            if not search_path:
                self.mw.send_error(
                    "Укажите папку для создания базы чертежей или файл с расширением .json"
                )
                return
            is_spec_path_set = self.set_specification_path(specification_path)
            if not is_spec_path_set:
                return
            self.proceed_database_source_path(search_path, need_to_merge=need_to_merge)

    def is_filters_required(self) -> bool:
        # If refresh button is clicked then all files will be filtered anyway
        # If merger button is clicked then script checks if files been already filtered,
        # and filter settings didn't change
        # If nothing changed script skips this step
        filters = self.mw.settings_window_data.filters
        if filters is None:
            return False
        return (
            filters != self.previous_filters
            or self.search_path != self.mw.source_of_draws_field.toPlainText()
        )

    def start_filter_thread(
        self, callback: Callable, draw_paths: list[FilePath] = None, filter_only=True
    ):
        draw_paths = draw_paths or self.mw.list_widget.get_items_text_data()
        self.previous_filters = self.mw.settings_window_data.filters
        thread_api = self.kompas_api.collect_thread_api(ThreadKompasAPI)

        self.filter_thread = FilterThread(
            draw_paths, self.mw.settings_window_data.filters, thread_api, filter_only
        )
        self.filter_thread.status.connect(self.mw.status_bar.showMessage)
        self.filter_thread.increase_step.connect(self.increase_step)
        self.filter_thread.finished.connect(callback)
        self.filter_thread.switch_button_group.connect(self.mw.switch_button_group)
        self.filter_thread.start()

    def handle_filter_results(
        self, draw_list: list[FilePath], errors_list: list[str], filter_only=True
    ):
        self.mw.status_bar.showMessage("Фильтрация успешно завершена")
        if not draw_list:
            self.mw.send_error(
                "Нету файлов .cdw или .spw, в выбранной папке(ах) с указанными параметрами"
            )
            self.current_progress = 0
            self.mw.progress_bar.setValue(int(self.current_progress))
            self.mw.change_list_widget_state(self.mw.list_widget.clear)
            return

        self.mw.change_list_widget_state(self.mw.list_widget.clear)
        self.mw.change_list_widget_state(self.mw.list_widget.fill_list, draw_list=draw_list)
        self.print_out_errors({ErrorType.FILE_ERRORS: errors_list})
        if filter_only:
            self.current_progress = 0
            self.mw.progress_bar.setValue(int(self.current_progress))
            self.mw.switch_button_group(True)
            return
        else:
            self.start_merge_process(draw_list)

    def get_all_draw_paths_in_folder(self, search_path: str) -> list[FilePath] | None:
        draw_list = []
        except_folders_list = self.mw.settings_window_data.except_folders_list

        self.mw.change_list_widget_state(self.mw.list_widget.clear)
        self.bypassing_folders_inside_previous_status = (
            self.mw.bypassing_folders_inside_checkbox.isChecked()
        )

        if (
            self.mw.bypassing_folders_inside_checkbox.isChecked()
            or self.mw.search_by_spec_radio_button.isChecked()
        ):
            for this_dir, dirs, files_here in os.walk(search_path, topdown=True):
                dirs[:] = [directory for directory in dirs if directory not in except_folders_list]
                if os.path.basename(this_dir) in except_folders_list:
                    continue
                else:
                    files = self.get_files_in_one_folder(this_dir)
                    draw_list += files
        else:
            draw_list = self.get_files_in_one_folder(search_path)
        if draw_list:
            return list(draw_list)
        else:
            self.mw.send_error(
                "Нету файлов .cdw или .spw, в выбранной папке(ах) с указанными параметрами"
            )

    def get_files_in_one_folder(self, folder_path: str) -> list[FilePath]:
        split_paths: list[tuple[str, str]] = [
            os.path.splitext(file_path)
            for file_path in os.listdir(folder_path)
            if os.path.splitext(file_path)[1] in KOMPAS_EXTENSIONS
        ]
        temp_list = sorted(split_paths, key=lambda split_path: split_path[1], reverse=True)
        # sorting in the way so .spw specification
        # sheet is the first if .cdw and .spw files has the same name
        sorted_split_paths = sorted(temp_list, key=lambda split_path: split_path[0])
        sorted_draws = [
            FilePath(os.path.normpath(os.path.join(folder_path, root + ext)))
            for root, ext in sorted_split_paths
        ]
        return sorted_draws

    def check_settings_conditions_before_merge(self) -> bool:
        if not self.mw.source_of_draws_field.toPlainText() and not self.draw_list:
            self.mw.send_error("Укажите исходную папку чертежей для слития файлов")
            raise FolderNotSelectedError
        if self.mw.search_in_folder_radio_button.isChecked() and os.path.isfile(
            self.mw.source_of_draws_field.toPlainText()
        ):
            self.mw.send_error("В качестве источники для чертежей выбран файл а не папка.")
            raise FolderNotSelectedError

        elif self.mw.search_by_spec_radio_button.isChecked() and (
            self.search_path != self.mw.source_of_draws_field.toPlainText()
            or self.specification_path != self.mw.path_to_spec_field.toPlainText()
            or self.mw.bypassing_sub_assemblies_chekcbox.isChecked()
            != self.bypassing_sub_assemblies_previous_status
        ):
            choice = self.mw.create_yes_no_box(
                "Изменения",
                "Настройки поиска файлов и/или спецификация были изменены.Обновить список файлов?",
            )
            if choice == QtWidgets.QMessageBox.Yes:
                self.refresh_draws_in_list(need_to_merge=True)
                return True
        elif self.mw.search_in_folder_radio_button.isChecked() and (
            self.search_path != self.mw.source_of_draws_field.toPlainText()
            or self.mw.bypassing_folders_inside_checkbox.isChecked()
            != self.bypassing_folders_inside_previous_status
        ):
            choice = self.mw.create_yes_no_box(
                "Изменения",
                "Путь или настройки поиска были изменены.Обновить список файлов?",
            )
            if choice == QtWidgets.QMessageBox.Yes:
                self.fill_list_widget_with_paths(self.mw.source_of_draws_field.toPlainText())

    def check_filter_changes_before_merge(self) -> bool:
        draws_list = self.mw.list_widget.get_items_text_data()
        self.calculate_progress_step(len(draws_list))
        self.mw.switch_button_group(False)
        if self.is_filters_required():
            self.start_filter_thread(self.handle_filter_results, draws_list, filter_only=False)
            return True

    def check_merge_changes(self):
        try:
            is_refresh_started = self.check_settings_conditions_before_merge()
        except FolderNotSelectedError:
            return
        if is_refresh_started:
            return
        is_filters_required = self.check_filter_changes_before_merge()
        if (
            not is_filters_required
        ):  # start_merge_process will be called by function handle_filter_results
            self.start_merge_process(self.mw.list_widget.get_items_text_data())

    def start_merge_process(self, draws_list: list[FilePath]):
        def choose_folder(signal):
            dict_for_pdf = filedialog.askdirectory(title="Укажите папку для сохранения")
            self.data_queue.put(dict_for_pdf or FILE_NOT_CHOSEN_MESSAGE)

        def collect_merge_data() -> schemas.MergerData:
            main_name = None
            if self.mw.search_by_spec_radio_button.isChecked():
                main_name = Path(self.specification_path).stem
            return schemas.MergerData(
                self.mw.delete_single_draws_after_merge_checkbox.isChecked(),
                specification_path=main_name,
            )

        self.mw.status_bar.showMessage("Открытие Kompas")

        search_path = (
            self.search_path
            if self.mw.search_in_folder_radio_button.isChecked()
            else os.path.dirname(self.specification_path)
        )
        thread_api = self.kompas_api.collect_thread_api(ThreadKompasAPI)
        self.merge_thread = MergeThread(
            files=draws_list,
            directory=search_path,
            data_queue=self.data_queue,
            kompas_thread_api=thread_api,
            settings_window_data=self.mw.settings_window_data,
            merger_data=collect_merge_data(),
        )
        self.merge_thread.buttons_enable.connect(self.mw.switch_button_group)
        self.merge_thread.increase_step.connect(self.increase_step)
        self.merge_thread.kill_thread.connect(self.stop_merge_thread)
        self.merge_thread.choose_folder.connect(choose_folder)
        self.merge_thread.send_errors.connect(self.mw.send_error)
        self.merge_thread.status.connect(self.mw.status_bar.showMessage)
        self.merge_thread.progress_bar.connect(self.mw.progress_bar.setValue)

        self.merge_thread.start()

    def stop_merge_thread(self):
        self.mw.switch_button_group(True)
        self.mw.status_bar.showMessage("Папка не выбрана, запись прервана")
        self.merge_thread.terminate()

    def clear_data(self):
        self.search_path = None
        self.specification_path = None
        self.mw.source_of_draws_field.clear()
        if self.mw.search_in_folder_radio_button.isChecked():
            self.mw.source_of_draws_field.setPlaceholderText(
                "Выберите папку с файлами в формате .cdw или .spw"
            )
        else:
            self.mw.source_of_draws_field.setPlaceholderText(
                (
                    "Выберите папку с файлами в формате .cdw или .spw"
                    " \n или файл с базой чертежей в формате .json"
                )
            )
        self.mw.path_to_spec_field.clear()
        self.mw.path_to_spec_field.setPlaceholderText("Укажите путь до файла со спецификацией")
        self.mw.change_list_widget_state(self.mw.list_widget.clear)

    def add_file_to_list(self):
        file_path = filedialog.askopenfilename(
            initialdir="",
            title="Выбор чертежей",
            filetypes=(("Спецификация", "*.spw"), ("Чертёж", "*.cdw")),
        )
        if file_path:
            self.mw.change_list_widget_state(self.mw.list_widget.fill_list, draw_list=[file_path])
            self.mw.merge_files_button.setEnabled(True)

    def add_folder_to_list(self):
        directory_path = filedialog.askdirectory(title="Укажите папку с чертежами")
        draw_list = []
        if directory_path:
            draw_list = self.get_files_in_one_folder(directory_path)
        if draw_list:
            self.mw.change_list_widget_state(self.mw.list_widget.fill_list, draw_list=draw_list)
            self.mw.merge_files_button.setEnabled(True)

    def choose_search_way(self):
        self.mw.choose_data_base_button.setEnabled(self.mw.search_by_spec_radio_button.isChecked())
        self.mw.choose_specification_button.setEnabled(
            self.mw.search_by_spec_radio_button.isChecked()
        )
        self.mw.path_to_spec_field.setEnabled(self.mw.search_by_spec_radio_button.isChecked())
        self.mw.bypassing_sub_assemblies_chekcbox.setEnabled(
            self.mw.search_by_spec_radio_button.isChecked()
        )
        self.mw.bypassing_folders_inside_checkbox.setEnabled(
            self.mw.search_in_folder_radio_button.isChecked()
        )
        if self.mw.search_in_folder_radio_button.isChecked():
            self.mw.source_of_draws_field.setPlaceholderText(
                "Выберите папку с файлами в формате .cdw или .spw"
            )
            self.mw.choose_folder_button.setText("Выбор папки \n с чертежами для поиска")

        else:
            self.mw.source_of_draws_field.setPlaceholderText(
                "Выберите папку с файлами в формате "
                ".cdw или .spw \n или файл с базой чертежей в формате .json"
            )
            self.mw.choose_folder_button.setText("Выбор папки для\n создания базы чертежей")

    def calculate_progress_step(self, number_of_files, filter_only=False, get_data_base=False):
        self.current_progress = 0
        self.progress_step = 0
        number_of_operations = 0

        if get_data_base:
            number_of_operations = 1
        if self.mw.settings_window_data.filters:
            if filter_only or self.is_filters_required():
                number_of_operations = 1
        if not filter_only and not get_data_base:
            number_of_operations += 2  # convert files and merger them
        if number_of_operations:
            self.progress_step = 100 / (number_of_operations * number_of_files)

    def increase_step(self, start=True):
        if start:
            self.current_progress += self.progress_step
            self.mw.progress_bar.setValue(int(self.current_progress))

    def proceed_database_source_path(
        self, source_of_draw_path: str | FilePath, need_to_merge=False
    ):
        if os.path.isdir(source_of_draw_path):
            if draw_list := self.proceed_folder_draw_list_search(source_of_draw_path):
                self.get_data_base_from_folder(draw_list, need_to_merge)
        else:
            self.search_paths_by_data_base_file(FilePath(source_of_draw_path), need_to_merge)

    def get_data_base_from_folder(self, draw_paths: list[FilePath], need_to_merge=False):
        self.calculate_progress_step(len(draw_paths), get_data_base=True)
        kompas_thread_api = self.kompas_api.collect_thread_api(ThreadKompasAPI)

        self.data_base_thread = DataBaseThread(draw_paths, need_to_merge, kompas_thread_api)
        self.data_base_thread.buttons_enable.connect(self.mw.switch_button_group)
        self.data_base_thread.calculate_step.connect(self.calculate_progress_step)
        self.data_base_thread.errors.connect(self.mw.send_error)
        self.data_base_thread.status.connect(self.mw.status_bar.showMessage)
        self.data_base_thread.progress_bar.connect(self.mw.progress_bar.setValue)
        self.data_base_thread.increase_step.connect(self.increase_step)
        self.data_base_thread.finished.connect(self.proceed_database_thread_results)

        self.data_base_thread.start()

    def proceed_database_thread_results(
        self,
        data_base: dict[DrawObozn, list[FilePath]],
        errors_info: dict[ErrorType, DrawErrorsType],
        need_to_merge=False,
    ):
        self.mw.progress_bar.setValue(0)
        self.mw.status_bar.showMessage("Завершено получение Базы Чертежей")

        if not data_base:
            self.mw.send_error("Нету файлов с обозначением в штампе")
            return
        self.print_out_errors(errors_info)

        self.data_base_file = data_base
        self.choose_database_storage_method()
        if self.specification_path:
            self.get_paths_to_specifications(need_to_merge)

    def choose_database_storage_method(self):
        choice = self.mw.create_yes_no_box(
            "База Чертежей",
            "Сохранить полученные данные на жёсткий диск?",
        )
        if choice == QtWidgets.QMessageBox.Yes:
            self.save_database_to_disk()
        else:
            self.mw.create_information_box(
                "Отмена записи",
                "База хранится в памяти и будет использована только для текущего запуска",
            )
            self.mw.save_data_base_file_button.setEnabled(True)

    def save_database_to_disk(self):
        data_base_path = self.mw.create_save_file_dialog(
            "Сохранить файл", ".", "Json file(*.json)"
        )[0]
        if not data_base_path:
            self.mw.create_information_box(
                "Отмена записи",
                "База хранится в памяти и будет использована только для текущего запуска",
            )
            return
        try:
            with open(data_base_path, "w") as file:
                json.dump(self.data_base_file, file, ensure_ascii=False)
        except Exception:
            self.mw.send_error("В базе чертежей имеются ошибки")
            return
        self.set_search_path(FilePath(data_base_path))
        if self.mw.save_data_base_file_button.isEnabled():
            self.mw.create_information_box("Запись данных", "Запись данных успешно произведена")
            self.mw.save_data_base_file_button.setEnabled(False)

    def select_data_base_file_path(self):
        file_path = filedialog.askopenfilename(
            initialdir="", title="Загрузить файл", filetypes=(("Json file", "*.json"),)
        )
        if file_path:
            self.search_paths_by_data_base_file(FilePath(file_path))
        else:
            self.mw.send_error("Файл с базой не выбран .json")

    def search_paths_by_data_base_file(self, file_path: FilePath, need_to_merge=False):
        response = self.load_data_base_file(file_path)
        if not response:
            return
        self.set_search_path(file_path)
        if self.specification_path:
            self.get_paths_to_specifications(need_to_merge)

    def load_data_base_file(self, file_path: FilePath):
        if not os.path.exists(file_path):
            self.mw.send_error("Указан несуществующий путь")
            return
        if not file_path.endswith(".json"):
            self.mw.send_error("Указанный файл не является файлом .json")
            return None
        with open(file_path) as file:
            try:
                self.data_base_file = json.load(file)
            except json.decoder.JSONDecodeError:
                self.mw.send_error("В Файл settings.json \n присутствуют ошибки \n синтаксиса json")
                return None
            except UnicodeDecodeError:
                self.mw.send_error(
                    "Указана неверная кодировка файл, попытайтесь еще раз сгенерировать данные"
                )
                return None
            else:
                return 1


def except_hook(cls, exception, traceback):
    # для вывода ошибок в консоль при разработке
    sys.__excepthook__(cls, exception, traceback)

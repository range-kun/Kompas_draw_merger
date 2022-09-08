# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import queue
import shutil
import sys
import time
from collections import defaultdict
from operator import itemgetter
from tkinter import filedialog
from typing import BinaryIO, Callable

import fitz
import pythoncom
import win32com
from PyPDF2 import PdfFileMerger, PdfFileWriter, PdfFileReader
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QThread, pyqtSignal

import kompas_api
import schemas
import utils
from kompas_api import StampCell, DocNotOpened, NoExecutions, NotSupportedSpecType, CoreKompass, Converter, KompasAPI, \
    SpecPathChecker, OboznSearcher
from pop_up_windows import SettingsWindow, RadioButtonsWindow, SaveType, Filters
from schemas import DrawData, DrawType, DrawObozn, SpecSectionData, DrawExecution, ThreadKompasAPI, DoublePathsData, \
    ErrorType
from utils import FilePath, FILE_NOT_EXISTS_MESSAGE, DrawOboznCreation, ErrorsPrinter
from widgets_tools import WidgetBuilder, MainListWidget, WidgetStyles

FILE_NOT_CHOSEN_MESSAGE = "Not_chosen"
EXECUTION_NOT_CHOSEN = "Исполнение не выбрано поиск завершен"


class FolderNotSelected(Exception):
    pass


class ExecutionNotSelected(Exception):
    pass


class SpecificationEmpty(Exception):
    pass


class DifferentDrawsForSameObozn(Exception):
    pass


class NoDraws(Exception):
    pass


class UiMerger(WidgetBuilder):
    def __init__(self, kompas_api: CoreKompass):
        WidgetBuilder.__init__(self, parent=None)
        self.kompas_api = kompas_api
        self.setFixedSize(929, 646)
        self.setWindowTitle("Конвертер")

        self.settings_window = SettingsWindow()
        self.settings_window_data = self.settings_window.collect_settings_window_info()
        self.style_class = WidgetStyles()

        self.setup_ui()

        self.kompas_ext = ['.cdw', '.spw']
        self.search_path: FilePath | None = None
        self.current_progress = 0
        self.progress_step = 0
        self.data_queue: queue.Queue = queue.Queue()
        self.missing_list: list[tuple | str] = []
        self.draw_list: list[FilePath] = []

        self.bypassing_folders_inside_previous_status = True
        self.bypassing_sub_assemblies_previous_status = False

        self.merge_thread: QThread | None = None
        self.filter_thread: QThread | None = None
        self.data_base_thread: QThread | None = None
        self.search_path_thread: QThread | None = None

        self.data_base_file: dict[DrawObozn, list[FilePath]] = {}
        self.specification_path: FilePath | None = None
        self.previous_filters: Filters = Filters()
        self.obozn_in_specification: list[SpecSectionData] = []

    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_Delete:
            self.change_list_widget_state(self.list_widget.remove_selected)

    def setup_ui(self):
        self.setup_styles()
        self.setFont(self.arial_12_font)
        self.grid_widget.setGeometry(QtCore.QRect(10, 10, 906, 611))

        self.setup_look_up_section()
        self.setup_spec_section()
        self.setup_look_up_parameters_section()
        self.setup_upper_list_buttons()
        self.setup_list_widget_section()
        self.setup_lower_list_buttons()
        self.setup_bottom_section()

    def setup_styles(self):
        self.central_widget = QtWidgets.QWidget(self)
        self.grid_widget = QtWidgets.QWidget(self.central_widget)
        self.setCentralWidget(self.central_widget)

        self.grid_layout = QtWidgets.QGridLayout(self.grid_widget)
        self.grid_layout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.grid_layout.setContentsMargins(0, 0, 0, 0)

        self.arial_12_font = self.style_class.arial_12_font
        self.arial_11_font = self.style_class.arial_11_font

        self.line_edit_size_policy = self.style_class.line_edit_size_policy
        self.sizepolicy_button = self.style_class.size_policy_button
        self.sizepolicy_button_2 = self.style_class.size_policy_button_2

    def setup_look_up_section(self):
        self.source_of_draws_field = self.make_text_edit(
            font=self.arial_11_font,
            placeholder="Выберите папку с файлами в формате .spw или .cdw",
            size_policy=self.line_edit_size_policy
        )
        self.grid_layout.addWidget(self.source_of_draws_field, 1, 0, 1, 2)

        self.choose_folder_button = self.make_button(
            text="Выбор папки \n с чертежами для поиска",
            font=self.arial_12_font,
            command=self.choose_source_of_draw_folder
        )
        self.grid_layout.addWidget(self.choose_folder_button, 1, 2, 1, 1)

        self.choose_data_base_button = self.make_button(
            text='Выбор файла\n с базой чертежей',
            font=self.arial_12_font,
            enabled=False,
            command=self.select_data_base_file_path
        )
        self.grid_layout.addWidget(self.choose_data_base_button, 1, 3, 1, 1)

    def setup_spec_section(self):
        self.choose_specification_button = self.make_button(
            text='Выбор \nспецификации',
            font=self.arial_12_font,
            enabled=False,
            command=self.choose_specification
        )
        self.grid_layout.addWidget(self.choose_specification_button, 2, 2, 1, 1)

        self.save_data_base_file_button = self.make_button(
            text='Сохранить \n базу чертежей',
            font=self.arial_12_font, enabled=False,
            command=self.save_database_to_disk
        )
        self.grid_layout.addWidget(self.save_data_base_file_button, 2, 3, 1, 1)

        self.path_to_spec_field = self.make_text_edit(
            font=self.arial_11_font,
            placeholder="Укажите путь до файла со спецификацией .sdw",
            size_policy=self.line_edit_size_policy
        )
        self.path_to_spec_field.setEnabled(False)
        self.grid_layout.addWidget(self.path_to_spec_field, 2, 0, 1, 2)

    def setup_look_up_parameters_section(self):

        self.serch_in_folder_radio_button = self.make_radio_button(
            text='Поиск по папке',
            font=self.arial_12_font,
            command=self.choose_search_way
        )
        self.serch_in_folder_radio_button.setChecked(True)
        self.grid_layout.addWidget(self.serch_in_folder_radio_button, 3, 2, 1, 1)

        self.search_by_spec_radio_button = self.make_radio_button(
            text='Поиск по спецификации',
            font=self.arial_12_font,
            command=self.choose_search_way
        )
        self.grid_layout.addWidget(self.search_by_spec_radio_button, 3, 0, 1, 1)

        self.bypassing_folders_inside_checkbox = self.make_checkbox(
            font=self.arial_12_font,
            text='С обходом всех папок внутри',
            activate=True,
        )
        self.grid_layout.addWidget(self.bypassing_folders_inside_checkbox, 3, 3, 1, 1)

        self.bypassing_sub_assemblies_chekbox = self.make_checkbox(
            font=self.arial_12_font,
            text='С поиском по подсборкам',
        )
        self.bypassing_sub_assemblies_chekbox.setEnabled(False)
        self.grid_layout.addWidget(self.bypassing_sub_assemblies_chekbox, 3, 1, 1, 1)

    def setup_upper_list_buttons(self):
        upper_items_list_layout = QtWidgets.QHBoxLayout()
        self.grid_layout.addLayout(upper_items_list_layout, 4, 0, 1, 4)

        clear_draw_list_button = self.make_button(
            text="Очистить спискок и выбор папки для поиска",
            font=self.arial_12_font,
            command=self.clear_data
        )
        upper_items_list_layout.addWidget(clear_draw_list_button)

        self.refresh_draw_list_button = self.make_button(
            text='Обновить файлы для склеивания',
            font=self.arial_12_font,
            command=self.refresh_draws_in_list
        )
        upper_items_list_layout.addWidget(self.refresh_draw_list_button)

        self.save_items_list = self.make_button(
            text='Скопироваты выбранные файлы',
            font=self.arial_12_font,
            command=self.copy_files_from_items_list
        )
        self.save_items_list.setEnabled(False)
        upper_items_list_layout.addWidget(self.save_items_list)

    def setup_list_widget_section(self):
        horizontal_layout = QtWidgets.QHBoxLayout()
        self.grid_layout.addLayout(horizontal_layout, 8, 0, 1, 4)

        self.list_widget = MainListWidget(self.grid_widget)
        horizontal_layout.addWidget(self.list_widget)

        vertical_layout = QtWidgets.QVBoxLayout()
        horizontal_layout.addLayout(vertical_layout)

        move_line_up_button = self.make_button(
            text='\n\n',
            size_policy=self.sizepolicy_button_2,
            command=self.list_widget.move_item_up
        )
        move_line_up_button.setIcon(QtGui.QIcon('img/arrow_up.png'))
        move_line_up_button.setIconSize(QtCore.QSize(50, 50))

        move_line_down_button = self.make_button(
            text='\n\n',
            size_policy=self.sizepolicy_button_2,
            command=self.list_widget.move_item_down
        )
        move_line_down_button.setIcon(QtGui.QIcon('img/arrow_down.png'))
        move_line_down_button.setIconSize(QtCore.QSize(50, 50))

        delete_list_widget_item = self.make_button(
            text="\n\n",
            size_policy=self.sizepolicy_button_2,
            command=lambda: self.change_list_widget_state(self.list_widget.remove_selected)
        )
        delete_list_widget_item.setIcon(QtGui.QIcon('img/red_cross.png'))
        delete_list_widget_item.setIconSize(QtCore.QSize(50, 50))

        vertical_layout.addWidget(move_line_up_button)
        vertical_layout.addWidget(move_line_down_button)
        vertical_layout.addWidget(delete_list_widget_item)

    def setup_lower_list_buttons(self):
        self.select_all_button = self.make_button(
            text='Выделить все',
            font=self.arial_12_font,
            command=self.list_widget.select_all,
            size_policy=self.sizepolicy_button
        )
        self.grid_layout.addWidget(self.select_all_button, 10, 0, 1, 1)

        self.remove_selection_button = self.make_button(
            text="Снять выделение",
            font=self.arial_12_font, command=self.list_widget.unselect_all,
            size_policy=self.sizepolicy_button
        )
        self.grid_layout.addWidget(self.remove_selection_button, 10, 1, 1, 1)
        self.switch_select_unselect_buttons(False)

        self.add_file_to_list_button = self.make_button(
            text="Добавить файл в список",
            font=self.arial_12_font,
            command=self.add_file_to_list
        )
        self.grid_layout.addWidget(self.add_file_to_list_button, 10, 2, 1, 1)

        self.add_folder_to_list_button = self.make_button(
            text="Добавить папку в список",
            font=self.arial_12_font,
            command=self.add_folder_to_list
        )
        self.grid_layout.addWidget(self.add_folder_to_list_button, 10, 3, 1, 1)

    def setup_bottom_section(self):

        self.additional_settings_button = self.make_button(
            text="Дополнительные настройки",
            font=self.arial_12_font,
            command=self.show_settings
        )
        self.grid_layout.addWidget(self.additional_settings_button, 11, 0, 1, 2)

        self.delete_single_draws_after_merge_checkbox = self.make_checkbox(
            font=self.arial_12_font,
            text='Удалить однодетальные pdf-чертежи по окончанию',
            activate=True
        )
        self.grid_layout.addWidget(self.delete_single_draws_after_merge_checkbox, 11, 2, 1, 2)

        self.merge_files_button = self.make_button(
            text="Склеить файлы",
            font=self.arial_12_font,
            command=self.check_merge_changes

        )
        self.grid_layout.addWidget(self.merge_files_button, 13, 0, 1, 4)

        self.progress_bar = QtWidgets.QProgressBar(self.grid_widget)
        self.progress_bar.setTextVisible(False)
        self.grid_layout.addWidget(self.progress_bar, 15, 0, 1, 4)

        self.status_bar = QtWidgets.QStatusBar(self)
        self.status_bar.setObjectName("statusbar")
        self.setStatusBar(self.status_bar)
        QtCore.QMetaObject.connectSlotsByName(self)

    def copy_files_from_items_list(self):
        file_paths = self.list_widget.get_items_text_data()

        dst_path = filedialog.askdirectory(title="Укажите папку для сохранения")
        if not dst_path:
            self.send_error('Папка не была указана')
            return
        for file_path in file_paths:
            try:
                base = os.path.basename(file_path)
                shutil.copyfile(fr'{file_path}', fr'{dst_path}/{base}')
            except PermissionError:
                self.send_error(
                    'У вас недостаточно прав для копирования в указанную папку копирование остановлено'
                )
                return
            except shutil.SameFileError:
                continue
        self.send_error('Копирование завершено')

    def choose_source_of_draw_folder(self):
        directory_path = filedialog.askdirectory(title="Укажите папку для поиска")
        if not directory_path:
            return

        if self.serch_in_folder_radio_button.isChecked():
            self.fill_list_widget_with_paths(search_path=directory_path)
        else:
            self.proceed_database_source_path(directory_path)

    def set_search_path(self, path: FilePath):
        self.search_path = path
        self.source_of_draws_field.setText(path)

    def proceed_folder_draw_list_search(self, folder_path: str):
        if draw_list := self.get_all_draw_paths_in_folder(folder_path):
            self.search_path = FilePath(folder_path)
            self.source_of_draws_field.setText(folder_path)
            return draw_list

    def choose_specification(self):
        spec_path = FilePath(filedialog.askopenfilename(
            initialdir="",
            title="Выбор cпецификации",
            filetypes=(("spec", "*.spw"),)
        ))

        is_spec_path_set = self.set_specification_path(spec_path)
        if not is_spec_path_set:
            return

        self.path_to_spec_field.setText(spec_path)
        if source_path := self.source_of_draws_field.toPlainText():
            if source_path == self.search_path and self.data_base_file:
                self.get_paths_to_specifications()
            else:
                self.proceed_database_source_path(source_path)

    def set_specification_path(self, spec_path: FilePath):
        if not spec_path:
            self.send_error('Файлы спецификации не выбран')
            return False
        try:
            utils.check_specification(spec_path)
        except utils.FileNotSpec as e:
            self.send_error(getattr(e, 'message', str(e)))
            return False
        self.specification_path = spec_path
        return True

    def get_paths_to_specifications(self, need_to_merge=False):
        self.change_list_widget_state(self.list_widget.clear)
        spec_path = self.path_to_spec_field.toPlainText()
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

        self.bypassing_sub_assemblies_previous_status = self.bypassing_sub_assemblies_chekbox.isChecked()

        only_one_specification = not self.bypassing_sub_assemblies_chekbox.isChecked()
        self.search_path_thread = SearchPathsThread(
            self.specification_path,
            self.data_base_file,
            only_one_specification,
            need_to_merge,
            self.data_queue,
            kompas_thread_api=self.kompas_api.collect_thread_api(ThreadKompasAPI)
        )
        self.search_path_thread.buttons_enable.connect(self.switch_button_group)
        self.search_path_thread.finished.connect(self.handle_search_path_thread_results)
        self.search_path_thread.status.connect(self.status_bar.showMessage)
        self.search_path_thread.errors.connect(self.send_error)
        self.search_path_thread.choose_spec_execution.connect(choose_spec_execution)
        self.search_path_thread.start()

    def handle_search_path_thread_results(
            self,
            missing_list: list[str | tuple],
            draw_list: list[FilePath],
            need_to_merge: bool
    ):
        self.status_bar.showMessage('Завершено получение файлов из спецификации')
        self.missing_list.extend(missing_list)
        self.draw_list = draw_list

        if self.missing_list:
            self.print_out_errors(ErrorType.FILE_MISSING)
        if self.draw_list:
            if need_to_merge:
                self.change_list_widget_state(self.list_widget.fill_list, draw_list=draw_list)
                self.check_merge_changes()
            else:
                self.calculate_progress_step(len(draw_list), filter_only=True)
                if self.progress_step:
                    self.start_filter_thread(self.handle_filter_results, draw_list)
                else:
                    self.change_list_widget_state(self.list_widget.fill_list, draw_list=draw_list)

    def print_out_errors(self, error_type: ErrorType):
        def save_errors_message_to_txt():
            filename = QtWidgets.QFileDialog.getSaveFileName(self, "Сохранить файл", ".", "txt(*.txt)")[0]
            if not filename:
                return
            try:
                with open(filename, 'w') as file:
                    file.write(errors_printer.message_for_file)
            except Exception:
                self.send_error("Ошибка записи")

        errors_printer = ErrorsPrinter(self.missing_list, error_type)
        title, message = errors_printer.create_error_message()
        self.missing_list = []

        choice = QtWidgets.QMessageBox.question(
            self, title, message,
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
        )
        if choice != QtWidgets.QMessageBox.Yes:
            return
        save_errors_message_to_txt()

    def fill_list_widget_with_paths(self, search_path: str):
        if not search_path:
            self.send_error('Укажите папку с чертежамиили')
            return

        draw_list = self.proceed_folder_draw_list_search(search_path)
        if not draw_list:
            return

        self.change_list_widget_state(self.list_widget.clear)
        self.calculate_progress_step(len(draw_list), filter_only=True)
        if self.progress_step:
            self.start_filter_thread(self.handle_filter_results, draw_list)
        else:
            self.change_list_widget_state(self.list_widget.fill_list, draw_list=draw_list)

    def refresh_draws_in_list(self, need_to_merge=False):
        search_path = self.source_of_draws_field.toPlainText()
        specification_path = self.path_to_spec_field.toPlainText()

        if self.serch_in_folder_radio_button.isChecked():
            self.fill_list_widget_with_paths(search_path)
        else:
            if not search_path:
                self.send_error('Укажите папку для создания базы чертежей или файл с расшерением .json')
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
        filters = self.settings_window_data.filters
        if filters is None:
            return False
        return filters != self.previous_filters or self.search_path != self.source_of_draws_field.toPlainText()

    def start_filter_thread(self, callback: Callable, draw_paths: list[FilePath] = None, filter_only=True):
        draw_paths = draw_paths or self.list_widget.get_items_text_data()
        self.previous_filters = self.settings_window_data.filters
        thread_api = self.kompas_api.collect_thread_api(ThreadKompasAPI)

        self.filter_thread = FilterThread(draw_paths, self.settings_window_data.filters, thread_api, filter_only)
        self.filter_thread.status.connect(self.status_bar.showMessage)
        self.filter_thread.increase_step.connect(self.increase_step)
        self.filter_thread.finished.connect(callback)
        self.filter_thread.switch_button_group.connect(self.switch_button_group)
        self.filter_thread.start()

    def handle_filter_results(self, draw_list: list[FilePath], errors_list: list[str], filter_only=True):
        self.status_bar.showMessage('Филтрация успешно завршена')
        if not draw_list:
            self.send_error('Нету файлов .cdw или .spw, в выбранной папке(ах) с указанными параметрами')
            self.current_progress = 0
            self.progress_bar.setValue(int(self.current_progress))
            self.change_list_widget_state(self.list_widget.clear)
            return

        self.change_list_widget_state(self.list_widget.clear)
        self.change_list_widget_state(self.list_widget.fill_list, draw_list=draw_list)
        if errors_list:
            self.missing_list = errors_list
            self.print_out_errors(ErrorType.FILE_NOT_OPENED)
        if filter_only:
            self.current_progress = 0
            self.progress_bar.setValue(int(self.current_progress))
            self.switch_button_group(True)
            return
        else:
            self.start_merge_process(draw_list)

    def get_all_draw_paths_in_folder(self, search_path: str) -> list[FilePath] | None:
        draw_list = []
        except_folders_list = self.settings_window_data.except_folders_list

        self.change_list_widget_state(self.list_widget.clear)
        self.bypassing_folders_inside_previous_status = self.bypassing_folders_inside_checkbox.isChecked()

        if self.bypassing_folders_inside_checkbox.isChecked() \
                or self.search_by_spec_radio_button.isChecked():
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
            self.send_error('Нету файлов .cdw или .spw, в выбраной папке(ах) с указанными параметрами')

    def get_files_in_one_folder(self, folder_path: str) -> list[FilePath]:
        split_paths: list[tuple[str, str]] = \
            [os.path.splitext(file_path) for file_path in os.listdir(folder_path)
             if os.path.splitext(file_path)[1] in self.kompas_ext]
        temp_list = sorted(split_paths, key=lambda split_path: split_path[1], reverse=True)
        # sorting in the way so .spw specification sheet is the first if .cdw and .spw files has the same name
        sorted_split_paths = sorted(temp_list, key=lambda split_path: split_path[0])
        sorted_draws = [
            FilePath(os.path.normpath(os.path.join(folder_path, root + ext))) for root, ext in sorted_split_paths
        ]
        return sorted_draws

    def check_settings_conditions_before_merge(self) -> bool:
        if not self.source_of_draws_field.toPlainText() and not self.draw_list:
            self.send_error('Укажите истоники чертежей для слития файлов')
            raise FolderNotSelected
        if self.serch_in_folder_radio_button.isChecked() and os.path.isfile(self.source_of_draws_field.toPlainText()):
            self.send_error('В качестве источники для чертежей выбран файл а не папка.')
            raise FolderNotSelected

        elif self.search_by_spec_radio_button.isChecked() \
                and (self.search_path != self.source_of_draws_field.toPlainText()
                     or self.specification_path != self.path_to_spec_field.toPlainText()
                     or self.bypassing_sub_assemblies_chekbox.isChecked()
                     != self.bypassing_sub_assemblies_previous_status):
            choice = QtWidgets.QMessageBox.question(
                self,
                "Изменения",
                "Настройки поиска файлов и/или спецификация были изменены.Обновить список файлов?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
            )
            if choice == QtWidgets.QMessageBox.Yes:
                self.refresh_draws_in_list(need_to_merge=True)
                return True
        elif self.serch_in_folder_radio_button.isChecked() and (
                self.search_path != self.source_of_draws_field.toPlainText()
                or self.bypassing_folders_inside_checkbox.isChecked()
                != self.bypassing_folders_inside_previous_status
        ):
            choice = QtWidgets.QMessageBox.question(
                self,
                'Изменения',
                "Путь или настройки поиска были изменены.Обновить список файлов?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
            )
            if choice == QtWidgets.QMessageBox.Yes:
                self.fill_list_widget_with_paths(self.source_of_draws_field.toPlainText())

    def check_filter_changes_before_merge(self) -> bool:
        draws_list = self.list_widget.get_items_text_data()
        self.calculate_progress_step(len(draws_list))
        self.switch_button_group(False)
        if self.is_filters_required():
            self.start_filter_thread(self.handle_filter_results, draws_list, filter_only=False)
            return True

    def check_merge_changes(self):
        try:
            is_refresh_started = self.check_settings_conditions_before_merge()
        except FolderNotSelected:
            return
        if is_refresh_started:
            return
        is_filters_required = self.check_filter_changes_before_merge()
        if not is_filters_required:  # start_merge_process will be called by function handle_filter_results
            self.start_merge_process(self.list_widget.get_items_text_data())

    def start_merge_process(self, draws_list: list[FilePath]):
        def choose_folder(signal):
            dict_for_pdf = filedialog.askdirectory(title="Укажите папку для сохранения")
            self.data_queue.put(dict_for_pdf or FILE_NOT_CHOSEN_MESSAGE)

        self.status_bar.showMessage("Открытие Kompas")

        search_path = self.search_path if self.serch_in_folder_radio_button.isChecked() \
            else os.path.dirname(self.specification_path)
        thread_api = self.kompas_api.collect_thread_api(ThreadKompasAPI)
        self.merge_thread = MergeThread(draws_list, search_path, self.data_queue, thread_api)
        self.merge_thread.buttons_enable.connect(self.switch_button_group)
        self.merge_thread.increase_step.connect(self.increase_step)
        self.merge_thread.kill_thread.connect(self.stop_merge_thread)
        self.merge_thread.choose_folder.connect(choose_folder)
        self.merge_thread.send_errors.connect(self.send_error)
        self.merge_thread.status.connect(self.status_bar.showMessage)
        self.merge_thread.progress_bar.connect(self.progress_bar.setValue)

        self.merge_thread.start()

    def stop_merge_thread(self):
        self.switch_button_group(True)
        self.status_bar.showMessage('Папка не выбрана, запись прервана')
        self.merge_thread.terminate()

    def clear_data(self):
        self.search_path = None
        self.specification_path = None
        self.source_of_draws_field.clear()
        if self.serch_in_folder_radio_button.isChecked():
            self.source_of_draws_field.setPlaceholderText("Выберите папку с файлами в формате .cdw или .spw")
        else:
            self.source_of_draws_field.setPlaceholderText(
                "Выберите папку с файлами в формате .cdw или .spw \n или файл с базой чертежей в формате .json"
            )
        self.path_to_spec_field.clear()
        self.path_to_spec_field.setPlaceholderText('Укажите путь до файла со спецификацией')
        self.change_list_widget_state(self.list_widget.clear)

    def add_file_to_list(self):
        file_path = filedialog.askopenfilename(
            initialdir="",
            title="Выбор чертежей",
            filetypes=(("Спецификация", "*.spw"), ("Чертёж", "*.cdw"))
        )
        if file_path:
            self.change_list_widget_state(self.list_widget.fill_list, draw_list=[file_path])
            self.merge_files_button.setEnabled(True)

    def add_folder_to_list(self):
        directory_path = filedialog.askdirectory(title="Укажите папку с чертежами")
        draw_list = []
        if directory_path:
            draw_list = self.get_files_in_one_folder(directory_path)
        if draw_list:
            self.change_list_widget_state(self.list_widget.fill_list, draw_list=draw_list)
            self.merge_files_button.setEnabled(True)

    def show_settings(self):
        self.settings_window.exec_()
        self.settings_window_data = self.settings_window.collect_settings_window_info()

    def choose_search_way(self):
        self.choose_data_base_button.setEnabled(self.search_by_spec_radio_button.isChecked())
        self.choose_specification_button.setEnabled(self.search_by_spec_radio_button.isChecked())
        self.path_to_spec_field.setEnabled(self.search_by_spec_radio_button.isChecked())
        self.bypassing_sub_assemblies_chekbox.setEnabled(self.search_by_spec_radio_button.isChecked())
        self.bypassing_folders_inside_checkbox.setEnabled(self.serch_in_folder_radio_button.isChecked())
        if self.serch_in_folder_radio_button.isChecked():
            self.source_of_draws_field.setPlaceholderText(
                "Выберите папку с файлами в формате .cdw или .spw"
            )
            self.choose_folder_button.setText("Выбор папки \n с чертежами для поиска")

        else:
            self.source_of_draws_field.setPlaceholderText(
                "Выберите папку с файлами в формате "
                ".cdw или .spw \n или файл с базой чертежей в формате .json"
            )
            self.choose_folder_button.setText("Выбор папки для\n создания базы чертежей")

    def calculate_progress_step(self, number_of_files, filter_only=False, get_data_base=False):
        self.current_progress = 0
        self.progress_step = 0
        number_of_operations = 0

        if get_data_base:
            number_of_operations = 1
        if self.settings_window_data.filters:
            if filter_only or self.is_filters_required():
                number_of_operations = 1
        if not filter_only and not get_data_base:
            number_of_operations += 2  # convert files and merger them
        if number_of_operations:
            self.progress_step = 100 / (number_of_operations * number_of_files)

    def increase_step(self, start=True):
        if start:
            self.current_progress += self.progress_step
            self.progress_bar.setValue(int(self.current_progress))

    def switch_button_group(self, switch=None):
        if not switch:
            switch = False if self.merge_files_button.isEnabled() else True

        self.merge_files_button.setEnabled(switch)
        self.choose_folder_button.setEnabled(switch)
        self.additional_settings_button.setEnabled(switch)
        self.refresh_draw_list_button.setEnabled(switch)
        self.choose_data_base_button.setEnabled(switch)
        self.choose_specification_button.setEnabled(switch)

    def change_list_widget_state(self, method: Callable, *args, **kwargs):
        method(*args, **kwargs)
        self.switch_select_unselect_buttons(self.list_widget.count() > 0)

    def switch_select_unselect_buttons(self, status: bool):
        self.select_all_button.setEnabled(status)
        self.remove_selection_button.setEnabled(status)
        self.save_items_list.setEnabled(status)

    def proceed_database_source_path(self, source_of_draw_path: str | FilePath, need_to_merge=False):
        if os.path.isdir(source_of_draw_path):
            if draw_list := self.proceed_folder_draw_list_search(source_of_draw_path):
                self.get_data_base_from_folder(draw_list, need_to_merge)
        else:
            self.search_paths_by_data_base_file(FilePath(source_of_draw_path), need_to_merge)

    def get_data_base_from_folder(self, draw_paths: list[FilePath], need_to_merge=False):
        self.calculate_progress_step(len(draw_paths), get_data_base=True)
        kompas_thread_api = self.kompas_api.collect_thread_api(ThreadKompasAPI)

        self.data_base_thread = DataBaseThread(draw_paths, need_to_merge, kompas_thread_api)
        self.data_base_thread.buttons_enable.connect(self.switch_button_group)
        self.data_base_thread.calculate_step.connect(self.calculate_progress_step)
        self.data_base_thread.errors.connect(self.send_error)
        self.data_base_thread.status.connect(self.status_bar.showMessage)
        self.data_base_thread.progress_bar.connect(self.progress_bar.setValue)
        self.data_base_thread.increase_step.connect(self.increase_step)
        self.data_base_thread.finished.connect(self.handle_data_base_creation_thread)

        self.data_base_thread.start()

    def handle_data_base_creation_thread(
            self,
            data_base: dict[DrawObozn, list[FilePath]],
            errors_list: list[str],
            need_to_merge=False
    ):
        self.progress_bar.setValue(0)
        self.status_bar.showMessage('Завершено получение Базы Чератежей')

        if not data_base:
            self.send_error('Нету файлов с обозначением в штампе')
            return

        if errors_list:
            self.missing_list = errors_list
            self.print_out_errors(ErrorType.FILE_NOT_OPENED)

        self.data_base_file = data_base
        self.choose_database_storage_method()
        if self.specification_path:
            self.get_paths_to_specifications(need_to_merge)

    def choose_database_storage_method(self):
        choice = QtWidgets.QMessageBox.question(
            self,
            'База Чертежей',
            "Сохранить полученные данные на жесткйи диск?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
        )
        if choice == QtWidgets.QMessageBox.Yes:
            self.save_database_to_disk()
        else:
            QtWidgets.QMessageBox.information(
                self,
                'Отмена записи',
                'База хранится в памяти и будет использована только для текущего запуска'
            )
            self.save_data_base_file_button.setEnabled(True)

    def save_database_to_disk(self):
        data_base_path = QtWidgets.QFileDialog.getSaveFileName(self, "Сохранить файл", ".", "Json file(*.json)")[0]
        if not data_base_path:
            QtWidgets.QMessageBox.information(
                self,
                'Отмена записи',
                'База хранится в памяти и будет использована только для текущего запуска'
            )
            return
        try:
            with open(data_base_path, 'w') as file:
                json.dump(self.data_base_file, file, ensure_ascii=False)
        except:
            self.send_error("В базе чертежей имеются ошибки")
            return
        self.set_search_path(FilePath(data_base_path))
        if self.save_data_base_file_button.isEnabled():
            QtWidgets.QMessageBox.information(
                self,
                'Запись данных',
                'Запись данных успешно произведена'
            )
            self.save_data_base_file_button.setEnabled(False)

    def select_data_base_file_path(self):
        file_path = filedialog.askopenfilename(
            initialdir="",
            title="Загрузить файл",
            filetypes=(("Json file", "*.json"),)
        )
        if file_path:
            self.search_paths_by_data_base_file(FilePath(file_path))
        else:
            self.send_error('Файл с базой не выбран .json')

    def search_paths_by_data_base_file(self, file_path: FilePath, need_to_merge=False):
        response = self.load_data_base_file(file_path)
        if not response:
            return
        self.set_search_path(file_path)
        if self.specification_path:
            self.get_paths_to_specifications(need_to_merge)

    def load_data_base_file(self, file_path: FilePath):
        if not os.path.exists(file_path):
            self.send_error('Указан несуществующий путь')
            return
        if not file_path.endswith('.json'):
            self.send_error('Указанный файл не является файлом .json')
            return None
        with open(file_path) as file:
            try:
                self.data_base_file = json.load(file)
            except json.decoder.JSONDecodeError:
                self.send_error('В Файл settings.txt \n присутсвуют ошибки \n синтаксиса json')
                return None
            except UnicodeDecodeError:
                self.send_error('Указана неверная кодировка файл, попытайтесь еще раз сгенерировать данные')
                return None
            else:
                return 1


class MergeThread(QThread):
    buttons_enable = pyqtSignal(bool)
    send_errors = pyqtSignal(str)
    status = pyqtSignal(str)
    kill_thread = pyqtSignal()
    increase_step = pyqtSignal(bool)
    progress_bar = pyqtSignal(int)
    choose_folder = pyqtSignal(bool)

    def __init__(
            self,
            files: list[FilePath],
            directory: FilePath,
            data_queue: queue.Queue,
            kompas_thread_api: ThreadKompasAPI
    ):
        self._files_path = files
        self._search_path = directory
        self.data_queue = data_queue
        self._kompas_thread_api = kompas_thread_api

        self._constructor_class = QtWidgets.QFileDialog()
        self._settings_window_data = merger.settings_window_data
        self._need_to_split_file = self._settings_window_data.split_file_by_size
        self._need_to_close_files: list[BinaryIO] = []

        QThread.__init__(self)

    def run(self):
        try:
            directory_to_save = self.select_save_folder()
        except FolderNotSelected:
            self._kill_thread()
            return

        self._file_paths_creator = self._initiate_file_paths_creator(directory_to_save)

        single_draw_dir = self._file_paths_creator.single_draw_dir
        os.makedirs(single_draw_dir)

        single_pdf_file_paths = self._file_paths_creator.pdf_file_paths
        self._convert_single_files_to_pdf(single_pdf_file_paths)

        merge_data = self._create_merger_data(single_pdf_file_paths)
        self._merge_pdf_files(merge_data)
        self._close_file_objects()

        if self._settings_window_data.watermark_path:
            self._add_watermark(list(merge_data.keys()))

        if merger.delete_single_draws_after_merge_checkbox.isChecked():
            shutil.rmtree(single_draw_dir)

        if not self._settings_window_data.split_file_by_size:
            pdf_file = self._file_paths_creator.create_main_pdf_file_path()
            os.startfile(pdf_file)

        os.system(f'explorer "{(os.path.normpath(os.path.dirname(single_draw_dir)))}"')

        self.buttons_enable.emit(True)
        self.progress_bar.emit(int(0))
        self.status.emit('Слитие успешно завершено')

    def _kill_thread(self):
        self.send_errors.emit('Запись прервана, папка не была найдена')
        self.buttons_enable.emit(True)
        self.progress_bar.emit(0)
        self.kill_thread.emit()

    def select_save_folder(self) -> FilePath | None:
        # If request folder from this thread later when trying to retrieve kompas api
        # Exception will be raised, that's why, folder is requested from main UiMerger class
        if self._settings_window_data.save_type == SaveType.AUTO_SAVE_FOLDER:
            return None
        self.choose_folder.emit(True)
        while True:
            time.sleep(0.1)
            try:
                directory_to_save = self.data_queue.get(block=False)
            except queue.Empty:
                pass
            else:
                break
        if directory_to_save == FILE_NOT_CHOSEN_MESSAGE:
            raise FolderNotSelected
        return directory_to_save

    def _initiate_file_paths_creator(self, directory_to_save: FilePath | None = None) -> utils.MergerFolderData:
        return utils.MergerFolderData(
            self._search_path,
            self._need_to_split_file,
            self._files_path,
            merger,
            directory_to_save
        )

    def _convert_single_files_to_pdf(self, pdf_file_paths: list[FilePath]):
        pythoncom.CoInitialize()
        _converter = Converter(self._kompas_thread_api)
        for file_path, pdf_file_path in zip(self._files_path, pdf_file_paths):
            self.increase_step.emit(True)
            self.status.emit(f'Конвертация {file_path}')
            _converter.convert_draw_to_pdf(file_path, pdf_file_path)

    @staticmethod
    def _merge_pdf_files(merge_data: dict[FilePath, PdfFileWriter | PdfFileMerger]):
        for pdf_file_path, merger_instance in merge_data.items():
            with open(pdf_file_path, 'wb') as pdf:
                merger_instance.write(pdf)

    def _create_merger_data(self, pdf_file_paths: list[FilePath]) -> dict[FilePath, PdfFileWriter | PdfFileMerger]:

        if self._need_to_split_file:
            merger_instance = self._create_split_merger_data(pdf_file_paths)
        else:
            merger_instance = self._create_single_merger_data(pdf_file_paths)

        return merger_instance

    def _create_split_merger_data(self, pdf_file_paths) -> dict[FilePath, PdfFileWriter]:
        # using different classes because PyPDF2.PdfFileWriter can add single page unlike PdfFileMerger
        merger_instance: defaultdict = defaultdict(PdfFileWriter)

        for pdf_file_path in pdf_file_paths:
            file = self._get_file_obj(pdf_file_path)
            merger_pdf_reader = PdfFileReader(file)

            for page in merger_pdf_reader.pages:
                size = page.mediaBox[2:]
                file_path = self._file_paths_creator.create_main_pdf_file_path(size)
                merger_instance[file_path].addPage(page)

            self.increase_step.emit(True)
            self.status.emit(f'Сливание {pdf_file_path}')

        return merger_instance

    def _create_single_merger_data(self, pdf_file_paths: list[FilePath]) -> dict[FilePath, PdfFileMerger]:
        # using different classes because PyPDF2.PdfFileWriter can add single page unlike PdfFileMerger
        merger_instance = PdfFileMerger()

        for single_pdf_file_path in pdf_file_paths:
            file = self._get_file_obj(single_pdf_file_path)
            merger_instance.append(fileobj=file)

            self.increase_step.emit(True)
            self.status.emit(f'Сливание {single_pdf_file_path}')

        file_path = self._file_paths_creator.create_main_pdf_file_path()
        return {file_path: merger_instance}

    def _get_file_obj(self, file_path: FilePath) -> BinaryIO:
        file = open(file_path, 'rb')  # files will be closed later, using with would lead to blank pdf lists
        self._need_to_close_files.append(file)
        return file

    def _close_file_objects(self):
        for fileobj in self._need_to_close_files:
            fileobj.close()
        self._need_to_close_files = []

    def _add_watermark(self, pdf_file_paths: list[FilePath]):
        def add_watermark_to_file(_pdf_file_path: FilePath):
            pdf_doc = fitz.open(_pdf_file_path)  # open the PDF
            rect = fitz.Rect(watermark_position)  # where to put image: use upper left corner
            for page in pdf_doc:
                if not page.is_wrapped:
                    page.wrap_contents()
                try:
                    page.insert_image(rect, filename=watermark_path, overlay=False)
                except ValueError:
                    self.send_errors.emit(
                        'Заданы неверные координаты, размещения картинки, водяной знак не был добавлен'
                    )
                    return
            pdf_doc.saveIncr()  # do an incremental save

        watermark_path = self._settings_window_data.watermark_path
        watermark_position = self._settings_window_data.watermark_position

        if not watermark_position:
            return
        if not os.path.exists(watermark_path) or watermark_path == FILE_NOT_EXISTS_MESSAGE:
            self.send_errors.emit('Путь к файлу с картинкой не существует')
            return

        for pdf_file_path in pdf_file_paths:
            add_watermark_to_file(pdf_file_path)


class FilterThread(QThread):
    finished = pyqtSignal(list, list, bool)
    increase_step = pyqtSignal(bool)
    status = pyqtSignal(str)
    switch_button_group = pyqtSignal(bool)

    def __init__(self, draw_paths_list, filters: Filters, kompas_thread_api: ThreadKompasAPI, filter_only=True):
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
        self.status.emit(f'Закрытие Kompas')
        self.switch_button_group.emit(True)
        self.finished.emit(filtered_paths_draw_list, self.errors_list, self.filter_only)

    def filter_draws(self) -> list[str]:
        draw_list = []

        self.status.emit("Открытие Kompas")
        for file_path in self.draw_paths_list:  # структура обработки для каждого документа
            self.status.emit(f'Применение фильтров к {file_path}')
            self.increase_step.emit(True)
            try:
                with self._kompas_api.get_draw_stamp(file_path) as draw_stamp:
                    if self.filters.date_range and not self.filter_by_date_cell(draw_stamp):
                        continue

                    file_is_ok = True
                    for data_list, stamp_cell in [
                        (self.filters.constructor_list, StampCell.CONSTRUCTOR_NAME_CELL),
                        (self.filters.checker_list, StampCell.CHECKER_NAME_CELL),
                        (self.filters.sortament_list, StampCell.GAUGE_CELL)
                    ]:
                        if data_list and not self.filter_file_by_cell_value(data_list, stamp_cell, draw_stamp):
                            file_is_ok = False
                            break
            except DocNotOpened:
                self.errors_list.append(
                    f'Не удалось открыть файл {file_path} возможно файл создан в более новой версии или был перемещен\n'
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
            except:
                return False
            if not date_1 <= date_in_stamp <= date_2:
                return False
            return True

    @staticmethod
    def filter_file_by_cell_value(filter_data_list: list[str], stamp_cell_number: StampCell, draw_stamp):
        data_in_stamp = draw_stamp.Text(stamp_cell_number).Str
        if any(filtered_data in data_in_stamp for filtered_data in filter_data_list):
            return True
        return False


class DataBaseThread(QThread):
    increase_step = pyqtSignal(bool)
    status = pyqtSignal(str)
    finished = pyqtSignal(dict, list, bool)
    progress_bar = pyqtSignal(int)
    calculate_step = pyqtSignal(int, bool, bool)
    buttons_enable = pyqtSignal(bool)
    errors = pyqtSignal(str)

    def __init__(self, draw_paths: list[FilePath], need_to_merge: bool, kompas_thread_api: ThreadKompasAPI):
        self.draw_paths = draw_paths
        self.need_to_merge = need_to_merge
        self.kompas_thread_api = kompas_thread_api

        self.draws_data_base: dict[DrawObozn, list[FilePath]] = {}
        self.errors_list: list = []
        QThread.__init__(self)

    def run(self):
        pythoncom.CoInitializeEx(0)
        self.shell = win32com.client.gencache.EnsureDispatch('Shell.Application', 0)  # подлкючаемся к винде

        self.buttons_enable.emit(False)
        obozn_meta_number = self._get_meta_obozn_number(os.path.dirname(self.draw_paths[0]))
        if not obozn_meta_number:
            self.errors.emit("Ошибка при создания базы чертежей")
            return

        self._kompas_api = KompasAPI(self.kompas_thread_api)  # нужно создовать именно в run для правильной работы
        self._create_data_base(obozn_meta_number)

        self.progress_bar.emit(0)
        if double_paths_list := self._get_list_of_paths_with_extra_obozn():
            self._proceed_double_paths(double_paths_list)

        self.progress_bar.emit(0)
        self.buttons_enable.emit(True)
        self.finished.emit(self.draws_data_base, self.errors_list, self.need_to_merge)

    def _get_meta_obozn_number(self, dir_name: str) -> int | None:
        dir_obj = self.shell.NameSpace(dir_name)  # получаем объект папки виндовс шелл
        for number in range(355):
            if dir_obj.GetDetailsOf(None, number) == 'Обозначение':
                return number
        return None

    def _create_data_base(self, meta_obozn_number: int):
        for draw_path in self.draw_paths:
            self.status.emit(f'Получение атрибутов {draw_path}')
            self.increase_step.emit(True)
            if draw_obozn := self._fetch_draw_obozn(draw_path, meta_obozn_number):
                if draw_obozn in self.draws_data_base.keys():
                    self.draws_data_base[draw_obozn].append(draw_path)
                else:
                    self.draws_data_base[draw_obozn] = [draw_path]

    def _fetch_draw_obozn(self, draw_path: FilePath, meta_obozn_number: int) -> DrawObozn:
        dir_obj = self.shell.NameSpace(os.path.dirname(draw_path))  # получаем объект папки виндовс шелл
        item = dir_obj.ParseName(os.path.basename(draw_path))  # указатель на файл (делаем именно объект винд шелл)
        draw_obozn = dir_obj.GetDetailsOf(item, meta_obozn_number).replace('$', '').replace('|', ''). \
            replace(' ', '').strip().lower()  # читаем обозначение мимо компаса, для увелечения скорости
        return draw_obozn

    def _get_list_of_paths_with_extra_obozn(self) -> list[DoublePathsData]:
        def filter_paths_by_extension(file_extension: str) -> list[FilePath]:
            return list(filter(lambda path: path.endswith(file_extension), paths))

        list_of_double_paths: list[DoublePathsData] = []
        for draw_obozn, paths in self.draws_data_base.items():
            if len(paths) < 2:
                continue
            list_of_double_paths.append(DoublePathsData(
                draw_obozn=draw_obozn,
                cdw_paths=filter_paths_by_extension("cdw"),
                spw_paths=filter_paths_by_extension("spw"))
            )
        return list_of_double_paths

    def _proceed_double_paths(self, double_paths_list: list[DoublePathsData]):
        """
            При создании чертежей возможно их дублирование в базе, данный код проверяет наличие таких чертежей
        """
        def create_output_error_list() -> list[tuple[DrawObozn, FilePath]]:
            # для группировки сообщений при последущей печати
            return [(path_data.draw_obozn, path) for path in path_data.cdw_paths + path_data.spw_paths]

        self.status.emit("Открытие Kompas")

        self.calculate_step.emit(len(double_paths_list), False, True)

        for path_data in double_paths_list:
            self.status.emit(f"Обработка путей для {path_data.draw_obozn}")
            self.increase_step.emit(True)

            correct_paths: list[FilePath] = []
            for draw_paths in [path_data.cdw_paths, path_data.spw_paths]:
                if not draw_paths:
                    continue

                if not self._confirm_same_draw_name_and_obozn(draw_paths):
                    self.errors_list.extend(create_output_error_list())
                    del self.draws_data_base[path_data.draw_obozn]
                    break
                correct_paths.append(self._get_right_path(draw_paths))

            self.draws_data_base[path_data.draw_obozn] = correct_paths

    @staticmethod
    def _confirm_same_draw_name_and_obozn(draw_paths: list[FilePath]) -> bool:
        file_names = set([os.path.basename(file_path).replace(" ", '').lower().strip() for file_path in draw_paths])
        if len(file_names) > 1:
            return False
        return True

    def _get_right_path(self, file_paths: list[FilePath]) -> FilePath:
        """
            Сначала сравнивает даты в штампе  и выбираем самый поздний.
            Если они равны считывает дату создания и выбирает наиболее раннюю версию
        """
        if len(file_paths) < 2:
            return file_paths[0]

        draws_data: list[tuple[FilePath, int, float]] = []
        for path in file_paths:
            try:
                with self._kompas_api.get_draw_stamp(path) as draw_stamp:
                    stamp_time_of_creation = self._get_stamp_time_of_creation(draw_stamp)
                    file_date_of_creation = os.stat(path).st_ctime
                    draws_data.append((path, stamp_time_of_creation, file_date_of_creation))
            except DocNotOpened:
                self.errors_list.append(
                    f'Не удалось открыть файл {path} возможно файл создан в более новой версии или был перемещен\n'
                )
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


class SearchPathsThread(QThread):
    # this thread will fill list widget with system paths from spec, by they obozn_in_specification
    # parameter, using data_base_file
    status = pyqtSignal(str)
    finished = pyqtSignal(list, list, bool)
    buttons_enable = pyqtSignal(bool)
    errors = pyqtSignal(str)
    choose_spec_execution = pyqtSignal(dict)
    kill_thread = pyqtSignal()

    def __init__(
            self,
            specification_path: FilePath,
            data_base_file,
            only_one_specification: bool,
            need_to_merge: bool,
            data_queue: queue.Queue,
            kompas_thread_api: ThreadKompasAPI
    ):
        self.draw_paths: list[FilePath] = []
        self.missing_list: list = []
        self.need_to_merge = need_to_merge
        self.specification_path = specification_path
        self.without_sub_assembles = only_one_specification
        self.data_base_file = data_base_file
        self.error = 1
        self.data_queue = data_queue

        self.kompas_thread_api = kompas_thread_api
        QThread.__init__(self)

    def run(self):
        self.buttons_enable.emit(False)
        self.status.emit(f'Обработка {os.path.basename(self.specification_path)}')
        pythoncom.CoInitialize()
        self._kompas_api = KompasAPI(self.kompas_thread_api)
        try:
            obozn_in_specification, errors = self._get_obozn_from_specification()
        except (FileExistsError, ExecutionNotSelected, SpecificationEmpty, NotSupportedSpecType) as e:
            self.errors.emit(getattr(e, 'message', str(e)))
        except NoExecutions:
            self.errors.emit(f"{os.path.basename(self.specification_path)} - Для груповой спецефикации"
                             f"не были получены исполнения")
        else:
            self.process_specification(obozn_in_specification)

        self.buttons_enable.emit(True)
        self.finished.emit(self.missing_list, self.draw_paths, self.need_to_merge)

    def _get_obozn_from_specification(self):
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
                raise ExecutionNotSelected(EXECUTION_NOT_CHOSEN)
            return execution

        obozn_searhcer = OboznSearcher(self.specification_path, self._kompas_api)
        column_numbers = None
        if obozn_searhcer.need_to_select_executions():
            column_numbers = select_execution(obozn_searhcer.get_all_spec_executions())

        obozn_in_specification, errors = obozn_searhcer.get_obozn_from_specification(column_numbers)
        self.missing_list.extend(errors)

        if not obozn_in_specification:
            raise SpecificationEmpty(f'{os.path.basename(self.specification_path)} - '
                                     f'Спецификация пуста, как и вся наша жизнь, обновите файл базы чертежей')

        return obozn_in_specification, errors

    def process_specification(self, draws_in_specification: list[SpecSectionData]):
        self.draw_paths.append(self.specification_path)
        self.recursive_path_searcher(self.specification_path, draws_in_specification)
        # have to put self.obozn_in_specification in because function calls herself with different input data

    def recursive_path_searcher(
            self,
            spec_path: FilePath,
            obozn_in_specification: list[SpecSectionData]
    ):
        # spec_path берется не None если идет рекурсия
        self.status.emit(f'Обработка {os.path.basename(spec_path)}')
        for section_data in obozn_in_specification:
            if section_data.draw_type in [DrawType.ASSEMBLY_DRAW, DrawType.DETAIL]:
                self.draw_paths.extend(self.get_cdw_paths_from_specification(section_data, spec_path=spec_path))
            else:  # Specification paths
                self.fill_draw_list_from_specification(section_data, spec_path=spec_path)

    def get_cdw_paths_from_specification(self, section_data: SpecSectionData, spec_path: FilePath, ) -> list[FilePath]:
        draw_paths = []
        for draw_data in section_data.draw_names:
            draw_file_path = FilePath(self.fetch_draw_path_from_data_base(
                draw_data,
                file_extension=".cdw",
                file_path=spec_path
            ) or "")
            if draw_file_path and draw_file_path not in draw_paths:  # одинаковые пути для одной спеки не добавляем
                draw_paths.append(draw_file_path)
        return draw_paths

    def fill_draw_list_from_specification(self, section_data: SpecSectionData, spec_path: FilePath):
        registered_draws: list[FilePath] = []
        for draw_data in section_data.draw_names:
            spw_file_path = FilePath(self.fetch_draw_path_from_data_base(
                draw_data,
                file_extension='.spw',
                file_path=spec_path
            ) or "")
            if not spw_file_path or spw_file_path in registered_draws:
                continue
            registered_draws.append(spw_file_path)

            try:
                obozn_searhcer = OboznSearcher(
                    spw_file_path,
                    self._kompas_api,
                    only_document_list=self.without_sub_assembles,
                    spec_obozn=draw_data.draw_obozn,
                )
                response = obozn_searhcer.get_obozn_from_specification()
            except (FileExistsError, NotSupportedSpecType) as e:
                self.missing_list.append(getattr(e, 'message', str(e)))
                continue
            except NoExecutions:
                self.missing_list.append(f"\n{os.path.basename(self.specification_path)} - Для груповой спецефикации"
                                         f"не были получены исполнения, обновите базу чертежей")
                continue
            draws_in_specification, errors = response
            if errors:
                self.missing_list.extend(errors)
            self.draw_paths.append(spw_file_path)
            self.recursive_path_searcher(spw_file_path, draws_in_specification)

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

    def fetch_draw_path_from_data_base(
            self,
            draw_data: DrawData,
            file_extension: str,
            file_path: FilePath) -> FilePath | None:
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
                missing_draw = [draw_obozn.upper(), draw_name.capitalize().replace('\n', ' ')]
                missing_draw_info = (spec_path, ' - '.join(missing_draw))
                self.missing_list.append(missing_draw_info)
                return None
        else:
            draw_path = self.get_correct_draw_path(draw_path, file_extension)

        if os.path.exists(draw_path):
            return draw_path
        else:
            if self.error == 1:  # print this message only once
                self.missing_list.append(f"\nПуть {draw_path} является недейстительным, обновите базу чертежей")
                self.error += 1
            return None

    def try_fetch_spec_path(self, spec_obozn: DrawObozn) -> FilePath | None:
        def look_for_path_by_obozn(draw_obozn_list: list[DrawObozn]) -> FilePath | None:
            for _draw_obozn in draw_obozn_list:
                if _spec_path := self.data_base_file.get(_draw_obozn):
                    return self.get_correct_draw_path(_spec_path, ".spw")

        def verify_its_correct_spec_path(_spec_path: FilePath, execution: DrawExecution):
            try:
                is_that_correct_spec_path = SpecPathChecker(
                    _spec_path,
                    self._kompas_api,
                    execution).verify_its_correct_spec_path()
            except FileExistsError:
                self.missing_list.append(f"\nПуть {_spec_path} является недейстительным, обновите базу чертежей")
            except NoExecutions:
                self.missing_list.append(
                    f"\n{_spec_path} - Для групповой спецефикации не были получены исполнения, обновите базу чертежей")
                return
            else:
                return is_that_correct_spec_path

        draw_obozn_creator = DrawOboznCreation(spec_obozn)
        spec_path = look_for_path_by_obozn(draw_obozn_creator.draw_obozn_list)
        if spec_path is None:
            return None
        if draw_obozn_creator.need_to_verify_path \
                and not verify_its_correct_spec_path(spec_path, draw_obozn_creator.execution):
            return None
        return spec_path


def except_hook(cls, exception, traceback):
    # для вывода ошибок в консоль при разработке
    sys.__excepthook__(cls, exception, traceback)


if __name__ == '__main__':
    _core_kompas_api = CoreKompass()
    app = QtWidgets.QApplication(sys.argv)
    app.aboutToQuit.connect(_core_kompas_api.exit_kompas)
    app.setStyle('Fusion')

    merger = UiMerger(_core_kompas_api)
    merger.show()
    sys.excepthook = except_hook
    sys.exit(app.exec_())

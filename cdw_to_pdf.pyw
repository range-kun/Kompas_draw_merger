# -*- coding: utf-8 -*-
from __future__ import annotations

import itertools
import json
import os
import queue
import shutil
import sys
import time
from collections import defaultdict
from enum import Enum
from operator import itemgetter
from tkinter import filedialog
from typing import BinaryIO


import fitz
import pythoncom
import win32com
from PyPDF2 import PdfFileMerger, PdfFileWriter, PdfFileReader
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QThread, pyqtSignal
from win32com.client import Dispatch

import kompas_api
import utils
from kompas_api import StampCell, DocNotOpened, NoExecutions, NotSupportedSpecType, CoreKompass, Converter, KompasAPI, \
    SpecPathChecker, OboznSearcher
from pop_up_windows import SettingsWindow, RadioButtonsWindow, SaveType, Filters
from schemas import DrawData, DrawType, DrawObozn, SpecSectionData, DrawExecution, ThreadKompasAPI
from utils import FilePath, FILE_NOT_EXISTS_MESSAGE, DrawOboznCreation
from widgets_tools import WidgetBuilder, MainListWidget, WidgetStyles

FILE_NOT_CHOSEN_MESSAGE = "Not_chosen"
EXECUTION_NOT_CHOSEN = "Исполнение не выбрано поиск завершен"


class FolderNotSelected(Exception):
    pass


class ExecutionNotSelected(Exception):
    pass


class SpecificationEmpty(Exception):
    pass


class ErrorType(Enum):
    FILE_MISSING = 1
    FILE_NOT_OPENED = 2


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
        self.data_queue = queue.Queue()
        self.missing_list = []
        self.draw_list = []

        self.bypassing_folders_inside_previous_status = True
        self.bypassing_sub_assemblies_previous_status = False

        self.merge_thread: QThread | None = None
        self.filter_thread: QThread | None = None
        self.data_base_thread: QThread | None = None
        self.search_path_thread: QThread | None = None

        self.data_base_file: dict[DrawObozn, [list[FilePath]]] | None = None
        self.specification_path = None
        self.previous_filters = {}
        self.obozn_in_specification: list[SpecSectionData] = []

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
            command=self.choose_initial_folder
        )
        self.grid_layout.addWidget(self.choose_folder_button, 1, 2, 1, 1)

        self.choose_data_base_button = self.make_button(
            text='Выбор файла\n с базой чертежей',
            font=self.arial_12_font,
            enabled=False,
            command=self.get_data_base_path
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
            command=self.apply_data_base_save
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

        save_items_list = self.make_button(
            text='Скопироваты выбранные файлы',
            font=self.arial_12_font,
            command=self.copy_files_from_items_list
        )
        upper_items_list_layout.addWidget(save_items_list)

    def setup_list_widget_section(self):
        horizontal_layout = QtWidgets.QHBoxLayout()
        self.grid_layout.addLayout(horizontal_layout, 8, 0, 1, 4)

        self.list_widget = MainListWidget(self.grid_widget)
        horizontal_layout.addWidget(self.list_widget)

        vertical_layout = QtWidgets.QVBoxLayout()
        horizontal_layout.addLayout(vertical_layout)

        self.move_line_up_button = self.make_button(
            text='\n\n',
            size_policy=self.sizepolicy_button_2,
            command=self.list_widget.move_item_up
        )
        self.move_line_up_button.setIcon(QtGui.QIcon('img/arrow_up.png'))
        self.move_line_up_button.setIconSize(QtCore.QSize(50, 50))

        self.move_line_down_button = self.make_button(
            text='\n\n',
            size_policy=self.sizepolicy_button_2,
            command=self.list_widget.move_item_down
        )
        self.move_line_down_button.setIcon(QtGui.QIcon('img/arrow_down.png'))
        self.move_line_down_button.setIconSize(QtCore.QSize(50, 50))

        vertical_layout.addWidget(self.move_line_up_button)
        vertical_layout.addWidget(self.move_line_down_button)

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
            command=self.merge_files
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

    def choose_initial_folder(self):
        directory_path = filedialog.askdirectory(title="Укажите папку для поиска")
        if not directory_path:
            return

        draw_list = self.check_search_path(directory_path)
        if not draw_list:
            self.send_error(f"В указанной папке отсутсвуют чертежи формата .cdw, .spw")
            return

        self.source_of_draws_field.setText(self.search_path)
        if self.serch_in_folder_radio_button.isChecked():
            self.calculate_step(len(draw_list), filter_only=True)
            if self.progress_step:
                self.apply_filters(draw_list)
            else:
                self.list_widget.fill_list(draw_list=draw_list)
        else:
            self.get_data_base(draw_list)

    def check_search_path(self, search_path):
        if os.path.isfile(search_path):
            if self.search_by_spec_radio_button.isChecked():
                self.check_data_base_file(search_path)
            else:
                self.send_error('Укажите папку для поиска с файламиа, а не файл')
        else:
            draw_list = self.get_all_files_in_folder(search_path)
            if draw_list:
                self.search_path = search_path
                return draw_list
            else:
                return None

    def choose_specification(self):
        spec_path = FilePath(filedialog.askopenfilename(
            initialdir="",
            title="Выбор cпецификации",
            filetypes=(("spec", "*.spw"),)
        ))
        if not spec_path:
            return

        is_spec_path_set = self.set_specification_path(spec_path)
        if not is_spec_path_set:
            return

        self.path_to_spec_field.setText(spec_path)
        cdw_file = self.source_of_draws_field.toPlainText()
        if cdw_file:
            if cdw_file == self.search_path and self.data_base_file:
                self.get_paths_to_specifications()
            else:
                self.get_data_base()

    def set_specification_path(self, spec_path: FilePath):
        try:
            utils.check_specification(spec_path)
        except utils.FileNotSpec as e:
            self.send_error(e)
            return False
        self.specification_path = spec_path
        return True

    def get_paths_to_specifications(self, refresh=False):
        self.list_widget.clear()
        spec_path = self.path_to_spec_field.toPlainText()
        if not spec_path:
            return
        if spec_path != self.specification_path:
            is_spec_path_set = self.set_specification_path(spec_path)
            if not is_spec_path_set:
                return
        self.start_search_paths_thread(refresh)

    def start_search_paths_thread(self, refresh: bool):
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
            refresh,
            self.data_queue,
            kompas_thread_api=self.kompas_api.collect_thread_api(ThreadKompasAPI)
        )
        self.search_path_thread.buttons_enable.connect(self.switch_button_group)
        self.search_path_thread.finished.connect(self.handle_search_path_thread_results)
        self.search_path_thread.status.connect(self.status_bar.showMessage)
        self.search_path_thread.errors.connect(self.send_error)
        self.search_path_thread.choose_spec_execution.connect(choose_spec_execution)
        self.search_path_thread.start()

    def handle_search_path_thread_results(self, missing_list, draw_list, refresh):
        self.status_bar.showMessage('Завершено получение файлов из спецификации')
        self.missing_list.extend(missing_list)
        self.draw_list = draw_list

        if self.missing_list:
            self.print_out_errors(ErrorType.FILE_MISSING)
        if self.draw_list:
            if refresh:
                self.list_widget.fill_list(draw_list=self.draw_list)
                self.start_merge_process(draw_list)
            else:
                self.calculate_step(len(draw_list), filter_only=True)
                if self.progress_step:
                    self.apply_filters(draw_list)
                else:
                    self.list_widget.fill_list(draw_list=self.draw_list)

    def print_out_errors(self, error_type: ErrorType) -> str | None:
        def group_missing_files_info():
            grouped_list = itertools.groupby(grouped_messages, itemgetter(0))
            grouped_list = [key + ':\n' + '\n'.join(['----' + v for k, v in value]) for key, value in grouped_list]
            missing_message = '\n'.join(grouped_list)
            return missing_message

        def create_missing_files_message() -> tuple[str, str]:
            window_title = "Отсуствующие чертежи"
            error_message = f"Не были найдены следующи чертежи:\n{missing_message} {''.join(one_line_messages)}" \
                            f"\nСохранить список?"
            return window_title, error_message

        def create_file_erorrs_message() -> tuple[str, str]:
            window_title = "Ошибки при обработке файлов"
            error_message = f"Были получены следующие ошибки при " \
                            f"обработке файлов:\n{missing_message} {''.join(one_line_messages)}" \
                            f"\nСохранить список?"
            return window_title, error_message

        one_line_messages = []
        grouped_messages = []
        for error in self.missing_list:
            if type(error) == str:
                one_line_messages.append(error)
            else:
                grouped_messages.append(error)
        missing_message = group_missing_files_info()

        if error_type == ErrorType.FILE_MISSING:
            title, message = create_missing_files_message()
        elif error_type == ErrorType.FILE_NOT_OPENED:
            title, message = create_file_erorrs_message()

        choice = QtWidgets.QMessageBox.question(
            self, title, message,
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
        )

        self.missing_list = []
        if choice != QtWidgets.QMessageBox.Yes:
            return
        filename = QtWidgets.QFileDialog.getSaveFileName(self, "Сохранить файл", ".", "txt(*.txt)")[0]
        if not filename:
            return
        try:
            with open(filename, 'w') as file:
                file.write(missing_message)
                file.writelines(one_line_messages)
        except Exception:
            self.send_error("Ошибка записи")
            return

    def refresh_draws_in_list(self, refresh=False):
        search_path = self.source_of_draws_field.toPlainText()
        if self.serch_in_folder_radio_button.isChecked():
            if not search_path:
                self.send_error('Укажите папку с чертежамиили')
                return

            draw_list = self.check_search_path(search_path)
            if not draw_list:
                return

            self.list_widget.clear()
            self.calculate_step(len(draw_list), filter_only=True)
            if self.progress_step:
                self.start_filter_thread(draw_list)
            else:
                self.send_error('Обновление завершено')
                self.list_widget.fill_list(draw_list=draw_list)
        else:
            specification_path = self.path_to_spec_field.toPlainText()
            if not search_path:
                self.send_error('Укажите папку с чертежамиили файл .json')
                return
            if not specification_path:
                self.send_error('Укажите файл спецификации')
                return
            is_spec_path_set = self.set_specification_path(specification_path)
            if not is_spec_path_set:
                return
            self.get_data_base(refresh=refresh)

    def apply_filters(self, draw_list=None, filter_only=True):
        # If refresh button is clicked then all files will be filtered anyway
        # If merger button is clicked then script checks if files been already filtered,
        # and filter settings didn't change
        # If nothing changed script skips this step
        if self.is_filters_required():
            self.start_filter_thread(draw_list, filter_only)

    def is_filters_required(self) -> bool:
        filters = self.settings_window_data.filters
        if filters is None:
            return False
        return filters != self.previous_filters or self.search_path != self.source_of_draws_field.toPlainText()

    def start_filter_thread(self, draw_list=None, filter_only=True):
        draw_list = draw_list or self.list_widget.get_items_text_data()
        self.previous_filters = self.settings_window_data.filters
        thread_api = self.kompas_api.collect_thread_api(ThreadKompasAPI)

        self.filter_thread = FilterThread(draw_list, self.settings_window_data.filters, thread_api, filter_only)
        self.filter_thread.status.connect(self.status_bar.showMessage)
        self.filter_thread.increase_step.connect(self.increase_step)
        self.filter_thread.finished.connect(self.handle_filter_results)
        self.filter_thread.start()

    def handle_filter_results(self, draw_list, errors_list: list[str], filter_only=True):
        self.status_bar.showMessage('Филтрация успешно завршена')
        if not draw_list:
            self.send_error('Нету файлов .cdw или .spw, в выбранной папке(ах) с указанными параметрами')
            self.current_progress = 0
            self.progress_bar.setValue(int(self.current_progress))
            self.list_widget.clear()
            self.switch_button_group(True)
            return

        self.list_widget.clear()
        self.list_widget.fill_list(draw_list=draw_list)
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

    def check_data_base_file(self, filename):
        if not os.path.exists(filename):
            self.send_error('Указан несуществующий путь')
            return
        if not filename.endswith('.json'):
            self.send_error('Указанный файл не является файлом .json')
            return None
        with open(filename) as file:
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

    def get_all_files_in_folder(self, search_path=None):
        search_path = search_path or self.search_path
        draw_list = []
        except_folders_list = self.settings_window_data.except_folders_list

        self.list_widget.clear()
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

    def get_files_in_one_folder(self, folder: FilePath):
        # sorting in the way so specification sheet is the first if .cdw and .spw files has the same name
        draw_list = [os.path.splitext(file_path) for file_path in os.listdir(folder)
                     if os.path.splitext(file_path)[1] in self.kompas_ext]
        draw_list = sorted([(i[0], '.adw' if i[1] == '.spw' else '.cdw') for i in draw_list], key=itemgetter(0, 1))
        draw_list = [os.path.join(folder, (i[0] + '.spw' if i[1] == '.adw' else i[0] + '.cdw')) for i in draw_list]
        draw_list = map(os.path.normpath, draw_list)
        return draw_list

    def merge_files(self):
        if self.serch_in_folder_radio_button.isChecked() and os.path.isfile(self.source_of_draws_field.toPlainText()):
            self.send_error('Укажите папку для сливания')
            return

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
                self.refresh_draws_in_list(refresh=True)
            else:
                self.merge_files_in_one()
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
                self.refresh_draws_in_list()
            self.merge_files_in_one()
        else:
            self.merge_files_in_one()

    def merge_files_in_one(self):
        draws_list = self.list_widget.get_items_text_data()
        if not draws_list:
            self.send_error('Нету файлов для слития')
            return
        self.calculate_step(len(draws_list))
        self.switch_button_group(False)
        if self.is_filters_required():
            self.start_filter_thread(draws_list, filter_only=False)  # at the end this method call this function again
        else:
            self.start_merge_process(draws_list)

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
        self.merge_thread.errors.connect(self.send_error)
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
        self.list_widget.clear()

    def add_file_to_list(self):
        file_path = filedialog.askopenfilename(
            initialdir="",
            title="Выбор чертежей",
            filetypes=(("Спецификация", "*.spw"), ("Чертёж", "*.cdw"))
        )
        if file_path:
            self.list_widget.fill_list(draw_list=[file_path])
            self.merge_files_button.setEnabled(True)

    def add_folder_to_list(self):
        directory_path = filedialog.askdirectory(title="Укажите папку с чертежами")
        draw_list = []
        if directory_path:
            draw_list = self.get_files_in_one_folder(directory_path)
        if draw_list:
            self.list_widget.fill_list(draw_list=draw_list)
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

    def calculate_step(self, number_of_files, filter_only=False, get_data_base=False):
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

    def get_data_base(self, files=None, refresh=False):
        if files:
            self.get_data_base_from_folder(files)
        elif self.source_of_draws_field.toPlainText():
            filename = self.source_of_draws_field.toPlainText()
            if os.path.isdir(filename):
                files = self.check_search_path(filename)
                if files:
                    self.get_data_base_from_folder(files, refresh)
            else:
                self.load_data_base(filename, refresh)
        else:
            self.send_error('Укажите папку для поиска файлов')

    def get_data_base_from_folder(self, files, refresh=False):
        self.calculate_step(len(files), get_data_base=True)
        kompas_thread_api = self.kompas_api.collect_thread_api(ThreadKompasAPI)

        self.data_base_thread = DataBaseThread(files, refresh, kompas_thread_api)
        self.data_base_thread.buttons_enable.connect(self.switch_button_group)
        self.data_base_thread.calculate_step.connect(self.calculate_step)
        self.data_base_thread.status.connect(self.status_bar.showMessage)
        self.data_base_thread.progress_bar.connect(self.progress_bar.setValue)
        self.data_base_thread.increase_step.connect(self.increase_step)
        self.data_base_thread.finished.connect(self.handle_data_base_results)

        self.data_base_thread.start()

    def handle_data_base_results(
            self,
            data_base: dict[str, list[str]],
            errors_list: list[str],
            refresh=False
    ):
        self.progress_bar.setValue(0)
        self.status_bar.showMessage('Завершено получение Базы Чератежей')

        if data_base:
            self.data_base_file = data_base
            self.save_data_base()
            self.get_paths_to_specifications(refresh)
        else:
            self.send_error('Нету файлов с обозначением в штампе')
        if errors_list:
            self.missing_list = errors_list
            self.print_out_errors(ErrorType.FILE_NOT_OPENED)

    def save_data_base(self):
        choice = QtWidgets.QMessageBox.question(
            self,
            'База Чертежей',
            "Сохранить полученные данные?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
        )
        if choice == QtWidgets.QMessageBox.Yes:
            self.apply_data_base_save()
        else:
            QtWidgets.QMessageBox.information(
                self,
                'Отмена записи',
                'Данные о связях хранятся в памяти'
            )
            self.save_data_base_file_button.setEnabled(True)

    def apply_data_base_save(self):
        filename = QtWidgets.QFileDialog.getSaveFileName(self, "Сохранить файл", ".", "Json file(*.json)")[0]
        if filename:
            try:
                with open(filename, 'w') as file:
                    json.dump(self.data_base_file, file, ensure_ascii=False)
            except:
                self.send_error("В базе чертежей имеются ошибки")
                return
            self.source_of_draws_field.setText(filename)
            self.search_path = filename
        if self.save_data_base_file_button.isEnabled():
            QtWidgets.QMessageBox.information(
                self,
                'Запись данных',
                'Запись данных успешно произведена'
            )
            self.save_data_base_file_button.setEnabled(False)

    def get_data_base_path(self):
        file_path = filedialog.askopenfilename(
            initialdir="",
            title="Загрузить файл",
            filetypes=(("Json file", "*.json"),)
        )
        if file_path:
            self.load_data_base(file_path)
        else:
            self.send_error('Файл с базой не выбран .json')

    def load_data_base(self, filename, refresh=False):
        filename = filename or self.search_path
        response = self.check_data_base_file(filename)
        if not response:
            return
        self.search_path = filename
        self.source_of_draws_field.setText(filename)
        self.get_paths_to_specifications(refresh)


class MergeThread(QThread):
    buttons_enable = pyqtSignal(bool)
    errors = pyqtSignal(str)
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

        self._file_paths_creator: utils.MergerFolderData | None = None
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
        self.errors.emit('Запись прервана, папка не была найдена')
        self.buttons_enable.emit(True)
        self.progress_bar.emit(0)
        self.kill_thread.emit()


    def select_save_folder(self) -> FilePath | None:
        # If request folder from this thread later when trying to retrieve kompas api
        # Exception will be raised, that's why, folder is requested from main UiMerger class
        if self._settings_window_data.save_type == SaveType.AUTO_SAVE_FOLDER:
            return
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
        merger_instance = defaultdict(PdfFileWriter)

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
            self.send_errors.emit(f'Путь к файлу с картинкой не существует')
            return

        for pdf_file_path in pdf_file_paths:
            add_watermark_to_file(pdf_file_path)


class FilterThread(QThread):
    finished = pyqtSignal(list, list, bool)
    increase_step = pyqtSignal(bool)
    status = pyqtSignal(str)

    def __init__(self, draw_paths_list, filters: Filters, kompas_thread_api: ThreadKompasAPI, filter_only=True):
        self.draw_paths_list = draw_paths_list
        self.filters = filters
        self.filter_only = filter_only
        self.errors_list = []

        self.kompas_thread_api = kompas_thread_api
        self._kompas_api: KompasAPI = None
        QThread.__init__(self)

    def run(self):
        self._kompas_api = KompasAPI(self.kompas_thread_api)
        filtered_paths_draw_list = self.filter_draws()
        self.status.emit(f'Закрытие Kompas')
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

    def __init__(self, draw_list, refresh, kompas_thread_api: ThreadKompasAPI):
        self.draw_list = draw_list
        self.refresh = refresh
        self.files_dict: dict[str, list[str]] = {}
        self.errors_list: list[str] = []

        self.kompas_thread_api = kompas_thread_api
        self._kompas_api: KompasAPI = None
        QThread.__init__(self)

    def run(self):
        self.buttons_enable.emit(False)
        self._kompas_api = KompasAPI(self.kompas_thread_api)
        self.create_data_base()
        self.progress_bar.emit(0)
        self.check_double_files()
        self.progress_bar.emit(0)
        self.buttons_enable.emit(True)
        self.finished.emit(self.files_dict, self.errors_list, self.refresh)

    def create_data_base(self):
        pythoncom.CoInitializeEx(0)
        shell = win32com.client.gencache.EnsureDispatch('Shell.Application', 0)  # подлкючаемся к винде
        dir_obj = shell.NameSpace(os.path.dirname(self.draw_list[0]))  # получаем объект папки виндовс шелл
        for x in range(355):
            cur_meta = dir_obj.GetDetailsOf(None, x)
            if cur_meta == 'Обозначение':
                meta_obozn = x  # присваиваем номер метаданных
                break

        for file in self.draw_list:
            self.status.emit(f'Получение атрибутов {file}')
            self.increase_step.emit(True)
            dir_obj = shell.NameSpace(os.path.dirname(file))  # получаем объект папки виндовс шелл
            item = dir_obj.ParseName(os.path.basename(file))  # указатель на файл (делаем именно объект винд шелл)
            doc_obozn = dir_obj.GetDetailsOf(item, meta_obozn).replace('$', '').\
                replace('|', '').replace(' ', '').strip().lower()
            # читаем обозначение мимо компаса, для увелечения скорости
            # doc_name = dir_obj.GetDetailsOf(item, meta_name)  # читаем наименование мимо компаса
            if doc_obozn:
                if doc_obozn not in self.files_dict.keys():  # если обозначения нет в списке
                    self.files_dict[doc_obozn] = [file]  # добавляем данные документа в основной словарь
                else:
                    self.files_dict[doc_obozn].append(file)

    def check_double_files(self):
        double_files = [(key, [i for i in v if i.endswith('cdw')], [i for i in v if i.endswith('spw')])
                        for key, v in self.files_dict.items() if len(v) > 1]
        if any(len(i[1]) > 1 or len(i[2]) > 1 for i in double_files):
            self.status.emit(f'Открытие Kompas')
            sum_len = len(double_files)
            self.calculate_step.emit(sum_len, False, True)
            for key, cdw_files, spw_files in double_files:
                temp_files = []
                self.increase_step.emit(True)

                if len(cdw_files) > 1:
                    self.files_dict[key] = [i for i in self.files_dict[key] if i.endswith('spw')]
                    path = self.get_right_path(cdw_files)
                    temp_files = [path]

                if len(spw_files) > 1:
                    self.files_dict[key] = [i for i in self.files_dict[key] if i.endswith('cdw')]
                    path = self.get_right_path(spw_files)
                    temp_files.append(path)
                if key in self.files_dict.keys():
                    self.files_dict[key].extend(temp_files)
                else:
                    self.files_dict[key] = temp_files

    def get_right_path(self, same_double_paths):
        # Сначала сравнивает даты в штампе если они равны
        # Считывает дату создания и выбирает наиболее раннюю версию
        temp_dict = {}
        max_date_in_stamp = 0
        for path in same_double_paths:
            self.status.emit(f'Сравнение даты для {path}')
            try:
                with self._kompas_api.get_draw_stamp(path) as draw_stamp:
                    date_in_stamp = draw_stamp.Text(StampCell.CONSTRUCTOR_DATE_CELL).Str
                    try:
                        date_in_stamp = utils.date_to_seconds(date_in_stamp)
                    except:
                        date_in_stamp = 0
                    if date_in_stamp >= max_date_in_stamp:
                        if date_in_stamp > max_date_in_stamp:
                            max_date_in_stamp = date_in_stamp
                            temp_dict = {}

                        date_of_creation = os.stat(path).st_ctime
                        temp_dict[path] = date_of_creation
            except DocNotOpened:
                self.errors_list.append(
                    f'Не удалось открыть файл {path} возможно файл создан в более новой версии или был перемещен\n'
                )

        sorted_paths = sorted(temp_dict, key=temp_dict.get)
        return sorted_paths[0]


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
            refresh: bool,
            data_queue: queue.Queue,
            kompas_thread_api: ThreadKompasAPI
    ):
        self.draw_paths: list[FilePath] = []
        self.missing_list: list = []
        self.refresh = refresh
        self.specification_path = specification_path
        self.without_sub_assembles = only_one_specification
        self.data_base_file = data_base_file
        self.error = 1
        self.data_queue = data_queue

        self.kompas_thread_api = kompas_thread_api
        self._kompas_api: KompasAPI = None
        QThread.__init__(self)

    def run(self):
        self.buttons_enable.emit(False)
        self.status.emit(f'Обработка {os.path.basename(self.specification_path)}')
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
        self.finished.emit(self.missing_list, self.draw_paths, self.refresh)

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

    def get_cdw_paths_from_specification(self, section_data: SpecSectionData, spec_path: FilePath,) -> list[FilePath]:
        draw_paths = []
        for draw_data in section_data.draw_names:
            draw_file_path = FilePath(self.fetch_draw_path_from_data_base(
                draw_data,
                file_extension=".cdw",
                file_path=spec_path
            ))
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
            ))
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
                return
        else:
            draw_path = self.get_correct_draw_path(draw_path, file_extension)

        if os.path.exists(draw_path):
            return draw_path
        else:
            if self.error == 1:  # print this message only once
                self.missing_list.append(f"\nПуть {draw_path} является недейстительным, обновите базу чертежей")
                self.error += 1
            return

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
        if not spec_path:
            return
        if draw_obozn_creator.need_to_verify_path:
            if not verify_its_correct_spec_path(spec_path, draw_obozn_creator.execution):
                return
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

# -*- coding: utf-8 -*-
from __future__ import annotations

import itertools
import json
import os
import queue
import re
import shutil
import sys
import time
from operator import itemgetter
from typing import Optional
from tkinter import filedialog

import PyPDF2
import fitz
import pythoncom
import win32com
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QThread, pyqtSignal
from win32com.client import Dispatch

import kompas_api
from Widgets_class import MakeWidgets
from settings_window import SettingsWindow


def except_hook(cls, exception, traceback):
    sys.__excepthook__(cls, exception, traceback)


class UiMerger(MakeWidgets):
    def __init__(self):
        MakeWidgets.__init__(self, parent=None)
        self.kompas_ext = ['.cdw', '.spw']
        self.setFixedSize(929, 646)
        self.setup_ui()
        self.search_path = None
        self.current_progress = 0
        self.progress_step = 0
        self.settings_window = SettingsWindow()
        self.filter_thread = None
        self.data_queue = None
        self.missing_list = []
        self.draw_list = []
        self.appl = None
        self.bypassing_folders_inside_checkbox_status = 'Yes'
        self.bypassing_folders_inside_checkbox_current_status = 'Yes'
        self.bypassing_sub_assemblies_chekbox_status = 'No'
        self.bypassing_sub_assemblies_chekbox_current_status = 'No'
        self.thread = None
        self.data_base_thread = None
        self.data_base_files = None
        self.specification_path = None
        self.previous_filters = {}
        self.draws_in_specification = {}
        self.recursive_thread = None

    def setup_ui(self):
        self.setObjectName("Merger")
        font = QtGui.QFont()
        sizepolicy = \
            QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Ignored)
        sizepolicy_button = \
            QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Ignored)
        sizepolicy_button_2 = \
            QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        font.setFamily("Arial")
        font.setPointSize(12)
        self.setFont(font)

        self.centralwidget = QtWidgets.QWidget(self)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(10, 10, 906, 611))
        self.gridLayoutWidget.setObjectName("gridLayoutWidget")
        self.gridLayout = QtWidgets.QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)

        self.select_all_button = self.make_button(
            text='Выделить все',
            font=font, command=self.select_all,
            size_policy=sizepolicy_button
        )
        self.gridLayout.addWidget(self.select_all_button, 10, 0, 1, 1)

        self.remove_selection_button = self.make_button(
            text="Снять выделение",
            font=font, command=self.unselect_all,
            size_policy=sizepolicy_button
        )
        self.gridLayout.addWidget(self.remove_selection_button, 10, 1, 1, 1)

        self.add_file_to_list_button = self.make_button(
            text="Добавить файл в список",
            font=font,
            command=self.add_file_to_list
        )
        self.gridLayout.addWidget(self.add_file_to_list_button, 10, 2, 1, 1)

        self.add_folder_to_list_button = self.make_button(
            text="Добавить папку в список",
            font=font,
            command=self.add_folder_to_list
        )
        self.gridLayout.addWidget(self.add_folder_to_list_button, 10, 3, 1, 1)

        self.merge_files_button = self.make_button(
            text="Склеить файлы",
            font=font,
            command=self.check_lines
        )
        self.gridLayout.addWidget(self.merge_files_button, 13, 0, 1, 4)

        self.choose_folder_button = self.make_button(
            text="Выбор папки \n с чертежами для поиска",
            font=font,
            command=self.choose_initial_folder
        )
        self.gridLayout.addWidget(self.choose_folder_button, 1, 2, 1, 1)

        self.clear_draw_list_button = self.make_button(
            text="Очистить спискок и выбор папки для поиска",
            font=font,
            command=self.clear_data
        )
        self.gridLayout.addWidget(self.clear_draw_list_button, 4, 0, 1, 2)

        self.additional_settings_button = self.make_button(
            text="Дополнительные настройки",
            font=font,
            command=self.show_settings
        )
        self.gridLayout.addWidget(self.additional_settings_button, 11, 0, 1, 2)

        self.refresh_draw_list_button = self.make_button(
            text='Обновить список файлов для склеивания',
            font=font,
            command=self.refresh_draws_in_list
        )
        self.gridLayout.addWidget(self.refresh_draw_list_button, 4, 2, 1, 2)

        self.choose_data_base_button = self.make_button(
            text='Выбор файла\n с базой чертежей',
            font=font, enabled=False,
            command=self.get_data_base_path
         )
        self.gridLayout.addWidget(self.choose_data_base_button, 1, 3, 1, 1)

        self.choose_specification_button = self.make_button(
            text='Выбор \nспецификации',
            font=font,
            enabled=False,
            command=self.choose_specification
        )
        self.gridLayout.addWidget(self.choose_specification_button, 2, 2, 1, 1)

        self.save_data_base_file_button = self.make_button(
            text='Сохранить \n базу чертежей',
            font=font, enabled=False,
            command=self.apply_data_base_save
        )
        self.gridLayout.addWidget(self.save_data_base_file_button, 2, 3, 1, 1)

        self.delete_single_draws_after_merge_checkbox = self.make_checkbox(
            font=font,
            text='Удалить однодетальные pdf-чертежи по окончанию',
            activate=True
        )
        self.gridLayout.addWidget(self.delete_single_draws_after_merge_checkbox, 11, 2, 1, 2)

        self.bypassing_folders_inside_checkbox = self.make_checkbox(
            font=font,
            text='С обходом всех папок внутри',
            activate=True,
            command=self.change_bypassing_folders_inside_checkbox_status
        )
        self.gridLayout.addWidget(self.bypassing_folders_inside_checkbox, 3, 3, 1, 1)

        self.bypassing_sub_assemblies_chekbox = self.make_checkbox(
            font=font,
            text='С поиском по подсборкам',
            command=self.change_bypassing_sub_assemblies_chekbox_status
        )
        self.bypassing_sub_assemblies_chekbox.setEnabled(False)
        self.gridLayout.addWidget(self.bypassing_sub_assemblies_chekbox, 3, 1, 1, 1)

        self.source_of_draws_field = self.make_text_edit(
            font=font,
            placeholder="Выберите папку с файлами в формате .spw или .cdw",
            size_policy=sizepolicy
        )
        self.gridLayout.addWidget(self.source_of_draws_field, 1, 0, 1, 2)

        self.path_to_spec_field = self.make_text_edit(
            font=font,
            placeholder="Укажите путь до файла со спецификацией .sdw",
            size_policy=sizepolicy
        )
        self.path_to_spec_field.setEnabled(False)
        self.gridLayout.addWidget(self.path_to_spec_field, 2, 0, 1, 2)

        self.serch_in_folder_radio_button = self.make_radio_button(
            text='Поиск по папке',
            font=font,
            command=self.choose_search_way
        )
        self.serch_in_folder_radio_button.setChecked(True)
        self.gridLayout.addWidget(self.serch_in_folder_radio_button, 3, 2, 1, 1)

        self.search_by_spec_radio_button = self.make_radio_button(
            text='Поиск по спецификации',
            font=font,
            command=self.choose_search_way
        )
        self.gridLayout.addWidget(self.search_by_spec_radio_button, 3, 0, 1, 1)

        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.gridLayout.addLayout(self.horizontalLayout, 8, 0, 1, 4)

        self.listWidget = QtWidgets.QListWidget(self.gridLayoutWidget)
        self.listWidget.itemDoubleClicked.connect(self.open_item)
        self.horizontalLayout.addWidget(self.listWidget)

        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.horizontalLayout.addLayout(self.verticalLayout)

        self.move_line_up_button = self.make_button(
            text='\n\n',
            size_policy=sizepolicy_button_2,
            command=self.move_item_up
        )
        self.move_line_up_button.setIcon(QtGui.QIcon('img/arrow_up.png'))
        self.move_line_up_button.setIconSize(QtCore.QSize(50, 50))

        self.move_line_down_button = self.make_button(
            text='\n\n', size_policy=sizepolicy_button_2,
            command=self.move_item_down
        )
        self.move_line_down_button.setIcon(QtGui.QIcon('img/arrow_down.png'))
        self.move_line_down_button.setIconSize(QtCore.QSize(50, 50))

        self.verticalLayout.addWidget(self.move_line_up_button)
        self.verticalLayout.addWidget(self.move_line_down_button)

        self.progress_bar = QtWidgets.QProgressBar(self.gridLayoutWidget)
        self.progress_bar.setTextVisible(False)
        self.gridLayout.addWidget(self.progress_bar, 15, 0, 1, 4)

        self.setCentralWidget(self.centralwidget)
        self.status_bar = QtWidgets.QStatusBar(self)
        self.status_bar.setObjectName("statusbar")
        self.setStatusBar(self.status_bar)
        QtCore.QMetaObject.connectSlotsByName(self)

    def choose_initial_folder(self):
        directory = QtWidgets.QFileDialog.getExistingDirectory(self, "Выбрать папку", ".")
        if directory:
            draw_list = self.check_search_path(directory)
            if not draw_list:
                return
            self.source_of_draws_field.setText(self.search_path)
            if self.serch_in_folder_radio_button.isChecked():
                self.calculate_step(len(draw_list), filter_only=True)
                if self.progress_step:
                    self.apply_filters(draw_list)
                else:
                    self.fill_list(draw_list=draw_list)
            else:
                self.get_data_base(draw_list)

    def check_search_path(self, search_path):
        if os.path.isfile(search_path):
            if self.search_by_spec_radio_button.isChecked():
                search_path = self.check_data_base_file(search_path)
            else:
                self.send_error('Укажите папку для поиска с файламиа, а не файл')
        else:
            draw_list = self.get_all_files_in_folder(search_path)
            if draw_list:
                self.search_path = search_path
                return draw_list
            else:
                return None

    def move_item_up(self):
        current_row = self.listWidget.currentRow()
        current_item = self.listWidget.takeItem(current_row)
        self.listWidget.insertItem(current_row - 1, current_item)
        self.listWidget.setCurrentRow(current_row - 1)

    def move_item_down(self):
        current_row = self.listWidget.currentRow()
        current_item = self.listWidget.takeItem(current_row)
        self.listWidget.insertItem(current_row + 1, current_item)
        self.listWidget.setCurrentRow(current_row + 1)

    def choose_specification(self):
        file_path = filedialog.askopenfilename(
            initialdir="",
            title="Выбор cпецификации",
            filetypes=(("spec", "*.spw"),)
        )
        if file_path:
            response = self.check_specification(file_path)
            if not response:
                return
            elif type(response) == str:
                self.send_error(response)
                return
            self.path_to_spec_field.setText(file_path)
            cdw_file = self.source_of_draws_field.toPlainText()
            if cdw_file:
                if cdw_file == self.search_path and self.data_base_files:
                    self.get_paths_to_specifications()
                else:
                    self.get_data_base()

    def check_specification(self, file_path: str):
        if file_path.endswith('.spw') and os.path.isfile(file_path):
            response = kompas_api.get_draws_from_specification(file_path)
            if type(response) == str:  # ошибка при открытии спецификации
                return response

            draws_in_specification, self.appl, errors = response
            self.missing_list.extend(errors)
            if draws_in_specification:
                self.draws_in_specification = draws_in_specification
                self.specification_path = file_path
            else:
                self.send_error('Спецификация пуста, как и вся наша жизнь')
                return
        else:
            self.send_error('Указанный файл не является спецификацией или не существует')
            return
        return 1

    def get_paths_to_specifications(self, refresh=False):
        self.listWidget.clear()
        filename = self.path_to_spec_field.toPlainText()
        if not filename:
            return
        if filename != self.specification_path:
            response = self.check_specification(filename)
            if not response:
                return
            if type(response) == str:
                self.send_error(response)  # ошибка при открытии спецификации
        self.start_recursion(refresh)

    def handle_specification_result(self, missing_list, draw_list, refresh):
        self.status_bar.showMessage('Завершено получение файлов из спецификации')
        self.appl = None
        self.missing_list.extend(missing_list)
        self.draw_list = draw_list
        if self.missing_list:
            self.print_out_missing_files()
        if self.draw_list:
            if refresh:
                self.fill_list(draw_list=self.draw_list)
                self.start_merge_process(draw_list)
            else:
                self.calculate_step(len(draw_list), filter_only=True)
                if self.progress_step:
                    self.apply_filters(draw_list)
                else:
                    self.fill_list(draw_list=self.draw_list)
                    self.switch_button_group(True)

    def start_recursion(self, refresh):
        only_one_specification = not self.bypassing_sub_assemblies_chekbox.isChecked()
        self.bypassing_sub_assemblies_chekbox_status = 'Yes' \
            if self.bypassing_sub_assemblies_chekbox.isChecked() else 'No'
        self.recursive_thread = RecursionThread(self.specification_path, self.draws_in_specification,
                                                self.data_base_files, only_one_specification, refresh)
        self.recursive_thread.buttons_enable.connect(self.switch_button_group)
        self.recursive_thread.finished.connect(self.handle_specification_result)
        self.recursive_thread.status.connect(self.status_bar.showMessage)
        self.recursive_thread.errors.connect(self.send_error)
        self.recursive_thread.start()

    def change_bypassing_sub_assemblies_chekbox_status(self):
        if self.bypassing_sub_assemblies_chekbox.isChecked():
            self.bypassing_sub_assemblies_chekbox_current_status = 'Yes'
        else:
            self.bypassing_sub_assemblies_chekbox_current_status = 'No'

    def change_bypassing_folders_inside_checkbox_status(self):
        if self.bypassing_folders_inside_checkbox.isChecked():
            self.bypassing_folders_inside_checkbox_current_status = 'Yes'
        else:
            self.bypassing_folders_inside_checkbox_current_status = 'No'

    def print_out_missing_files(self):
        one_line_messages = []
        grouped_messages = []
        for error in self.missing_list:
            if type(error) == str:
                one_line_messages.append(error)
            else:
                grouped_messages.append(error)

        grouped_list = itertools.groupby(grouped_messages, itemgetter(0))
        grouped_list = [key + ':\n' + '\n'.join(['----' + v for k, v in value]) for key, value in grouped_list]
        missing_message = '\n'.join(grouped_list)
        choice = QtWidgets.QMessageBox.question(
            self, 'Отсуствующие чертежи',
            f"Не были найдены следующи чертежи:\n{missing_message}, {''.join(one_line_messages)}"
            f"\nСохранить список?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
        )
        self.missing_list = []
        if choice == QtWidgets.QMessageBox.Yes:
            filename = QtWidgets.QFileDialog.getSaveFileName(self, "Сохранить файл", ".", "txt(*.txt)")[0]
            if filename:
                try:
                    with open(filename, 'w') as file:
                        file.write(missing_message)
                        file.writelines(one_line_messages)
                except:
                    self.send_error("Ошибка записи")
                    return

    def refresh_draws_in_list(self, refresh=False):
        if self.serch_in_folder_radio_button.isChecked():
            search_path = self.source_of_draws_field.toPlainText()
            if not search_path:
                self.send_error('Укажите папку с чертежамиили')
                return
            draw_list = self.check_search_path(search_path)
            if not draw_list:
                return
            self.listWidget.clear()
            self.calculate_step(len(draw_list), filter_only=True)
            if self.progress_step:
                self.apply_filters(draw_list)
            else:
                self.send_error('Обновление завершено')
                self.fill_list(draw_list=draw_list)
        else:
            search_path = self.source_of_draws_field.toPlainText()
            specification = self.path_to_spec_field.toPlainText()
            if not search_path:
                self.send_error('Укажите папку с чертежамиили файл .json')
                return
            if not specification:
                self.send_error('Укажите файл спецификации')
                return
            response_1 = self.check_specification(specification)
            if not response_1:
                return
            if type(response_1) == str:
                self.send_error(response_1)
                return
            self.get_data_base(refresh=refresh)

    def apply_filters(self, draw_list=None, filter_only=True):
        # If refresh button is clicked then all files will be filtered anyway
        # If merger button is clicked then script checks if files been already filtered,
        # and settings didn't change
        # If so script skips this step
        draw_list = draw_list or self.get_items_in_list(self.listWidget)
        filters = self.get_all_filters()
        if not filter_only:
            if filters == self.previous_filters and self.search_path == self.source_of_draws_field.toPlainText():
                return 1
        self.previous_filters = filters
        self.filter_thread = FilterThread(draw_list, filters, filter_only)
        self.filter_thread.status.connect(self.status_bar.showMessage)
        self.filter_thread.increase_step.connect(self.increase_step)
        self.filter_thread.finished.connect(self.handle_filter_results)
        self.filter_thread.start()

    def handle_filter_results(self, draw_list, filter_only=True):
        self.status_bar.showMessage('Филтрация успешно завршена')
        self.appl = None
        if not draw_list:
            self.send_error('Нету файлов .cdw или .spw, в выбранной папке(ах) с указанными параметрами')
            self.current_progress = 0
            self.progress_bar.setValue(self.current_progress)
            self.listWidget.clear()
            self.switch_button_group(True)
            return
        if filter_only:
            self.current_progress = 0
            self.progress_bar.setValue(self.current_progress)
            self.fill_list(draw_list=draw_list)
            self.switch_button_group(True)
            return
        else:
            self.start_merge_process(draw_list)

    def closeEvent(self, event):
        if self.appl:
            self.status.emit(f'Закрытие Kompas')
            kompas_api.exit_kompas(self.appl)
        event.accept()

    def handle_data_base_results(self, data_base, appl, refresh=False):
        self.status_bar.showMessage('Завершено получение Базы Данных')
        if appl:
            self.appl = None
        if data_base:
            self.data_base_files = data_base
            self.save_data_base()
            self.get_paths_to_specifications(refresh)
        else:
            self.send_error('Нету файлов с обозначением в штампе')

    def save_data_base(self):
        choice = QtWidgets.QMessageBox.question(
            self,
            'База данных',
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
                    json.dump(self.data_base_files, file, ensure_ascii=False)
            except:
                self.send_error("В базе данных имеются ошибки")
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
        filename = QtWidgets.QFileDialog.getOpenFileName(self, "Загрузить файл", ".", "Json file(*.json)")[0]
        if filename:
            if filename:
                self.load_data_base(filename)
            else:
                self.send_error('Указанный файл не является файлом .json')

    def load_data_base(self, filename, refresh=False):
        filename = filename or self.search_path
        response = self.check_data_base_file(filename)
        if not response:
            return
        self.search_path = filename
        self.source_of_draws_field.setText(filename)
        self.get_paths_to_specifications(refresh)

    def check_data_base_file(self, filename):
        if not os.path.exists(filename):
            self.send_error('Указан несуществующий путь')
            return
        if not filename.endswith('.json'):
            self.send_error('Указанный файл не является файлом .json')
            return None
        with open(filename) as file:
            try:
                self.data_base_files = json.load(file)
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
        except_folders_list = self.get_items_in_list(self.settings_window.listWidget)
        self.listWidget.clear()
        self.bypassing_folders_inside_checkbox_status = 'Yes' \
            if self.bypassing_folders_inside_checkbox.isChecked() else 'No'
        if self.bypassing_folders_inside_checkbox.isChecked() \
                or self.search_by_spec_radio_button.isChecked():
            for this_dir, dirs, files_here in os.walk(search_path, topdown=True):
                dirs[:] = [d for d in dirs if d not in except_folders_list]
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

    def get_files_in_one_folder(self, folder):
        # sorting in the way so specification sheet is the first if .cdw and .spw files has the same name
        draw_list = [os.path.splitext(i) for i in os.listdir(folder) if os.path.splitext(i)[1] in self.kompas_ext]
        draw_list = sorted([(i[0], '.adw' if i[1] == '.spw' else '.cdw') for i in draw_list], key=itemgetter(0, 1))
        draw_list = [os.path.join(folder, (i[0] + '.spw' if i[1] == '.adw' else i[0] + '.cdw')) for i in draw_list]
        draw_list = map(os.path.normpath, draw_list)
        return draw_list

    def check_lines(self):
        if self.serch_in_folder_radio_button.isChecked() and os.path.isfile(self.source_of_draws_field.toPlainText()):
            self.send_error('Укажите папку для сливания')
            return
        elif self.search_by_spec_radio_button.isChecked() \
                and (self.search_path != self.source_of_draws_field.toPlainText()
                     or self.specification_path != self.path_to_spec_field.toPlainText()
                     or self.bypassing_sub_assemblies_chekbox_current_status
                     != self.bypassing_sub_assemblies_chekbox_status):
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
                or self.bypassing_folders_inside_checkbox_current_status
                != self.bypassing_folders_inside_checkbox_status
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
        else:
            self.merge_files_in_one()

    def merge_files_in_one(self):
        draws_list = self.get_items_in_list(self.listWidget)
        if not draws_list:
            self.send_error('Нету файлов для слития')
            return
        self.calculate_step(len(draws_list))
        self.switch_button_group(False)
        reply = self.apply_filters(draws_list, False)
        if reply:
            self.start_merge_process(draws_list)

    def start_merge_process(self, draws_list):
        self.data_queue = queue.Queue()
        search_path = self.search_path if self.serch_in_folder_radio_button.isChecked() \
            else os.path.dirname(self.specification_path)
        self.thread = MyBrandThread(draws_list, search_path, self.data_queue)
        self.thread.buttons_enable.connect(self.switch_button_group)
        self.thread.increase_step.connect(self.increase_step)
        self.thread.kill_thread.connect(self.stop_merge_thread)
        self.thread.errors.connect(self.send_error)
        self.thread.choose_folder.connect(self.choose_folder)
        self.thread.status.connect(self.status_bar.showMessage)
        self.thread.progress_bar.connect(self.progress_bar.setValue)
        self.appl = None
        self.thread.start()

    def stop_merge_thread(self):
        self.switch_button_group(True)
        self.status_bar.showMessage('Папка не выбрана, запись прервана')
        self.thread.terminate()

    def clear_data(self):
        self.search_path = None
        self.specification_path = None
        self.source_of_draws_field.clear()
        if self.serch_in_folder_radio_button.isChecked():
            self.source_of_draws_field.setPlaceholderText("Выберите папку с файлами в формате .cdw или .spw")
        else:
            self.line_edit.setPlaceholderText("Выберите папку с файлами в формате "
                                              ".cdw или .spw \n или файл с базой данных в формате .json")
        self.path_to_spec_field.clear()
        self.path_to_spec_field.setPlaceholderText('Укажите путь до файла со спецификацией')
        self.listWidget.clear()

    def select_all(self):
        files = (self.listWidget.item(i) for i
                 in range(self.listWidget.count()) if not self.listWidget.item(i).checkState())
        if files:
            for file in files:
                file.setCheckState(QtCore.Qt.Checked)

    def unselect_all(self):
        files = (self.listWidget.item(i) for i
                 in range(self.listWidget.count()) if self.listWidget.item(i).checkState())
        if files:
            for file in files:
                file.setCheckState(QtCore.Qt.Unchecked)

    def add_file_to_list(self):
        filename = [QtWidgets.QFileDialog.getOpenFileName(
            self, "Выбрать файл", ".", "Чертж(*.cdw);;Спецификация(*.spw)")[0]]
        if filename[0]:
            self.fill_list(draw_list=filename)
            self.merge_files_button.setEnabled(True)

    def add_folder_to_list(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Выбрать папку", ".")
        draw_list = []
        if folder:
            draw_list = self.get_files_in_one_folder(folder)
        if draw_list:
            self.fill_list(draw_list=draw_list)
            self.merge_files_button.setEnabled(True)

    @staticmethod
    def open_item(item):
        path = item.text()
        os.system(fr'explorer "{os.path.normpath(os.path.dirname(path))}"')
        os.startfile(path)

    def show_settings(self):
        self.settings_window.exec_()

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
        else:
            self.source_of_draws_field.setPlaceholderText(
                "Выберите папку с файлами в формате "
                ".cdw или .spw \n или файл с базой данных в формате .json"
            )

    def calculate_step(self, number_of_files, filter_only=False, get_data_base=False):
        self.current_progress = 0
        self.progress_step = 0
        number_of_operations = 0
        filters = self.get_all_filters()
        if get_data_base:
            number_of_operations = 1
        if filters:
            if not filters == self.previous_filters or filter_only:
                number_of_operations = 1
        if not filter_only and not get_data_base:
            number_of_operations += 2  # convert files and merger them
        if number_of_operations:
            self.progress_step = 100 / (number_of_operations * number_of_files)

    def increase_step(self, start=True):
        if start:
            self.current_progress += self.progress_step
            self.progress_bar.setValue(self.current_progress)

    def switch_button_group(self, switch=None):
        if not switch:
            switch = False if self.merge_files_button.isEnabled() else True
        self.merge_files_button.setEnabled(switch)
        self.choose_folder_button.setEnabled(switch)
        self.additional_settings_button.setEnabled(switch)
        self.refresh_draw_list_button.setEnabled(switch)
        self.choose_data_base_button.setEnabled(switch)
        self.choose_specification_button.setEnabled(switch)

    def choose_folder(self, signal):
        # signal отправляется из треда MyBrandThread
        dict_for_pdf = QtWidgets.QFileDialog.getExistingDirectory(self, "Выбрать папку", ".",
                                                                  QtWidgets.QFileDialog.ShowDirsOnly)
        if not dict_for_pdf:
            self.data_queue.put('Not_chosen')
        else:
            self.data_queue.put(dict_for_pdf)

    def get_all_filters(self):
        filters = {}
        if self.settings_window.checkBox.isChecked():
            date_1 = self.settings_window.dateEdit.dateTime().toSecsSinceEpoch()
            date_2 = self.settings_window.dateEdit_2.dateTime().toSecsSinceEpoch()
            filters['date_1'] = date_1
            filters['date_2'] = date_2
        if self.settings_window.checkBox_2.isChecked():
            if self.settings_window.search_by_spec_radio_button.isChecked():
                constructor_name = self.settings_window.lineEdit.text()
            else:
                constructor_name = str(self.settings_window.comboBox.currentText())
            if constructor_name:
                filters['constructor_name'] = constructor_name
        if self.settings_window.checkBox_3.isChecked():
            if self.settings_window.radio_button_4.isChecked():
                checker_name = self.settings_window.lineEdit_2.text()
            else:
                checker_name = str(self.settings_window.comboBox_2.currentText())
            if checker_name:
                filters['checker_name'] = checker_name
        return filters

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
        self.data_base_thread = DataBaseThread(files, refresh)
        self.data_base_thread.buttons_enable.connect(self.switch_button_group)
        self.data_base_thread.calculate_step.connect(self.calculate_step)
        self.data_base_thread.status.connect(self.status_bar.showMessage)
        self.data_base_thread.progress_bar.connect(self.progress_bar.setValue)
        self.data_base_thread.increase_step.connect(self.increase_step)
        self.data_base_thread.finished.connect(self.handle_data_base_results)
        self.data_base_thread.start()


class MyBrandThread(QThread):
    buttons_enable = pyqtSignal(bool)
    errors = pyqtSignal(str)
    status = pyqtSignal(str)
    choose_folder = pyqtSignal(bool)
    kill_thread = pyqtSignal()
    increase_step = pyqtSignal(bool)
    progress_bar = pyqtSignal(float)

    def __init__(self, files, directory, data_queue):
        self.files = files
        self.search_path = directory
        self.data_queue = data_queue
        self.constructor_class = QtWidgets.QFileDialog()
        QThread.__init__(self)

    def run(self):
        single_draw_dir, base_pdf_dir, main_name = self.create_folders()
        if not single_draw_dir and not base_pdf_dir:
            self.errors.emit('Запись прервана, папка не была найдена')
            self.buttons_enable.emit(True)
            self.progress_bar.emit(0)
            self.kill_thread.emit()
        self.cdw_to_pdf(self.files, single_draw_dir)
        pdf_file = self.merge_pdf_files(single_draw_dir, main_name)
        if merger.delete_single_draws_after_merge_checkbox.isChecked():
            shutil.rmtree(single_draw_dir)
        self.progress_bar.emit(0)
        self.buttons_enable.emit(True)
        os.system(f'explorer "{(os.path.normpath(os.path.dirname(single_draw_dir)))}"')
        if not merger.settings_window.checkBox_5.isChecked():
            os.startfile(pdf_file)
        self.status.emit(f'Закрытие Kompas')
        kompas_api.exit_kompas(self.appl)
        self.status.emit('Слитие успешно завершено')

    def create_folders(self):
        if merger.settings_window.radio_button_8.isChecked():
            base_pdf_dir, single_draw_dir, main_name = self.make_paths()
        else:
            self.choose_folder.emit(True)
            while True:
                time.sleep(0.1)
                try:
                    directory_to_save = self.data_queue.get(block=False)
                except queue.Empty:
                    pass
                else:
                    break
            if directory_to_save != 'Not_chosen':
                base_pdf_dir, single_draw_dir, main_name = self.make_paths(directory_to_save)
            else:
                return None, None, None
        os.makedirs(single_draw_dir)
        return single_draw_dir, base_pdf_dir, main_name

    def make_paths(self, directory_to_save=None):
        today_date = time.strftime("%d.%m.%Y")
        if merger.search_by_spec_radio_button.isChecked():
            main_name = os.path.basename(merger.specification_path)[:-4]
        else:
            main_name = os.path.basename(self.search_path)
        if not merger.settings_window.checkBox_5.isChecked():  # if not required to divide file
            base_pdf_dir = rf'{directory_to_save or self.search_path}\pdf'
            pdf_file = r'%s\%s - 01 %s.pdf' % (base_pdf_dir, main_name, today_date)
        else:
            base_pdf_dir = r'%s\pdf\%s - 01 %s' % (directory_to_save or self.search_path,
                                                   main_name, today_date)
            pdf_file = r'%s\%s.pdf' % (base_pdf_dir, main_name)
        single_draw_dir = os.path.splitext(pdf_file)[0] + " Однодетальные"
        # next code check if folder or file with same name exists if so:
        # get maximum number of file and folder and incriminate +1 to
        # name of new file and folder
        check_path = os.path.dirname(base_pdf_dir) if merger.settings_window.checkBox_5.isChecked() else base_pdf_dir
        if os.path.exists(check_path) and main_name in ' '.join(os.listdir(check_path)):
            string_of_files = ' '.join(os.listdir(check_path))
            today_update = max(map(int, re.findall(rf'{main_name} - (\d\d)(?= {today_date})', string_of_files)),
                               default=0)
            if today_update:
                today_update = str(today_update + 1) if today_update > 8 else '0' + str(today_update + 1)
                if merger.settings_window.checkBox_5.isChecked():
                    single_draw_dir = r'%s\pdf\%s - %s %s\Однодетальные' % (directory_to_save or self.search_path,
                                                                            main_name, today_update, today_date)
                else:
                    single_draw_dir = r'%s\pdf\%s - %s %s Однодетальные' % (directory_to_save or self.search_path,
                                                                            main_name, today_update, today_date)
        return base_pdf_dir, single_draw_dir, main_name

    def cdw_to_pdf(self, files, single_draw_dir):
        self.status.emit('Открытие Kompas')
        kompas_api7_module, application, const = kompas_api.get_kompas_api7()
        kompas6_api5_module, kompas_object, kompas6_constants = kompas_api.get_kompas_api5()
        self.appl = kompas_object
        doc_app, converter, _ = kompas_api.get_kompas_settings(application, kompas_object)
        number = 0
        for file in files:
            number += 1
            self.increase_step.emit(True)
            self.status.emit(f'Конвертация {file}')
            converter.Convert(file, single_draw_dir + "\\" +
                              f'{number} ' + os.path.basename(file) + ".pdf", 0, False
                              )

    def merge_pdf_files(self, directory_with_draws, main_name):
        files = sorted(os.listdir(directory_with_draws), key=lambda fil: int(fil.split()[0]))
        merger_instance = self.create_merger_instance(directory_with_draws, files)
        if type(merger_instance) is dict:  # if file gonna be divided
            for key, item in merger_instance.items():
                format_name, merger_writer, _ = item
                pdf_file = os.path.join(os.path.dirname(directory_with_draws),
                                        f'{format_name}-{main_name}.pdf')
                with open(pdf_file, 'wb') as pdf:
                    merger_writer.write(pdf)
                if merger.settings_window.checkBox_4.isChecked():
                    self.add_watermark(pdf_file)
            for key, item in merger_instance.items():
                _, _, files = item
                for file in files:
                    file.close()
        else:
            pdf_file = os.path.join(os.path.dirname(directory_with_draws),
                                    f'{os.path.basename(directory_with_draws)[:-14]}.pdf')
            with open(pdf_file, 'wb') as pdf:
                merger_instance.write(pdf)
            for file in merger_instance.inputs:
                file[0].close()
            if merger.settings_window.checkBox_4.isChecked():
                self.add_watermark(pdf_file)

        return pdf_file

    def create_merger_instance(self, directory, files):
        #  если нужно разбиваем файлы или сливаем всё в один
        if merger.settings_window.checkBox_5.isChecked():
            merger_instance = {595: ["A4", PyPDF2.PdfFileWriter(), []], 841: ["A3", PyPDF2.PdfFileWriter(), []],
                               1190: ["A2", PyPDF2.PdfFileWriter(), []], 1683: ["A1", PyPDF2.PdfFileWriter(), []]}
        else:
            merger_instance = PyPDF2.PdfFileMerger()
        for filename in files:
            file = open(os.path.join(directory, filename), 'rb')
            merger_pdf_reader = PyPDF2.PdfFileReader(file)
            if type(merger_instance) == dict:
                for page in merger_pdf_reader.pages:
                    size = int(sorted(page.mediaBox[2:])[0])
                    merger_instance[size][1].addPage(page)
                    merger_instance[size][2].append(file)
            else:
                merger_instance.append(fileobj=file)
            self.increase_step.emit(True)
            self.status.emit(f'Сливание {filename}')
        if merger.settings_window.checkBox_5.isChecked():
            merger_instance = {key: value for key, value in merger_instance.items() if
                               value[1].getNumPages()}
        return merger_instance

    def add_watermark(self, pdf_file):
        image = merger.settings_window.lineEdit_3.text()
        position = merger.settings_window.watermark_position
        if not image or not position:
            return
        if not os.path.exists(image) or image == 'Стандартный путь из настроек не существует':
            self.send_errors.emit(f'Путь к файлу с картинкой не существует')
            return
        pdf_doc = fitz.open(pdf_file)  # open the PDF
        rect = fitz.Rect(position)  # where to put image: use upper left corner
        for page in pdf_doc:
            if not page._isWrapped:
                page._wrapContents()
            try:
                page.insertImage(rect, filename=image, overlay=False)
            except ValueError:
                self.send_errors.emit(
                    'Заданы неверные координаты, размещения картинки, водяной знак не был добавлен'
                )
                return
        pdf_doc.saveIncr()  # do an incremental save


class FilterThread(QThread):
    finished = pyqtSignal(list, bool)
    increase_step = pyqtSignal(bool)
    status = pyqtSignal(str)

    def __init__(self, draw_list, filters, filter_only=True):
        self.draw_list = draw_list
        self.filters = filters
        self.appl = None
        self.filter_only = filter_only
        QThread.__init__(self)

    def run(self):
        self.filter_draws(self.draw_list, **self.filters)
        self.status.emit(f'Закрытие Kompas')
        kompas_api.exit_kompas(self.appl)
        self.finished.emit(self.draw_list, self.filter_only)

    def filter_draws(self, files, *, date_1=None, date_2=None,
                     constructor_name=None, checker_name=None):
        self.status.emit("Открытие Kompas")
        kompas_api7_module, application, const = kompas_api.get_kompas_api7()
        self.appl = application
        app = application.Application
        docs = app.Documents
        draw_list = []
        if date_1:
            date_1_in_seconds, date_2_in_seconds = sorted([date_1, date_2])
        app.HideMessage = const.ksHideMessageNo  # отключаем отображение сообщений Компас, отвечая на всё "нет"
        for file in files:  # структура обработки для каждого документа
            self.status.emit(f'Применение фильтров к {file}')
            self.increase_step.emit(True)
            doc, doc2d = kompas_api.get_right_api(file, docs, kompas_api7_module)
            draw_stamp = doc2d.LayoutSheets.Item(0).Stamp  # массив листов документа
            if date_1:
                date_in_stamp = draw_stamp.Text(130).Str
                if date_in_stamp:
                    try:
                        date_in_stamp = kompas_api.date_to_seconds(date_in_stamp)
                    except:
                        continue
                    if not date_1_in_seconds <= date_in_stamp <= date_2_in_seconds:
                        continue
            if constructor_name:
                constructor_in_stamp = draw_stamp.Text(110).Str
                if constructor_name not in constructor_in_stamp:
                    continue
            if checker_name:
                checker_in_stamp = draw_stamp.Text(115).Str
                if checker_in_stamp not in checker_in_stamp:
                    continue
            draw_list.append(file)
            doc.Close(const.kdDoNotSaveChanges)
        self.draw_list = draw_list


class DataBaseThread(QThread):
    increase_step = pyqtSignal(bool)
    status = pyqtSignal(str)
    finished = pyqtSignal(dict, bool, bool)
    progress_bar = pyqtSignal(float)
    calculate_step = pyqtSignal(int, bool, bool)
    buttons_enable = pyqtSignal(bool)

    def __init__(self, draw_list, refresh):
        self.draw_list = draw_list
        self.refresh = refresh
        self.appl = None
        self.files_dict = {}
        QThread.__init__(self)

    def run(self):
        self.buttons_enable.emit(False)
        self.create_data_base()
        self.progress_bar.emit(0)
        self.check_double_files()
        self.progress_bar.emit(0)
        self.buttons_enable.emit(True)
        if self.appl:
            reset = True
        else:
            reset = False
        self.finished.emit(self.files_dict, reset, self.refresh)

    def create_data_base(self):
        pythoncom.CoInitializeEx(0)
        shell = win32com.client.gencache.EnsureDispatch('Shell.Application', 0)  # подлкючаемся к винде
        dir_obj = shell.NameSpace(os.path.dirname(self.draw_list[0]))  # получаем объект папки виндовс шелл
        for x in range(355):
            cur_meta = dir_obj.GetDetailsOf(None, x)
            if cur_meta == 'Обозначение':
                meta_obozn = x  # присваиваем номер метаданных
                break
            # if cur_meta == 'Наименование':
            #     meta_name = x  # присваиваем номер метаданных
        for file in self.draw_list:
            self.status.emit(f'Получение атрибутов {file}')
            self.increase_step.emit(True)
            dir_obj = shell.NameSpace(os.path.dirname(file))  # получаем объект папки виндовс шелл
            item = dir_obj.ParseName(os.path.basename(file))  # указатель на файл (делаем именно объект винд шелл)
            doc_obozn = dir_obj.GetDetailsOf(item, meta_obozn).replace('$', '').\
                replace('|', '').replace(' ', '').strip().lower()
            # читаем обозначение мимо компаса, хз быстрее или нет
            # doc_name = dir_obj.GetDetailsOf(item, meta_name)  # читаем наименование мимо компаса, хз быстрее или нет
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
            kompas_api7_module, application, const = kompas_api.get_kompas_api7()
            self.appl = application
            app = application.Application
            docs = app.Documents
            app.HideMessage = const.ksHideMessageNo  # отключаем отображение сообщений Компас, отвечая на всё "нет"
            sum_len = len(double_files)
            self.calculate_step.emit(sum_len, False, True)
            for key, cdw_files, spw_files in double_files:
                temp_files = []
                self.increase_step.emit(True)
                if len(cdw_files) > 1:
                    self.files_dict[key] = [i for i in self.files_dict[key] if i.endswith('spw')]
                    path = self.get_right_path(cdw_files, const, kompas_api7_module, docs)
                    temp_files = [path]
                if len(spw_files) > 1:
                    self.files_dict[key] = [i for i in self.files_dict[key] if i.endswith('cdw')]
                    path = self.get_right_path(spw_files, const, kompas_api7_module, docs)
                    temp_files.append(path)
                if key in self.files_dict.keys():
                    self.files_dict[key].extend(temp_files)
                else:
                    self.files_dict[key] = temp_files
            self.status.emit(f'Закрытие Kompas')
            kompas_api.exit_kompas(self.appl)

    def get_right_path(self, same_double_paths, const, kompas_api7_module, docs):
        # Сначала сравнивает даты в штампе если они равны
        # Считывает дату создания и выбирает наиболее раннюю версию
        temp_dict = {}
        max_date_in_stamp = 0
        for path in same_double_paths:
            self.status.emit(f'Сравнение даты для {path}')
            date_of_creation = os.stat(path).st_ctime
            doc, doc2d = kompas_api.get_right_api(path, docs, kompas_api7_module)
            draw_stamp = doc2d.LayoutSheets.Item(0).Stamp  # массив листов документа
            date_in_stamp = draw_stamp.Text(130).Str
            try:
                date_in_stamp = kompas_api.date_to_seconds(date_in_stamp)
            except:
                date_in_stamp = 0
            if date_in_stamp >= max_date_in_stamp:
                if date_in_stamp > max_date_in_stamp:
                    max_date_in_stamp = date_in_stamp
                    temp_dict = {}
                temp_dict[path] = date_of_creation
            doc.Close(const.kdDoNotSaveChanges)
        sorted_paths = sorted(temp_dict, key=temp_dict.get)
        return sorted_paths[0]


class RecursionThread(QThread):
    status = pyqtSignal(str)
    finished = pyqtSignal(list, list, bool)
    buttons_enable = pyqtSignal(bool)
    errors = pyqtSignal(str)

    def __init__(self, specification, draws_in_specification, data_base_files, only_one_specification, refresh):
        self.draw_list = []
        self.missing_list = []
        self.appl = None
        self.refresh = refresh
        self.specification_path = specification
        self.only_one_specification = only_one_specification
        self.draws_in_specification = draws_in_specification
        self.data_base_files = data_base_files
        self.error = 1
        QThread.__init__(self)

    def run(self):
        self.buttons_enable.emit(False)
        self.process_specification()
        if self.appl:
            self.status.emit(f'Закрытие Kompas')
            kompas_api.exit_kompas(self.appl)
        self.finished.emit(self.missing_list, self.draw_list, self.refresh)

    def process_specification(self):
        self.draw_list.append(self.specification_path)
        self.recursive_draws_traversal(self.draws_in_specification)

    def recursive_draws_traversal(self, draw_list, spec_path: Optional[str] = None):
        # spec_path берется не None если идет рекурсия
        spec_path = spec_path or self.specification_path
        self.status.emit(f'Обработка {os.path.basename(spec_path)}')

        for group_name, draws_description in draw_list.items():
            if group_name in ["Сборочные чертежи", "Детали"]:
                for draw_full_name in draws_description:
                    cdw_file_path = self.fetch_draw_path(
                        draw_full_name,
                        file_extension='.cdw',
                        file_path=spec_path
                    )
                    if cdw_file_path:
                        self.draw_list.append(cdw_file_path)
            else:  # сборочные единицы
                for draw_full_name in draws_description:
                    spw_file_path = self.fetch_draw_path(
                        draw_full_name,
                        file_extension='.spw',
                        file_path=spec_path
                    )
                    if not spw_file_path:
                        continue
                    self.draw_list.append(spw_file_path)

                    if self.only_one_specification:
                        response = kompas_api.get_draws_from_specification(
                            spw_file_path,
                            only_document_list=True,
                            draw_obozn=draw_full_name[0]
                        )
                        if type(response) == str:  # ошибка при открытии спецификации
                            self.missing_list.append(response)
                            continue

                        cdw_file, self.appl, errors = response
                        if errors:
                            self.missing_list.extend(errors)
                    else:
                        response = kompas_api.get_draws_from_specification(
                            spw_file_path,
                            draw_obozn=draw_full_name[0]
                        )
                        if type(response) == str:  # ошибка при открытии спецификации
                            self.missing_list.append(response)
                            continue

                        draws_in_specification, self.appl, errors = response
                        if errors:
                            self.missing_list.extend(errors)
                        self.recursive_draws_traversal(draws_in_specification, spw_file_path)

    def fetch_draw_path(self, item: tuple[str, str], file_extension: str, file_path: str) -> Optional[str]:
        draw_obozn, draw_name = item
        try:
            draw_path = self.data_base_files[draw_obozn.lower()]
        except KeyError:
            draw_path = []
            if file_extension == ".spw":
                draw_path = self.fetch_spec_group_path(draw_obozn)
            if file_extension == ".cdw" or not draw_path:
                spec_path = os.path.basename(file_path)
                missing_draw = [draw_obozn.upper(), draw_name.capitalize().replace('\n', ' ')]
                missing_draw_info = (spec_path, ' - '.join(missing_draw))
                self.missing_list.append(missing_draw_info)
                return
        if len(draw_path) > 1:
            draw_path = [i for i in draw_path if i.endswith(file_extension)]

        if os.path.exists(draw_path[0]):
            return draw_path[0]
        else:
            if self.error == 1:  # print this message only once
                self.errors.emit('Некоторые пути в базе являются недействительным,'
                                 'обновите базу данных')
                self.error += 1
            return

    def fetch_spec_group_path(self, spec_obozn: str) -> Optional[str]:
        draw_info = kompas_api.fetch_obozn_execution_and_name(spec_obozn)
        if not draw_info:
            last_symbol = ""
            if not spec_obozn[-1].isdigit():
                last_symbol = spec_obozn[-1]

            db_obozn = spec_obozn
            spec_path = ""
            for num in range(1, 4): #  обычно максимальное количество исполнений до -03
                db_obozn += spec_obozn + f"-0{num}{last_symbol}"
                spec_path = self.data_base_files.get(db_obozn)
                if spec_path:
                    break
            if spec_path:
                return spec_path
            return

        obozn, execution = draw_info
        last_symbol = ""
        if not execution[-1].isdigit():  # если есть буква в конце
            last_symbol = execution[-1]
            execution = execution[:-1]

        # чертежи идут с четными номерами, парсим нечетные
        number_of_execution = int(execution) - 1
        spc_execution = ""  # исполнеение для поиска по бд
        if number_of_execution:
            spc_execution = f"-0{number_of_execution}" + last_symbol

        obozn += spc_execution
        spec_path = self.data_base_files.get(obozn.strip())
        is_that_correct_path = kompas_api.verify_its_group_spec_path(spec_path[0], execution)
        if not is_that_correct_path:
            return

        return spec_path


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    merger = UiMerger()
    merger.show()
    sys.excepthook = except_hook
    sys.exit(app.exec_())

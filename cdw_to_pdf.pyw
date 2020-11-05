# -*- coding: utf-8 -*-
import sys
import os
import queue
import time
import kompas_api
import PyPDF2
import shutil
import pythoncom
import itertools
import json
from win32com.client import Dispatch
import win32com
import fitz
from operator import itemgetter
from Widgets_class import MakeWidgets
from settings_window import SettingsWindow
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QThread, pyqtSignal

dict_for_pdf = None


def except_hook(cls, exception, traceback):
    sys.__excepthook__(cls, exception, traceback)


class Ui_Merger(MakeWidgets):
    def __init__(self, parent=None):
        MakeWidgets.__init__(self, parent=None)
        self.kompas_ext = ['.cdw', '.spw']
        self.setFixedSize(929, 646)
        self.setupUi(self)
        self.search_path = None
        self.current_progress = 0
        self.progress_step = 0
        self.settings_window = SettingsWindow()
        self.filter_thread = None
        self.data_queue = None
        self.missing_list = []
        self.draw_list = []
        self.appl = None
        self.checkBox_4_status = 'Yes'
        self.checkBox_4_current_status = 'Yes'
        self.checkBox_5_status = 'No'
        self.checkBox_5_current_status = 'No'
        self.thread = None
        self.data_base_thread = None
        self.data_base_files = None
        self.specification_path = None
        self.previous_filters = {}
        self.draws_in_specification = {}
        self.recursive_thread = None

    def setupUi(self, Merger):
        Merger.setObjectName("Merger")
        font = QtGui.QFont()
        sizepolicy =\
            QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Ignored)
        sizepolicy_button = \
            QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Ignored)
        font.setFamily("Arial")
        font.setPointSize(12)
        Merger.setFont(font)

        self.centralwidget = QtWidgets.QWidget(Merger)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(10, 10, 906, 611))
        self.gridLayoutWidget.setObjectName("gridLayoutWidget")
        self.gridLayout = QtWidgets.QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)

        self.pushButton = self.make_button(text='Выделить все', font=font, command=self.select_all,
                                           size_policy=sizepolicy_button)
        self.gridLayout.addWidget(self.pushButton, 10, 0, 1, 1)

        self.pushButton_2 = self.make_button(text="Снять выделение", font=font, command=self.unselect_all,
                                             size_policy=sizepolicy_button)
        self.gridLayout.addWidget(self.pushButton_2, 10, 1, 1, 1)

        self.pushButton_3 = self.make_button(text="Добавить файл в список", font=font,
                                             command=self.add_file_to_list)
        self.gridLayout.addWidget(self.pushButton_3, 10, 2, 1, 1)

        self.pushButton_4 = self.make_button(text="Добавить папку в список", font=font,
                                             command=self.add_folder_to_list)
        self.gridLayout.addWidget(self.pushButton_4, 10, 3, 1, 1)

        self.pushButton_5 = self.make_button(text="Склеить файлы", font=font, command=self.check_lines)
        self.gridLayout.addWidget(self.pushButton_5, 13, 0, 1, 4)

        self.pushButton_6 = self.make_button(text="Выбор папки \n с чертежами для поиска", font=font,
                                             command=self.choose_initial_folder)
        self.gridLayout.addWidget(self.pushButton_6, 1, 2, 1, 1)

        self.pushButton_7 = self.make_button(text="Очистить спискок и выбор папки для поиска", font=font,
                                             command=self.clear_data)
        self.gridLayout.addWidget(self.pushButton_7, 4, 0, 1, 2)

        self.pushButton_8 = self.make_button(text="Дополнительные настройки", font=font,
                                             command=self.show_settings)
        self.gridLayout.addWidget(self.pushButton_8, 11, 0, 1, 2)

        self.pushButton_9 = self.make_button(text='Обновить список файлов для склеивания', font=font,
                                             command=self.refresh_draws_in_list)
        self.gridLayout.addWidget(self.pushButton_9, 4, 2, 1, 2)

        self.pushButton_10 = self.make_button(text='Выбор файла\n с базой чертежей', font=font, enabled=False,
                                              command=self.get_data_base_path)
        self.gridLayout.addWidget(self.pushButton_10,  1, 3, 1, 1)

        self.pushButton_11 = self.make_button(text='Выбор \nспецификации', font=font, enabled=False,
                                              command=self.choose_specification)
        self.gridLayout.addWidget(self.pushButton_11, 2, 2, 1, 1)

        self.pushButton_13 = self.make_button(text='Сохранить \n базу чертежей', font=font, enabled=False,
                                              command=self.apply_data_base_save)
        self.gridLayout.addWidget(self.pushButton_13, 2, 3, 1, 1)


        self.checkBox_3 = self.make_checkbox(font=font, text='Удалить однодетальные pdf-чертежи по окончанию',
                                             activate=True)
        self.gridLayout.addWidget(self.checkBox_3, 11, 2, 1, 2)

        self.checkBox_4 = self.make_checkbox(font=font, text='С обходом всех папок внутри', activate=True,
                                             command=self.change_checkbox_4_status)
        self.gridLayout.addWidget(self.checkBox_4, 3, 3, 1, 1)

        self.checkBox_5 = self.make_checkbox(font=font, text='С поиском по подсборкам',
                                             command=self.change_checkbox_5_status)
        self.checkBox_5.setEnabled(False)
        self.gridLayout.addWidget(self.checkBox_5, 3, 1, 1, 1)

        self.line_edit = self.make_text_edit(font=font, placeholder="Выберите папку с файлами в формате .spw или"
                                             " .cdw", size_policy=sizepolicy)
        self.gridLayout.addWidget(self.line_edit, 1, 0, 1, 2)

        self.line_edit_2 = self.make_text_edit(font=font, placeholder="Укажите путь до файла со спецификацией"
                                               " .sdw", size_policy=sizepolicy)
        self.line_edit_2.setEnabled(False)
        self.gridLayout.addWidget(self.line_edit_2, 2, 0, 1, 2)

        self.radio_button = self.make_radio_button(text='Поиск по папке', font=font, command=self.choose_search_way)
        self.radio_button.setChecked(True)
        self.gridLayout.addWidget(self.radio_button, 3, 2, 1, 1)

        self.radio_button_2 = self.make_radio_button(text='Поиск по спецификации', font=font,
                                                     command=self.choose_search_way)
        self.gridLayout.addWidget(self.radio_button_2, 3, 0, 1, 1)


        self.listWidget = QtWidgets.QListWidget(self.gridLayoutWidget)
        self.listWidget.itemDoubleClicked.connect(self.open_item)
        self.gridLayout.addWidget(self.listWidget, 8, 0, 1, 4)

        self.progressBar = QtWidgets.QProgressBar(self.gridLayoutWidget)
        self.progressBar.setTextVisible(False)
        self.gridLayout.addWidget(self.progressBar, 15, 0, 1, 4)

        Merger.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(Merger)
        self.statusbar.setObjectName("statusbar")
        Merger.setStatusBar(self.statusbar)
        QtCore.QMetaObject.connectSlotsByName(Merger)

    def choose_initial_folder(self):
        directory = QtWidgets.QFileDialog.getExistingDirectory(self, "Выбрать папку", ".")
        if directory:
            draw_list = self.check_search_path(directory)
            if not draw_list:
                return
            self.line_edit.setText(self.search_path)
            if self.radio_button.isChecked():
                self.calculate_step(len(draw_list), filter_only=True)
                if self.progress_step:
                    self.apply_filters(draw_list)
                else:
                    self.fill_list(draw_list=draw_list)
            else:
                self.get_data_base(draw_list)

    def check_search_path(self, search_path):
        if os.path.isfile(search_path):
            if self.radio_button_2.isChecked():
                search_path = self.check_data_base_file(search_path)
            else:
                self.error('Укажите папку для поиска с файламиа, а не файл')
        else:
            draw_list = self.get_all_files_in_folder(search_path)
            if draw_list:
                self.search_path = search_path
                return draw_list
            else:
                return None

    def choose_specification(self):
        filename = [QtWidgets.QFileDialog.getOpenFileName(self, "Выбрать файл", ".",
                                                          "Спецификация(*.spw)")[0]][0]
        if filename:
            response = self.check_specification(filename)
            if not response:
                return
            self.line_edit_2.setText(filename)
            cdw_file = self.line_edit.toPlainText()
            if cdw_file:
                if cdw_file == self.search_path and self.data_base_files:
                    self.get_paths_to_specifications()
                else:
                    self.get_data_base()

    def check_specification(self, filename):
        if filename.endswith('.spw') and os.path.isfile(filename):
            draws_in_specification, self.appl = kompas_api.get_draws_from_specification(filename)
            if draws_in_specification:
                self.draws_in_specification = draws_in_specification
                self.specification_path = filename
            else:
                self.error('Спецификация пуста, как и вся наша жизнь')
                return
        else:
            self.error('Указанный файл не является спецификацией или не существует')
            return
        return 1

    def get_paths_to_specifications(self, refresh=False):
        self.listWidget.clear()
        filename = self.line_edit_2.toPlainText()
        if not filename:
            return
        if filename != self.specification_path:
            response = self.check_specification(filename)
            if not response:
                return
        self.start_recursion(refresh)

    def handle_specification_result(self, missing_list, draw_list, refresh):
        self.statusbar.showMessage('Завершено получение файлов из спецификации')
        self.appl = None
        self.missing_list = missing_list
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
        only_one_specification = self.checkBox_5.isChecked()
        self.checkBox_5_status = 'Yes' if self.checkBox_5.isChecked() else 'No'
        self.recursive_thread = RecursionThread(self.specification_path, self.draws_in_specification,
                                                self.data_base_files, only_one_specification, refresh)
        self.recursive_thread.buttons_enable.connect(self.switch_button_group)
        self.recursive_thread.finished.connect(self.handle_specification_result)
        self.recursive_thread.status.connect(self.statusbar.showMessage)
        self.recursive_thread.start()

    def change_checkbox_5_status(self):
        if self.checkBox_5.isChecked():
            self.checkBox_5_current_status = 'Yes'
        else:
            self.checkBox_5_current_status = 'No'

    def change_checkbox_4_status(self):
        if self.checkBox_4.isChecked():
            self.checkBox_4_current_status = 'Yes'
        else:
            self.checkBox_4_current_status = 'No'

    def print_out_missing_files(self):
        grouped_list = itertools.groupby(self.missing_list, itemgetter(0))
        grouped_list = [key + ':\n' + '\n'.join(['----' + v for k, v in value]) for key, value in grouped_list]
        missing_message = '\n'.join(grouped_list)
        choice = QtWidgets.QMessageBox.question(self, 'Отсуствующие чертежи',
                                                f"Не были найдены следующи чертежи:\n{missing_message},"
                                                f"\nСохранить список?",
                                                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if choice == QtWidgets.QMessageBox.Yes:
            filename = QtWidgets.QFileDialog.getSaveFileName(self, "Сохранить файл", ".", "txt(*.txt)")[0]
            if filename:
                try:
                    with open(filename, 'w') as file:
                        file.write(missing_message)
                except:
                    self.error("Ошибка записи")
                    return

    def refresh_draws_in_list(self, refresh=False):
        if self.radio_button.isChecked():
            search_path = self.line_edit.toPlainText()
            if not search_path:
                self.error('Укажите папку с чертежамиили')
                return
            draw_list = self.check_search_path(search_path)
            if not draw_list:
                return
            self.listWidget.clear()
            self.calculate_step(len(draw_list), filter_only=True)
            if self.progress_step:
                self.apply_filters(draw_list)
            else:
                self.error('Обновление завершено')
                self.fill_list(draw_list=draw_list)
        else:
            search_path = self.line_edit.toPlainText()
            specification = self.line_edit_2.toPlainText()
            if not search_path:
                self.error('Укажите папку с чертежамиили файл .json')
                return
            if not specification:
                self.error('Укажите файл спецификации')
                return
            response_1 = self.check_specification(specification)
            if not response_1:
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
            if filters == self.previous_filters and self.search_path == self.line_edit.toPlainText():
                return 1
        self.previous_filters = filters
        self.filter_thread = FilterThread(draw_list, filters, filter_only)
        self.filter_thread.status.connect(self.statusbar.showMessage)
        self.filter_thread.increase_step.connect(self.increase_step)
        self.filter_thread.finished.connect(self.handle_filter_results)
        self.filter_thread.start()

    def handle_filter_results(self, draw_list, filter_only=True):
        self.statusbar.showMessage('Филтрация успешно завршена')
        self.appl = None
        if not draw_list:
            self.error('Нету файлов .cdw или .spw, в выбранной папке(ах) с указанными параметрами')
            self.current_progress = 0
            self.progressBar.setValue(self.current_progress)
            self.listWidget.clear()
            self.switch_button_group(True)
            return
        if filter_only:
            self.current_progress = 0
            self.progressBar.setValue(self.current_progress)
            self.fill_list(draw_list=draw_list)
            self.switch_button_group(True)
            return
        else:
            self.start_merge_process(draw_list)

    def closeEvent(self, event):
        if self.appl:
            kompas_api.exit_kompas(self.appl)
        event.accept()

    def handle_data_base_results(self, data_base, appl, refresh=False):
        self.statusbar.showMessage('Завершено получение Базы Данных')
        if appl:
            self.appl = None
        if data_base:
            self.data_base_files = data_base
            self.save_data_base()
            self.get_paths_to_specifications(refresh)
        else:
            self.error('Нету файлов с обозначением в штампе')

    def save_data_base(self):
        choice = QtWidgets.QMessageBox.question(self, 'База данных',
                                                "Сохранить полученные данные?",
                                                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if choice == QtWidgets.QMessageBox.Yes:
            self.apply_data_base_save()
        else:
            QtWidgets.QMessageBox.information(self, 'Отмена записи', 'Данные о связях хранятся в памяти')
            self.pushButton_13.setEnabled(True)

    def apply_data_base_save(self):
        filename = QtWidgets.QFileDialog.getSaveFileName(self, "Сохранить файл", ".", "Json file(*.json)")[0]
        if filename:
            try:
                with open(filename, 'w') as file:
                    json.dump(self.data_base_files, file, ensure_ascii=False)
            except:
                self.error("В базе данных имеются ошибки")
                return
            self.line_edit.setText(filename)
            self.search_path = filename
        if self.pushButton_13.isEnabled():
            QtWidgets.QMessageBox.information(self, 'Запись данных', 'Запись данных успешно произведена')
            self.pushButton_13.setEnabled(False)

    def get_data_base_path(self):
        filename = QtWidgets.QFileDialog.getOpenFileName(self, "Загрузить файл", ".", "Json file(*.json)")[0]
        if filename:
            if filename:
                self.load_data_base(filename)
            else:
                self.error('Указанный файл не является файлом .json')

    def load_data_base(self, filename, refresh=False):
        filename = filename or self.search_path
        response = self.check_data_base_file(filename)
        if not response:
            return
        self.search_path = filename
        self.line_edit.setText(filename)
        self.get_paths_to_specifications(refresh)

    def check_data_base_file(self, filename):
        if not os.path.exists(filename):
            self.error('Указан несуществующий путь')
            return
        if not filename.endswith('.json'):
            self.error('Указанный файл не является файлом .json')
            return None
        with open(filename) as file:
            try:
                self.data_base_files = json.load(file)
            except json.decoder.JSONDecodeError as e:
                self.error('В Файл settings.txt \n присутсвуют ошибки \n синтаксиса json')
                return None
            except UnicodeDecodeError:
                self.error('Указана неверная кодировка файл, попытайтесь еще раз сгенерировать данные')
                return None
            else:
                return 1

    def get_all_files_in_folder(self, search_path=None):
        search_path = search_path or self.search_path
        draw_list = []
        except_folders_list = self.get_items_in_list(self.settings_window.listWidget)
        self.listWidget.clear()
        self.checkBox_4_status = 'Yes' if self.checkBox_4.isChecked() else 'No'
        if self.checkBox_4.isChecked() or self.radio_button_2.isChecked():
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
            self.error('Нету файлов .cdw или .spw, в выбраной папке(ах) с указанными параметрами')

    def get_files_in_one_folder(self, folder):
        # sorting in the way so specification sheet is the first if .cdw and .spw files has the same name
        draw_list = [os.path.splitext(i) for i in os.listdir(folder) if os.path.splitext(i)[1] in self.kompas_ext]
        draw_list = sorted([(i[0], '.adw' if i[1] == '.spw' else '.cdw') for i in draw_list], key=itemgetter(0, 1))
        draw_list = [os.path.join(folder, (i[0] + '.spw' if i[1] == '.adw' else i[0] + '.cdw')) for i in draw_list]
        draw_list = map(os.path.normpath, draw_list)
        return draw_list

    def check_lines(self):
        if self.radio_button_2.isChecked() and (self.search_path != self.line_edit.toPlainText()
                                                or self.specification_path != self.line_edit_2.toPlainText()
                                                or self.checkBox_5_current_status != self.checkBox_5_status):
            choice = QtWidgets.QMessageBox.question(self, 'Изменения',
                                                    f"Настройки поиска файлов и/или спецификация были изменены. "
                                                    f"Обновить список файлов?",
                                                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            if choice == QtWidgets.QMessageBox.Yes:
                self.refresh_draws_in_list(refresh=True)
            else:
                self.merge_files_in_one()
        elif self.radio_button.isChecked() and (self.search_path != self.line_edit.toPlainText()
                                                or self.checkBox_4_current_status != self.checkBox_4_status) :
            choice = QtWidgets.QMessageBox.question(self, 'Изменения',
                                                    f"Путь или настройки поиска были изменены. "
                                                    f"Обновить список файлов?",
                                                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
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
            self.error('Нету файлов для слития')
            return
        self.calculate_step(len(draws_list))
        self.switch_button_group(False)
        reply = self.apply_filters(draws_list, False)
        if reply:
            self.start_merge_process(draws_list)

    def start_merge_process(self, draws_list):
        self.data_queue = queue.Queue()
        self.thread = MyBrandThread(draws_list, self.search_path, self.data_queue)
        self.thread.buttons_enable.connect(self.switch_button_group)
        self.thread.increase_step.connect(self.increase_step)
        self.thread.kill_thread.connect(self.stop_merge_thread)
        self.thread.errors.connect(self.error)
        self.thread.choose_folder.connect(self.choose_folder)
        self.thread.status.connect(self.statusbar.showMessage)
        self.thread.progress_bar.connect(self.progressBar.setValue)
        self.appl = None
        self.thread.start()

    def stop_merge_thread(self):
        self.switch_button_group(True)
        self.statusbar.showMessage('Папка не выбрана, запись прервана')
        self.thread.terminate()

    def clear_data(self):
        self.search_path = None
        self.specification_path = None
        self.line_edit.clear()
        if self.radio_button.isChecked():
            self.line_edit.setPlaceholderText("Выберите папку с файлами в формате .cdw или .spw")
        else:
            self.line_edit.setPlaceholderText("Выберите папку с файлами в формате "
                                              ".cdw или .spw \n или файл с базой данных в формате .json")
        self.line_edit_2.clear()
        self.line_edit_2.setPlaceholderText('Укажите путь до файла со спецификацией')
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
        filename = [QtWidgets.QFileDialog.getOpenFileName(self, "Выбрать файл", ".",
                                                                "Чертж(*.cdw);;"
                                                                "Спецификация(*.spw)")[0]]
        if filename[0]:
            self.fill_list(draw_list=filename)
            self.pushButton_5.setEnabled(True)

    def add_folder_to_list(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Выбрать папку", ".")
        draw_list = []
        if folder:
            draw_list = self.get_files_in_one_folder(folder)
        if draw_list:
            self.fill_list(draw_list=draw_list)
            self.pushButton_5.setEnabled(True)

    @staticmethod
    def open_item(item):
        path = item.text()
        os.system(fr'explorer "{os.path.normpath(os.path.dirname(path))}"')
        os.startfile(path)

    def show_settings(self):
        self.settings_window.exec_()

    def choose_search_way(self):
        self.pushButton_10.setEnabled(self.radio_button_2.isChecked())
        self.pushButton_11.setEnabled(self.radio_button_2.isChecked())
        self.line_edit_2.setEnabled(self.radio_button_2.isChecked())
        self.checkBox_5.setEnabled(self.radio_button_2.isChecked())
        self.checkBox_4.setEnabled(self.radio_button.isChecked())
        if self.radio_button.isChecked():
            self.line_edit.setPlaceholderText("Выберите папку с файлами в формате .cdw или .spw")
        else:
            self.line_edit.setPlaceholderText("Выберите папку с файлами в формате "
                                              ".cdw или .spw \n или файл с базой данных в формате .json")

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
            self.progressBar.setValue(self.current_progress)

    def switch_button_group(self, switch=None):
        if not switch:
            switch = False if self.pushButton_5.isEnabled() else True
        self.pushButton_5.setEnabled(switch)
        self.pushButton_6.setEnabled(switch)
        self.pushButton_8.setEnabled(switch)
        self.pushButton_9.setEnabled(switch)
        self.pushButton_10.setEnabled(switch)
        self.pushButton_11.setEnabled(switch)

    def choose_folder(self, signal):
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
            if self.settings_window.radio_button_2.isChecked():
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
        elif self.line_edit.toPlainText():
            filename = self.line_edit.toPlainText()
            if os.path.isdir(filename):
                files = self.check_search_path(filename)
                if files:
                    self.get_data_base_from_folder(files, refresh)
            else:
                self.load_data_base(filename, refresh)
        else:
            self.error('Укажите папку для поиска файлов')

    def get_data_base_from_folder(self, files, refresh=False):
        self.calculate_step(len(files), get_data_base=True)
        self.data_base_thread = DataBaseThread(files, refresh)
        self.data_base_thread.buttons_enable.connect(self.switch_button_group)
        self.data_base_thread.calculate_step.connect(self.calculate_step)
        self.data_base_thread.status.connect(self.statusbar.showMessage)
        self.data_base_thread.progress_bar.connect(self.progressBar.setValue)
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
        directory_pdf, base_pdf_dir = self.create_folders()
        if not directory_pdf and not base_pdf_dir:
            self.errors.emit('Запись прервана, папка не была найдена')
            self.buttons_enable.emit(True)
            self.progress_bar.emit(0)
            self.kill_thread.emit()
        self.cdw_to_pdf(self.files, directory_pdf)
        pdf_file = self.merge_pdf_files(directory_pdf)
        if merger.checkBox_3.isChecked():
            shutil.rmtree(directory_pdf)
        self.progress_bar.emit(0)
        self.buttons_enable.emit(True)
        os.system(f'explorer "{os.path.normpath(base_pdf_dir)}"')
        os.startfile(pdf_file)
        kompas_api.exit_kompas(self.appl)
        self.status.emit('Слитие успешно завершено')

    def create_folders(self):
        main_name, base_pdf_dir, pdf_file, directory_pdf = self.make_paths()
        if not os.path.exists(base_pdf_dir):
            try:
                os.makedirs(base_pdf_dir)
            except FileNotFoundError:
                self.choose_folder.emit(True)
                while True:
                    time.sleep(0.1)
                    try:
                        data = self.data_queue.get(block=False)
                    except queue.Empty:
                        pass
                    else:
                        break
                if data != 'Not_chosen':
                    self.search_path = data
                    main_name, base_pdf_dir, pdf_file, directory_pdf = self.make_paths()
                else:
                    return None, None

        if os.path.exists(pdf_file) or os.path.exists(directory_pdf):
            # check if folder or file with same name exists if so:
            # get maximum number of file and folder and incriminate +1 to
            # name of new file and folder
            created_files = max([int(i.split()[-2]) for i in os.listdir(base_pdf_dir)
                                 if i.endswith(f'{time.strftime("%d.%m.%Y.pdf")}')], default=0)
            created_folders = max([int(i.split()[-2]) for i in os.listdir(base_pdf_dir)
                                   if i.endswith(f'{time.strftime("%d.%m.%Y")}')], default=0)
            today_update = created_files if created_files >= created_folders else created_folders
            today_update = str(today_update + 1) if today_update > 8 else '0' + str(today_update + 1)
            directory_pdf = r'%s\pdf\%s - %s %s' % \
                            (self.search_path, main_name, today_update, time.strftime("%d.%m.%Y"))
        os.makedirs(directory_pdf)
        return directory_pdf, base_pdf_dir

    def make_paths(self):
        main_name = os.path.basename(self.search_path)
        base_pdf_dir = rf'{self.search_path}\pdf'
        pdf_file = r'%s\%s - 01 %s.pdf' % (base_pdf_dir, main_name, time.strftime("%d.%m.%Y"))
        directory_pdf = os.path.splitext(pdf_file)[0]
        return main_name, base_pdf_dir, pdf_file, directory_pdf

    def cdw_to_pdf(self, files, directory_pdf):
        kompas_api7_module, application, const = kompas_api.get_kompas_api7()
        kompas6_api5_module, kompas_object, kompas6_constants = kompas_api.get_kompas_api5()
        self.appl = kompas_object
        doc_app , iConverter, _ = kompas_api.get_kompas_settings(application, kompas_object)
        number = 0
        for file in files:
            number += 1
            self.increase_step.emit(True)
            self.status.emit(f'Конвертация {file}')
            iConverter.Convert(file, directory_pdf + "\\" +
                               f'{number} ' + os.path.basename(file) + ".pdf", 0, False)

    def merge_pdf_files(self, directory):
        files = sorted(os.listdir(directory), key=lambda fil: int(fil.split()[0]))
        merger_pdf = PyPDF2.PdfFileMerger()
        for filename in files:
            merger_pdf.append(fileobj=open(os.path.join(directory, filename), 'rb'))
            self.increase_step.emit(True)
            self.status.emit(f'Сливание {filename}')
        if merger.settings_window.checkBox_5.isChecked():
            input_pages = sorted([(i, i.pagedata['/MediaBox'][2:]) for i in merger_pdf.pages], key=itemgetter(1))
            merger_pdf.pages = [i[0] for i in input_pages]
        pdf_file = os.path.join(os.path.dirname(directory), f'{os.path.basename(directory)}.pdf')
        with open(pdf_file, 'wb') as pdf:
            merger_pdf.write(pdf)
        if merger.settings_window.checkBox_4.isChecked():
            self.add_watermark(pdf_file)
        for file in merger_pdf.inputs:
            file[0].close()
        return pdf_file

    def add_watermark(self, pdf_file):
        image = merger.settings_window.lineEdit_3.text()
        position = merger.settings_window.watermark_position
        if not image or not position:
            return
        if not os.path.exists(image) or image == 'Стандартный путь из настроек не существует':
            self.errors.emit(f'Путь к файлу с картинкой не существует')
            return
        pdf_doc = fitz.open(pdf_file)  # open the PDF
        rect = fitz.Rect(position)  # where to put image: use upper left corner
        for page in pdf_doc:
            if not page._isWrapped:
                page._wrapContents()
            try:
                page.insertImage(rect, filename=image, overlay=False)
            except ValueError:
                self.errors.emit('Заданы неверные координаты, размещения картинки, водяной знак не был добавлен')
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
        kompas_api.exit_kompas(self.appl)
        self.finished.emit(self.draw_list, self.filter_only)

    def filter_draws(self, files, *, date_1=None, date_2=None,
                     constructor_name=None, checker_name=None):
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
            iStamp = doc2d.LayoutSheets.Item(0).Stamp  # массив листов документа
            if date_1:
                date_in_stamp = iStamp.Text(130).Str
                if date_in_stamp:
                    try:
                        date_in_stamp = kompas_api.date_to_seconds(date_in_stamp)
                    except:
                        continue
                    if not date_1_in_seconds <= date_in_stamp <= date_2_in_seconds:
                        continue
            if constructor_name:
                constructor_in_Stamp = iStamp.Text(110).Str
                if not constructor_name in constructor_in_Stamp:
                    continue
            if checker_name:
                checker_in_Stamp = iStamp.Text(115).Str
                if not checker_in_Stamp in checker_in_Stamp:
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
        self.get_data_base()
        self.progress_bar.emit(0)
        self.check_double_files()
        self.progress_bar.emit(0)
        self.buttons_enable.emit(True)
        if self.appl:
            reset = True
        else:
            reset = False
        self.finished.emit(self.files_dict, reset, self.refresh)

    def get_data_base(self):
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
            doc_obozn = dir_obj.GetDetailsOf(item, meta_obozn).replace('$', '').replace('|', '').replace(' ',
                                                                                                         '').strip().lower()
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
            iStamp = doc2d.LayoutSheets.Item(0).Stamp  # массив листов документа
            date_in_stamp = iStamp.Text(130).Str
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

    def __init__(self, specification, draws_in_specification, data_base_files, only_one_specification, refresh):
        self.draw_list = []
        self.missing_list = []
        self.appl = None
        self.refresh = refresh
        self.specification_path = specification
        self.only_one_specification = only_one_specification
        self.draws_in_specification = draws_in_specification
        self.data_base_files = data_base_files
        QThread.__init__(self)

    def run(self):
        self.buttons_enable.emit(False)
        self.process_specification()
        if self.appl:
            kompas_api.exit_kompas(self.appl)
        self.finished.emit(self.missing_list, self.draw_list, self.refresh)

    def process_specification(self):
        self.draw_list.append(self.specification_path)
        self.recursive_traversal(self.draws_in_specification)

    def recursive_traversal(self, draw_list, drawisimo=None):
        drawisimo = drawisimo or self.specification_path
        self.status.emit(f'Обработка {os.path.basename(drawisimo)}')
        for key, value in draw_list.items():
            if key == 'Сборочные чертежи' or key == 'Детали':
                for item in value:
                    draw = self.check_item(item, '.cdw', drawisimo)
                    if draw:
                        self.draw_list.append(draw[0])
            else:
                for item in value:
                    draw = self.check_item(item, '.spw', drawisimo)
                    if draw:
                        if self.only_one_specification:
                            self.draw_list.append(draw[0])
                            draw_list, self.appl = kompas_api.get_draws_from_specification(draw[0])
                            self.recursive_traversal(draw_list, draw[0])
                        else:
                            self.draw_list.append(draw[0])
                            cdw_file, self.appl = kompas_api.get_draws_from_specification(draw[0], True)
                            cdw_file = self.check_item(cdw_file[0], '.cdw', drawisimo)
                            if cdw_file:
                                self.draw_list.append(cdw_file[0])

    def check_item(self, item, extension, drawisimo):
        try:
            draw = self.data_base_files[item[0].lower()]
        except KeyError:
            spec_path = os.path.basename(drawisimo)
            item = [item[0].upper(), item[1].capitalize().replace('\n', ' ')]
            new_item = (spec_path, ' - '.join(item))
            self.missing_list.append(new_item)
            return
        else:
            if len(draw) > 1:
                draw = [i for i in draw if i.endswith(extension)]
        return draw


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    merger = Ui_Merger()
    merger.show()
    sys.excepthook = except_hook
    sys.exit(app.exec_())
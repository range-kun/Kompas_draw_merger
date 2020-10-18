# -*- coding: utf-8 -*-
import sys
import os
import time
import kompas_api
import PyPDF2
import shutil
import fitz
from operator import itemgetter
from Widgets_class import MakeWidgets
from settings_window import SettingsWindow
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QThread, pyqtSignal


def except_hook(cls, exception, traceback):
    sys.__excepthook__(cls, exception, traceback)


class Ui_Merger(MakeWidgets):
    def __init__(self, parent=None):
        MakeWidgets.__init__(self, parent=None)
        self.kompas_ext = ['.cdw', '.spw']
        self.setFixedSize(720, 562)
        self.setupUi(self)
        self.directory = None
        self.current_progress = 0
        self.progress_step = 0
        self.settings_window = SettingsWindow()

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
        self.gridLayoutWidget.setGeometry(QtCore.QRect(10, 10, 700, 530))
        self.gridLayoutWidget.setObjectName("gridLayoutWidget")
        self.gridLayout = QtWidgets.QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)

        self.pushButton = self.make_button(text='Выделить все', font=font, command=self.select_all,
                                           size_policy=sizepolicy_button)
        self.gridLayout.addWidget(self.pushButton, 12, 0, 1, 1)

        self.pushButton_2 = self.make_button(text="Снять выделение", font=font, command=self.unselect_all,
                                             size_policy=sizepolicy_button)
        self.gridLayout.addWidget(self.pushButton_2, 12, 1, 1, 1)

        self.pushButton_3 = self.make_button(text="Добавить файл \nв список", font=font,
                                             command=self.add_file_to_list)
        self.gridLayout.addWidget(self.pushButton_3, 12, 2, 1, 1)

        self.pushButton_4 = self.make_button(text="Добавить папку \nв список", font=font,
                                             command=self.add_folder_to_list)
        self.gridLayout.addWidget(self.pushButton_4, 12, 3, 1, 1)

        self.pushButton_5 = self.make_button(text="Склеить файлы", font=font, command=self.merge_files_in_one)
        self.gridLayout.addWidget(self.pushButton_5, 15, 0, 1, 4)

        self.pushButton_6 = self.make_button(text="Выбор стартовой \nпапки", font=font,
                                             command=self.choose_initial_folder)
        self.gridLayout.addWidget(self.pushButton_6, 0, 2, 1, 1)

        self.pushButton_7 = self.make_button(text="Очистить спискок и выбор стартовой папки", font=font,
                                             command=self.clear_data)
        self.gridLayout.addWidget(self.pushButton_7, 1, 0, 1, 2)

        self.pushButton_8 = self.make_button(text="Дополнительные \nнастройки", font=font,
                                             command=self.show_settings)
        self.gridLayout.addWidget(self.pushButton_8, 0, 3, 1, 1)

        self.pushButton_9 = self.make_button(text='Обновить список файлов', font=font, command=self.refresh_settings)
        self.gridLayout.addWidget(self.pushButton_9, 1, 2, 1, 2)

        self.checkBox_3 = self.make_checkbox(font=font, text='Удалить папку с однодетальными \n'
                                                             'pdf-файлами по окончанию', activate=True)
        self.gridLayout.addWidget(self.checkBox_3, 13, 2, 1, 2)

        self.checkBox_4 = self.make_checkbox(font=font, text='С обходом всех папок'
                                                             'в выбраной папке', activate=True)
        self.gridLayout.addWidget(self.checkBox_4, 13, 0, 1, 2)

        self.label = self.make_text_edit(font=font, placeholder="Выберите папку с файлами в формате .spw или"
                                                                " .cdw", size_policy=sizepolicy)
        self.gridLayout.addWidget(self.label, 0, 0, 1, 2)

        self.listWidget = QtWidgets.QListWidget(self.gridLayoutWidget)
        self.gridLayout.addWidget(self.listWidget, 6, 0, 5, 4)

        self.progressBar = QtWidgets.QProgressBar(self.gridLayoutWidget)
        self.progressBar.setTextVisible(False)
        self.gridLayout.addWidget(self.progressBar, 16, 0, 1, 4)

        Merger.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(Merger)
        self.statusbar.setObjectName("statusbar")
        Merger.setStatusBar(self.statusbar)
        QtCore.QMetaObject.connectSlotsByName(Merger)

    def choose_initial_folder(self):
        directory = QtWidgets.QFileDialog.getExistingDirectory(self, "Выбрать папку", ".")
        if directory:
            self.directory = directory
            self.label.setText(self.directory)
            draw_list = self.get_all_files_in_folder()
            if draw_list:
                draw_list = self.filter_files(draw_list)
                if draw_list:
                    self.fill_list(draw_list=draw_list)
                    self.pushButton_5.setEnabled(True)

    def refresh_settings(self):
        folder = self.label.toPlainText()
        if folder:
            self.directory = folder
            self.listWidget.clear()
            draw_list = self.get_all_files_in_folder()
            if draw_list:
                draw_list = self.filter_files(draw_list)
                if draw_list:
                    self.fill_list(draw_list=draw_list)
        else:
            self.error('Укажите папку с файлами .cdw и .spw')

    def filter_files(self, draw_list=None):
        draw_list = draw_list or self.get_items_in_list(self.listWidget)
        filters = self.get_all_filters()
        if filters:
            self.calculate_step(len(draw_list), filter_only=True)
            draw_list = kompas_api.filter_draws(draw_list, **filters, instance=self)
            self.current_progress = 100
            self.progressBar.setValue(self.current_progress)
            if not draw_list:
                self.error('Нету файлов .cdw или .spw, в выбранной папке(ах) с указанными параметрами')
                return
        return draw_list

    def get_all_files_in_folder(self):
        draw_list = []
        except_folders_list = self.get_items_in_list(self.settings_window.listWidget)
        self.listWidget.clear()
        if self.checkBox_4.isChecked():
            for (this_dir, _, files_here) in os.walk(self.directory):
                if not os.path.basename(this_dir) in except_folders_list:
                    files = self.get_files_in_one_folder(this_dir)
                    draw_list += files
        else:
            draw_list = self.get_files_in_folder(self.directory)
        if draw_list:
            return draw_list
        else:
            self.error('Нету файлов .cdw или .spw, в выбраной папке(ах) с указанными параметрами')

    def get_files_in_one_folder(self, folder):
        # sorting in the way so specification sheet is the first if .cdw and .spw files has the same name
        draw_list = [os.path.splitext(i) for i in os.listdir(folder) if os.path.splitext(i)[1] in self.kompas_ext]
        draw_list = sorted([(i[0], '.adw' if i[1] == '.spw' else '.cdw') for i in draw_list], key=itemgetter(0, 1))
        draw_list = [os.path.join(folder, (i[0] + '.spw' if i[1] == '.adw' else i[0] + '.cdw')) for i in draw_list]
        draw_list = map(os.path.normpath, draw_list)
        return draw_list

    def merge_files_in_one(self):
        self.pushButton_5.setEnabled(False)
        if self.settings_window.checkBox.isChecked():
            files = self.filter_files()
        if not files:
            self.error('Нету файлов для сливания')
            return
        directory_pdf,  base_pdf_dir = self.create_folders()
        self.thread = MyBrandThread(files, directory_pdf, base_pdf_dir)
        self.thread.button_enable.connect(self.pushButton_5.setEnabled)
        self.thread.progress_bar.connect(self.progressBar.setValue)
        self.thread.errors.connect(self.error)
        self.thread.start()

    def create_folders(self):
        main_name = os.path.basename(self.directory)
        base_pdf_dir = rf'{self.directory}\pdf'
        pdf_file = r'%s\pdf\%s - 01 %s.pdf' % (self.directory, main_name, time.strftime("%d.%m.%Y"))
        directory_pdf = os.path.splitext(pdf_file)[0]

        if not os.path.exists(base_pdf_dir):
            os.makedirs(base_pdf_dir)
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
                            (self.directory, main_name, today_update, time.strftime("%d.%m.%Y"))
        os.makedirs(directory_pdf)
        return directory_pdf, base_pdf_dir

    def merge_pdf_files(self, directory):
        # Get list of files in directory
        files = sorted(os.listdir(directory), key=lambda fil: int(fil.split()[0]))
        pdf_document = PyPDF2.PdfFileMerger()
        for filename in files:
            pdf_document.append(fileobj=open(os.path.join(directory, filename), 'rb'))
        # check if sorting option activate
        if self.settings_window.checkBox_5.isChecked():
            input_pages = sorted([(i, i.pagedata['/MediaBox'][2:]) for i in merger.pages], key=itemgetter(1))
            pdf_document.pages = [i[0] for i in input_pages]
        pdf_file_location = os.path.join(os.path.dirname(directory), f'{os.path.basename(directory)}.pdf')
        with open(pdf_file_location, 'wb') as pdf:
            pdf_document.write(pdf)
        for file in pdf_document.inputs:
            file[0].close()
        return pdf_file_location

    def clear_data(self):
        self.directory = None
        self.label.clear()
        self.label.setPlaceholderText("Выберите папку с файлами \n в формате .cdw или .spw")
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

    def show_settings(self):
        self.settings_window.exec_()

    def calculate_step(self, number_of_files, filter_only=False):
        self.current_progress = 0
        number_of_operations = 0
        number_of_operations = self.settings_window.checkBox.isChecked()*1 or \
                 self.settings_window.checkBox_2.isChecked()*1 or \
                 self.settings_window.checkBox_3.isChecked()*1
        if not filter_only:
            number_of_operations += 2  # convert files and merger them
        self.progress_step = int(100 / (number_of_operations * number_of_files))

    def increase_step(self):
        self.current_progress += self.progress_step
        self.progressBar.setValue(self.current_progress)

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


class MyBrandThread(QThread):
    button_enable = pyqtSignal(bool)
    errors =pyqtSignal(str)
    progress_bar = pyqtSignal(float)

    def __init__(self, files, directory_pdf, base_pdf_dir):
        self.files = files
        self.base_pdf_dir = base_pdf_dir
        self.directory_pdf = directory_pdf
        self.progress_step = int(100/(2 * len(files)))
        self.current_progress = 0  # обнуляем прогресс в начале новой задачи
        QThread.__init__(self)

    def run(self):
        self.cdw_to_pdf(self.files, self.directory_pdf)
        pdf_file = self.merge_pdf_files(self.directory_pdf)
        if merger.checkBox_3.isChecked():
            shutil.rmtree(self.directory_pdf)
        os.system(f'explorer "{os.path.normpath(self.base_pdf_dir)}"')
        os.startfile(pdf_file)
        kompas_api.exit_kompas()
        self.progress_bar.emit(100-self.current_progress)
        self.button_enable.emit(True)

    def cdw_to_pdf(self, files, directory_pdf):
        kompas_api7_module, application, const = kompas_api.get_kompas_api7()
        kompas6_api5_module, kompas_object, kompas6_constants = kompas_api.get_kompas_api5()
        doc_app , iConverter, _ = kompas_api.get_kompas_settings(application, kompas_object)
        number = 0
        self.progress_bar.emit(self.current_progress)
        for file in files:
            number += 1
            iConverter.Convert(file, directory_pdf + "\\" +
                               f'{number} ' + os.path.basename(file) + ".pdf", 0, False)
            self.increase_step()

    def merge_pdf_files(self, directory):
        files = sorted(os.listdir(directory), key=lambda fil: int(fil.split()[0]))
        merger_pdf = PyPDF2.PdfFileMerger()
        for filename in files:
            merger_pdf.append(fileobj=open(os.path.join(directory, filename), 'rb'))
            self.increase_step()
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

    def increase_step(self):
        self.current_progress += self.progress_step
        self.progress_bar.emit(self.current_progress)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    merger = Ui_Merger()
    merger.show()
    sys.excepthook = except_hook
    sys.exit(app.exec_())


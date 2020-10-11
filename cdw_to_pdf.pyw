# -*- coding: utf-8 -*-
import sys
import os
import time
import kompas_api
import PyPDF2
import shutil
from operator import itemgetter
from Widgets_class import MakeWidgets
from _datetime import datetime
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QThread, pyqtSignal


def except_hook(cls, exception, traceback):
    sys.__excepthook__(cls, exception, traceback)

class SettingsWindow(QtWidgets.QWidget):
    def __init__(self):
        QtWidgets.QWidget.__init__(self)
        self.construct_class = MakeWidgets()
        self.date_today = [int(i) for i in str(datetime.date(datetime.now())).split('-')]
        self.setupUi(self)

    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(640, 461)

        self.gridLayoutWidget = QtWidgets.QWidget(Form)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(10, 10, 624, 441))
        self.gridLayoutWidget.setObjectName("gridLayoutWidget")
        self.gridLayout = QtWidgets.QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)

        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)

        font_1 = QtGui.QFont()
        font_1.setFamily("Arial")
        font_1.setPointSize(11)

        font_2 = QtGui.QFont()
        font_2.setFamily("MS Shell Dlg 2")
        font_2.setPointSize(12)

        datePolicy = \
            QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.MinimumExpanding)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Ignored, QtWidgets.QSizePolicy.Ignored)

        self.checkBox = self.construct_class.make_checkbox(
            font=font, text='Файлы с датой только за указанный период', command=self.select_date,
            parent=self.gridLayoutWidget)
        self.gridLayout.addWidget(self.checkBox, 1, 0, 1, 1)

        self.dateEdit = self.construct_class.make_date(date_par=self.date_today, font=font,
                                                       parent=self.gridLayoutWidget)
        self.dateEdit.setEnabled(False)
        self.gridLayout.addWidget(self.dateEdit, 1, 2, 1, 1)

        self.dateEdit_2 = self.construct_class.make_date(date_par=self.date_today, font=font,
                                                         parent=self.gridLayoutWidget)
        self.dateEdit_2.setEnabled(False)
        self.dateEdit_2.setSizePolicy(datePolicy)
        self.gridLayout.addWidget(self.dateEdit_2, 1, 1, 1, 1)


        self.checkBox_2 = self.construct_class.make_checkbox(
            font=font, text='С указанной фамилией разработчика', command=self.select_constructor,
            parent=self.gridLayoutWidget)
        self.gridLayout.addWidget(self.checkBox_2,  2, 0, 3, 1)

        self.radio_button = self.construct_class.make_radio_button(text='Фамилия из списка', font=font_1,
                                                                   parent=self.gridLayoutWidget,
                                                                   command=self.constructor_name_option)
        self.radio_button.setEnabled(False)
        self.gridLayout.addWidget(self.radio_button, 2, 1, 1, 1)

        self.radio_button_2 = self.construct_class.make_radio_button(text='Другая', font=font_1,
                                                                     parent=self.gridLayoutWidget,
                                                                     command=self.constructor_name_option)
        self.radio_button_2.setEnabled(False)
        self.gridLayout.addWidget(self.radio_button_2, 2, 2, 1, 1)

        self.btngroup = QtWidgets.QButtonGroup()
        self.btngroup.addButton(self.radio_button)
        self.btngroup.addButton(self.radio_button_2)

        self.comboBox = self.construct_class.make_combobox(font=font_2, parent=self.gridLayoutWidget)
        self.comboBox.setEnabled(False)
        self.gridLayout.addWidget(self.comboBox, 3, 1, 1, 1)

        self.lineEdit = self.construct_class.make_line_edit(parent=self.gridLayoutWidget, font=font)
        self.lineEdit.setSizePolicy(sizePolicy)
        self.lineEdit.setEnabled(False)
        self.gridLayout.addWidget(self.lineEdit, 3, 2, 1, 1)


        self.checkBox_3 = self.construct_class.make_checkbox(
            font=font, text='С указанной фамилией проверяющего', command=self.select_checker,
            parent=self.gridLayoutWidget)
        self.gridLayout.addWidget(self.checkBox_3, 5, 0, 2, 1)

        self.radio_button_3 = self.construct_class.make_radio_button(text='Фамилия из списка', font=font_1,
                                                                     parent=self.gridLayoutWidget,
                                                                     command=self.checker_name_option)
        self.radio_button_3.setEnabled(False)
        self.gridLayout.addWidget(self.radio_button_3, 5, 1, 1, 1)

        self.radio_button_4 = self.construct_class.make_radio_button(text='Другая', font=font_1,
                                                                     parent=self.gridLayoutWidget,
                                                                     command=self.checker_name_option)
        self.radio_button_4.setEnabled(False)
        self.gridLayout.addWidget(self.radio_button_4, 5, 2, 1, 1)

        self.btngroup_2 = QtWidgets.QButtonGroup()
        self.btngroup_2.addButton(self.radio_button_3)
        self.btngroup_2.addButton(self.radio_button_4)

        self.comboBox_2 = self.construct_class.make_combobox(font=font_2, parent=self.gridLayoutWidget)
        self.comboBox_2.setEnabled(False)
        self.gridLayout.addWidget(self.comboBox_2, 6, 1, 1, 1)

        self.lineEdit_2 = self.construct_class.make_line_edit(parent=self.gridLayoutWidget, font=font_1)
        self.lineEdit_2.setSizePolicy(sizePolicy)
        self.lineEdit_2.setEnabled(False)
        self.gridLayout.addWidget(self.lineEdit_2, 6, 2, 1, 1)


        self.checkBox_4 = self.construct_class.make_checkbox(text='Добавить водяной знак',
                                                             font=font, parent=self.gridLayoutWidget,
                                                             command=self.water_mark_option,
                                                             activate=True)
        self.gridLayout.addWidget(self.checkBox_4, 7, 0, 2, 1)

        self.push_button = self.construct_class.make_button(text='Стандартный', font=font,
                                                            parent=self.gridLayoutWidget,
                                                            command=self.add_default_watermark
                                                            )
        self.gridLayout.addWidget(self.push_button, 7, 1, 1, 1)

        self.push_button_2 = self.construct_class.make_button(text='Свое изображение', font=font,
                                                              parent=self.gridLayoutWidget,
                                                              command=self.add_custom_watermark)
        self.gridLayout.addWidget(self.push_button_2, 7, 2, 1, 1)

        self.checkBox_5 = self.construct_class.make_checkbox(text='Отсортировать файлы по формату',
                                                             font=font, parent=self.gridLayoutWidget)
        self.gridLayout.addWidget(self.checkBox_5, 9, 0, 1, 1)

        self.lineEdit_3 = self.construct_class.make_line_edit(parent=self.gridLayoutWidget, font=font_1)
        self.gridLayout.addWidget(self.lineEdit_3, 8, 1, 1, 2)


        self.label = self.construct_class.make_label(text='Исключить следующие папки:', font=font,
                                                     parent=self.gridLayoutWidget)
        self.gridLayout.addWidget(self.label, 10, 0, 1, 3)

        self.listWidget = QtWidgets.QListWidget(self.gridLayoutWidget)
        self.listWidget.setFont(font)
        self.gridLayout.addWidget(self.listWidget, 11, 0, 2, 2)


        self.pushButton_3 = self.construct_class.make_button(text='Удалить выбранную\n папку',
                                                             parent=self.gridLayoutWidget,
                                                             font=font, size_policy=datePolicy)
        self.gridLayout.addWidget(self.pushButton_3,  11, 2, 1, 1)

        self.pushButton_4 = self.construct_class.make_button(text='Добавить папку',
                                                             parent=self.gridLayoutWidget,
                                                             font=font, size_policy=datePolicy)
        self.gridLayout.addWidget(self.pushButton_4, 12, 2, 1, 1)

        self.pushButton_5 = self.construct_class.make_button(text='Сбросить настройки',
                                                             parent=self.gridLayoutWidget,
                                                             font=font, command=self.set_default_settings)
        self.gridLayout.addWidget(self.pushButton_5, 13, 0, 1, 3)

        QtCore.QMetaObject.connectSlotsByName(Form)

    def select_date(self):
        self.dateEdit.setEnabled(self.checkBox.isChecked())
        self.dateEdit_2.setEnabled(self.checkBox.isChecked())

    def select_constructor(self):
        self.comboBox.setEnabled(self.checkBox_2.isChecked())
        self.lineEdit.setEnabled(self.checkBox_2.isChecked())
        self.radio_button.setEnabled(self.checkBox_2.isChecked())
        self.radio_button_2.setEnabled(self.checkBox_2.isChecked())
        if self.checkBox_2.isChecked():
            self.btngroup.setExclusive(True)
            self.radio_button.setChecked(self.checkBox_2.isChecked())
        else:
            self.btngroup.setExclusive(False)
            self.radio_button.setChecked(self.checkBox_2.isChecked())
            self.radio_button_2.setChecked(self.checkBox_2.isChecked())


    def select_checker(self):
        self.comboBox_2.setEnabled(self.checkBox_3.isChecked())
        self.lineEdit_2.setEnabled(self.checkBox_3.isChecked())
        self.radio_button_3.setEnabled(self.checkBox_3.isChecked())
        self.radio_button_4.setEnabled(self.checkBox_3.isChecked())
        if self.checkBox_3.isChecked():
            self.btngroup_2.setExclusive(True)
            self.radio_button_3.setChecked(self.checkBox_3.isChecked())
        else:
            self.btngroup_2.setExclusive(False)
            self.radio_button_3.setChecked(self.checkBox_3.isChecked())
            self.radio_button_4.setChecked(self.checkBox_3.isChecked())

    def constructor_name_option(self):
        choosed_combo_box = self.sender().text() == 'Фамилия из списка'
        self.comboBox.setEnabled(choosed_combo_box)
        self.lineEdit.setEnabled(not choosed_combo_box)

    def checker_name_option(self):
        choosed_combo_box_2 = self.sender().text() == 'Фамилия из списка'
        self.comboBox_2.setEnabled(choosed_combo_box_2)
        self.lineEdit_2.setEnabled(not choosed_combo_box_2)

    def water_mark_option(self):
        self.push_button.setEnabled(self.checkBox_4.isChecked())
        self.push_button_2.setEnabled(self.checkBox_4.isChecked())
        self.lineEdit_3.setEnabled(self.checkBox_4.isChecked())

    def add_default_watermark(self):
        pass

    def add_custom_watermark(self):
        filename = QtWidgets.QFileDialog.getOpenFileName(self, "Выбрать файл", ".",
                                                          "jpg(*.jpg);;"
                                                          "png(*.png);;"
                                                          "bmp(*.bmp);;")[0]
        if filename:
            self.lineEdit_3.setText(filename)

    def set_default_settings(self):
        pass

class Ui_Merger(MakeWidgets):
    def __init__(self, parent=None):
        MakeWidgets.__init__(self, parent=None)
        self.kompas_ext = ['.cdw', '.spw']
        self.setFixedSize(720, 562)
        self.setupUi(self)
        self.directory = None
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

        self.pushButton_9 = self.make_button(text='Обновить список файлов', font=font)
        self.gridLayout.addWidget(self.pushButton_9, 1, 2, 1, 2)

        self.checkBox_3 = self.make_checkbox(font=font, text='Удалить папку с однодетальными \n'
                                                             'pdf-файлами по окончанию', activate=True)
        self.gridLayout.addWidget(self.checkBox_3, 13, 2, 1, 2)

        self.checkBox_4 = self.make_checkbox(font=font, text='С обходом всех папок'
                                                             'в указанной папке', activate=True)
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
        self.directory = QtWidgets.QFileDialog.getExistingDirectory(self, "Выбрать папку", ".")
        draw_list = []
        if self.directory:
            self.listWidget.clear()
            if self.checkBox_4.isChecked():
                for (this_dir, _, files_here) in os.walk(self.directory):
                    if 'Былое' not in this_dir and 'Проработка' not in this_dir:
                        files = self.get_files_in_folder(this_dir)
                        draw_list += files
            else:
                draw_list = self.get_files_in_folder(self.directory)
        if draw_list:
            if self.settings_window.checkBox.isChecked():
                draw_list = self.filter_files(draw_list)
            if draw_list:
                self.label.setText(self.directory)
                self.fill_list(draw_list)

    def filter_files(self, draw_list):
        date_1 = self.settings_window.dateEdit.dateTime().toSecsSinceEpoch()
        date_2 = self.settings_window.dateEdit_2.dateTime().toSecsSinceEpoch()
        return kompas_api.filter_by_date(draw_list, date_1, date_2)

    def get_files_in_folder(self, folder):
        # sorting in the way so specification sheet is the first if .cdw and .spw files has the same name
        draw_list = [os.path.splitext(i) for i in os.listdir(folder) if os.path.splitext(i)[1] in self.kompas_ext]
        draw_list = sorted([(i[0], '.adw' if i[1] == '.spw' else '.cdw') for i in draw_list], key=itemgetter(0, 1))
        draw_list = [os.path.join(folder, (i[0] + '.spw' if i[1] == '.adw' else i[0] + '.cdw')) for i in draw_list]
        draw_list = map(os.path.normpath, draw_list)
        return draw_list

    def merge_files_in_one(self):
        self.pushButton_5.setEnabled(False)
        files = [str(self.listWidget.item(i).text()) for i
                 in range(self.listWidget.count()) if self.listWidget.item(i).checkState()]
        if self.checkBox.isChecked():
            files = self.filter_files(files)
        if not files:
            self.error('Нету файлов для сливания')
            return
        directory_pdf,  base_pdf_dir = self.create_folders()
        self.thread = MyBrandThread(files, directory_pdf, base_pdf_dir)
        self.thread.button_enable.connect(self.pushButton_5.setEnabled)
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
            self.fill_list(filename)

    def add_folder_to_list(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Выбрать папку", ".")
        draw_list = []
        if folder:
            draw_list = self.get_files_in_folder(folder)
        if draw_list:
            self.fill_list(draw_list)

    def show_settings(self):
        self.settings_window.show()


class MyBrandThread(QThread):
    button_enable = pyqtSignal(bool)

    def __init__(self, files, directory_pdf, base_pdf_dir):
        self.files = files
        self.base_pdf_dir = base_pdf_dir
        self.directory_pdf = directory_pdf
        QThread.__init__(self)

    def run(self):
        self.cdw_to_pdf(self.files, self.directory_pdf)
        pdf_file = self.merge_pdf_files(self.directory_pdf)
        if merger.checkBox_3.isChecked():
            shutil.rmtree(self.directory_pdf)
        os.system(f'explorer "{os.path.normpath(self.base_pdf_dir)}"')
        os.startfile(pdf_file)
        self.button_enable.emit(True)

    @staticmethod
    def cdw_to_pdf(files, directory_pdf):
        kompas_api7_module, application, const = kompas_api.get_kompas_api7()
        kompas6_api5_module, kompas_object, kompas6_constants = kompas_api.get_kompas_api5()
        doc_app = application.Application
        iConverter = doc_app.Converter(kompas_object.ksSystemPath(5) + r"\Pdf2d.dll")
        number = 0
        for file in files:
            number += 1
            iConverter.Convert(file, directory_pdf + "\\" +
                               f'{number} ' + os.path.basename(file) + ".pdf", 0, False)

    def merge_pdf_files(self, directory):
        # Получаем список файлов в переменную files
        files = sorted(os.listdir(directory), key=lambda fil: int(fil.split()[0]))
        merger_pdf = PyPDF2.PdfFileMerger()
        for filename in files:
            merger_pdf.append(fileobj=open(os.path.join(directory, filename), 'rb'))
        if merger.settings_window.checkBox_5.isChecked():
            input_pages = sorted([(i, i.pagedata['/MediaBox'][2:]) for i in merger.pages], key=itemgetter(1))
            merger_pdf.pages = [i[0] for i in input_pages]
        pdf_file = os.path.join(os.path.dirname(directory), f'{os.path.basename(directory)}.pdf')
        with open(pdf_file, 'wb') as pdf:
            merger_pdf.write(pdf)
        for file in merger_pdf.inputs:
            file[0].close()
        return pdf_file


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    merger = Ui_Merger()
    merger.show()
    sys.excepthook = except_hook
    try:
        sys.exit(app.exec_())
    except:
        kompas_api.exit_kompas()
        sys.exit(app.exec_())

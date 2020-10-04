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
#from threading import Thread


class Ui_Merger(MakeWidgets):
    def __init__(self, parent=None):
        MakeWidgets.__init__(self, parent=None)
        self.kompas_ext = ['.cdw', '.spw']
        self.date_today = [int(i) for i in str(datetime.date(datetime.now())).split('-')]
        self.setFixedSize(707, 552)
        self.setupUi(self)
        self.directory = None

    def setupUi(self, Merger):
        Merger.setObjectName("Merger")
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        Merger.setFont(font)

        self.centralwidget = QtWidgets.QWidget(Merger)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(10, 10, 693, 521))
        self.gridLayoutWidget.setObjectName("gridLayoutWidget")
        self.gridLayout = QtWidgets.QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)

        self.pushButton = self.make_button(text='Выделить \nвсе', font=font, command=self.select_all)
        self.gridLayout.addWidget(self.pushButton, 12, 0, 1, 1)

        self.pushButton_2 = self.make_button(text="Снять \nвыделение", font=font, command=self.unselect_all)
        self.gridLayout.addWidget(self.pushButton_2, 12, 1, 1, 1)

        self.pushButton_3 = self.make_button(text="Добавить файл \nв список", font=font,
                                             command=self.add_file_to_list)
        self.gridLayout.addWidget(self.pushButton_3, 12, 2, 1, 1)

        self.pushButton_4 = self.make_button(text="Добавить папку \nв список", font=font,
                                             command=self.add_folder_to_list)
        self.gridLayout.addWidget(self.pushButton_4, 12, 3, 1, 1)

        self.pushButton_5 = self.make_button(text="Склеить файлы", font=font, command=self.merge_files_in_one)
        self.pushButton_5.setEnabled(False)
        self.gridLayout.addWidget(self.pushButton_5, 16, 0, 1, 4)

        self.pushButton_6 = self.make_button(text="Выбор стартовой папки", font=font, command=self.choose_initial_folder)
        self.gridLayout.addWidget(self.pushButton_6, 1, 0, 1, 2)

        self.pushButton_7 = self.make_button(text="Очистить спискок и выбор стартовой папки", font=font,
                                             command=self.clear_data)
        self.gridLayout.addWidget(self.pushButton_7, 1, 2, 1, 2)

        self.dateEdit = self.make_date(date_par=self.date_today, font=font)
        self.dateEdit.setEnabled(False)
        self.gridLayout.addWidget(self.dateEdit, 13, 3, 1, 1)

        self.dateEdit_2 = self.make_date(date_par=self.date_today, font=font)
        self.dateEdit_2.setEnabled(False)
        self.gridLayout.addWidget(self.dateEdit_2, 13, 2, 1, 1)

        self.checkBox = self.make_checkbox(font=font, text='Файлы с датой только за указанный период',
                                           command=self.select_date)
        self.gridLayout.addWidget(self.checkBox, 13, 0, 1, 2)

        self.checkBox_2 = self.make_checkbox(font=font, text='Отсортировать файлы по формату')
        self.gridLayout.addWidget(self.checkBox_2, 14, 0, 1, 2)

        self.checkBox_3 = self.make_checkbox(font=font, text='Удалить папку с однодетальными \n'
                                                             'pdf-файлами по окончанию', activate=True)
        self.gridLayout.addWidget(self.checkBox_3, 14, 2, 1, 2)

        self.checkBox_4 = self.make_checkbox(font=font, text='С обходом всех папок\n'
                                                             'в указанной папке', activate=True)
        self.gridLayout.addWidget(self.checkBox_4, 0, 2, 1, 2)

        self.label = self.make_label(font=font, text="Выберите папку с файлами \n "
                                                     "в формате .cdw или .spw")
        self.gridLayout.addWidget(self.label, 0, 0, 1, 2)

        self.listWidget = QtWidgets.QListWidget(self.gridLayoutWidget)
        self.gridLayout.addWidget(self.listWidget, 6, 0, 5, 4)
        Merger.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(Merger)
        self.statusbar.setObjectName("statusbar")
        Merger.setStatusBar(self.statusbar)
        QtCore.QMetaObject.connectSlotsByName(Merger)

    def select_date(self):
        self.dateEdit.setEnabled(self.checkBox.isChecked())
        self.dateEdit_2.setEnabled(self.checkBox.isChecked())

    def choose_initial_folder(self):
        self.directory = QtWidgets.QFileDialog.getExistingDirectory(self, "Выбрать папку", ".")
        draw_list = []
        if self.directory:
            self.listWidget.clear()
            if self.checkBox_4.isChecked():
                for (this_dir, _, files_here) in os.walk(self.directory):
                    if not this_dir.endswith('Былое') and 'Проработка' not in this_dir:
                        files = self.get_files_in_folder(this_dir)
                        draw_list += files
            else:
                draw_list = self.get_files_in_folder(self.directory)
        if draw_list:
            if self.checkBox.isChecked():
                draw_list = self.filter_files(draw_list)
            if draw_list:
                self.label.setText(self.directory)
                self.fill_list(draw_list)

    def filter_files(self, draw_list):
        date_1 = self.dateEdit.dateTime().toSecsSinceEpoch()
        date_2 = self.dateEdit_2.dateTime().toSecsSinceEpoch()
        return kompas_api.filter_by_date(draw_list, date_1, date_2)

    def get_files_in_folder(self, folder):
        draw_list = [os.path.join(folder, i) for i in os.listdir(folder) if os.path.splitext(i)[1] in self.kompas_ext]
        draw_list = sorted(draw_list, key=lambda fil: os.path.splitext(fil)[1], reverse=True)
        map(os.path.normpath, draw_list)
        return draw_list

    def merge_files_in_one(self):
        files = [str(self.listWidget.item(i).text()) for i
                         in range(self.listWidget.count()) if self.listWidget.item(i).checkState()]
        directory_pdf = self.create_folders()
        if self.checkBox.isChecked():
            files = self.filter_files(files)
        kompas_api.cdw_to_pdf(files, directory_pdf)
        pdf_file = self.merge_pdf(directory_pdf)
        if self.checkBox_3.isChecked():
            shutil.rmtree(directory_pdf)
        os.startfile(pdf_file)

    def create_folders(self):
        main_name = [os.path.splitext(i)[0] for i in os.listdir(self.directory) if os.path.splitext(i)[1] == '.cdw'][0]
        if not main_name:
            main_name = \
            [os.path.splitext(i)[0] for i in os.listdir(self.directory) if os.path.splitext(i)[1] == '.spw'][0]
        base_pdf_dir = f'{self.directory}\pdf'
        pdf_file = r'%s\pdf\%s - 01 %s.pdf' % (self.directory, main_name, time.strftime("%d.%m.%Y"))
        directory_pdf = os.path.splitext(pdf_file)[0]

        if not os.path.exists(base_pdf_dir):
            os.makedirs(base_pdf_dir)
        if os.path.exists(pdf_file) or os.path.exists(directory_pdf):
            created_files = max([int(i.split()[-2]) for i in os.listdir(base_pdf_dir)
                                 if i.endswith(f'{time.strftime("%d.%m.%Y.pdf")}')], default=0)
            created_folders = max([int(i.split()[-2]) for i in os.listdir(base_pdf_dir)
                                   if i.endswith(f'{time.strftime("%d.%m.%Y")}')], default=0)
            today_update = created_files if created_files >= created_folders else created_folders
            today_update = str(today_update + 1) if today_update > 8 else '0' + str(today_update + 1)
            directory_pdf = r'%s\pdf\%s - %s %s' % (self.directory, main_name, today_update, time.strftime("%d.%m.%Y"))
        os.makedirs(directory_pdf)
        return directory_pdf

    def merge_pdf(self, directory):
        # Получаем список файлов в переменную files
        files = sorted(os.listdir(directory), key=lambda fil: int(fil.split()[0]))
        merger = PyPDF2.PdfFileMerger()
        for filename in files:
            merger.append(fileobj=open(os.path.join(directory, filename), 'rb'))
        if self.checkBox_2.isChecked():
            input_pages = sorted([(i, i.pagedata['/MediaBox'][2:]) for i in merger.pages], key=itemgetter(1))
            merger.pages = [i[0] for i in input_pages]
        pdf_file = os.path.join(os.path.dirname(directory), f'{os.path.basename(directory)}.pdf')
        with open(pdf_file, 'wb') as pdf:
            merger.write(pdf)
        for file in merger.inputs:
            file[0].close()
        return pdf_file

    def clear_data(self):
        self.directory = None
        self.label.setText("Выберите папку с файлами \n в формате .cdw или .spw")
        self.listWidget.clear()
        self.pushButton_5.setEnabled(False)

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
                                                         "Чертж(*.cdw);;Спецификация(*.spw)")[0]]
        if filename:
            self.fill_list(filename)

    def add_folder_to_list(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Выбрать папку", ".")
        draw_list = []
        if folder:
            draw_list = self.get_files_in_folder(folder)
        if draw_list:
            self.fill_list(draw_list)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    merger = Ui_Merger()
    merger.show()
    sys.exit(app.exec_())

from _datetime import datetime
from PyQt5 import QtCore, QtGui, QtWidgets
from Widgets_class import MakeWidgets
import json
import os


class SettingsWindow(QtWidgets.QDialog):
    def __init__(self):
        QtWidgets.QDialog.__init__(self)
        self.construct_class = MakeWidgets()
        self.date_today = [int(i) for i in str(datetime.date(datetime.now())).split('-')]
        self.setupUi(self)
        self.except_folders_list = []
        self.constructor_list = []
        self.checker_list = []
        self.add_default_watermark = False
        self.watermark_path = ''
        self.watermark_position = []
        self.sort_files = False
        self.load_settings = self.get_settings()
        if self.load_settings:
            self.apply_settings()

    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(675, 441)

        self.gridLayoutWidget = QtWidgets.QWidget(Form)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(10, 10, 660, 421))
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
        check_box_policy = QtWidgets.QSizePolicy\
            (QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Fixed)

        self.checkBox = self.construct_class.make_checkbox(
            font=font, text='C датой только за указанный период', command=self.select_date,
            parent=self.gridLayoutWidget)
        self.checkBox.setSizePolicy(check_box_policy)
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
                                                                   command=self.constructor_name_radio_option)
        self.radio_button.setEnabled(False)
        self.gridLayout.addWidget(self.radio_button, 2, 1, 1, 1)

        self.radio_button_2 = self.construct_class.make_radio_button(text='Другая', font=font_1,
                                                                     parent=self.gridLayoutWidget,
                                                                     command=self.constructor_name_radio_option)
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
                                                                     command=self.checker_name_radio_option)
        self.radio_button_3.setEnabled(False)
        self.gridLayout.addWidget(self.radio_button_3, 5, 1, 1, 1)

        self.radio_button_4 = self.construct_class.make_radio_button(text='Другая', font=font_1,
                                                                     parent=self.gridLayoutWidget,
                                                                     command=self.checker_name_radio_option)
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
                                                             command=self.select_watermark)
        self.gridLayout.addWidget(self.checkBox_4, 7, 0, 2, 1)

        self.radio_button_5 = self.construct_class.make_radio_button(text='Стандартный', font=font,
                                                                  parent=self.gridLayoutWidget,
                                                                  command=self.watermark_option
                                                                  )
        self.gridLayout.addWidget(self.radio_button_5, 7, 1, 1, 1)
        self.radio_button_5.setChecked(True)

        self.radio_button_6 = self.construct_class.make_radio_button(text='Свое изображение', font=font,
                                                                    parent=self.gridLayoutWidget,
                                                                    command=self.watermark_option
                                                                    )
        self.gridLayout.addWidget(self.radio_button_6, 7, 2, 1, 1)
        self.btngroup_3 = QtWidgets.QButtonGroup()
        self.btngroup_3.addButton(self.radio_button_5)
        self.btngroup_3.addButton(self.radio_button_6)

        self.checkBox_5 = self.construct_class.make_checkbox(text='Отсортировать файлы по формату',
                                                             font=font, parent=self.gridLayoutWidget)
        self.gridLayout.addWidget(self.checkBox_5, 9, 0, 1, 1)

        self.lineEdit_3 = self.construct_class.make_line_edit(parent=self.gridLayoutWidget, font=font_1)
        self.gridLayout.addWidget(self.lineEdit_3, 8, 1, 1, 2)
        self.lineEdit_3.setEnabled(False)


        self.label = self.construct_class.make_label(text='Исключить следующие папки:', font=font,
                                                     parent=self.gridLayoutWidget)
        self.gridLayout.addWidget(self.label, 10, 0, 1, 3)

        self.listWidget = QtWidgets.QListWidget(self.gridLayoutWidget)
        self.listWidget.setFont(font)
        self.gridLayout.addWidget(self.listWidget, 11, 0, 2, 2)


        self.pushButton = self.construct_class.make_button(text='Добавить папку',
                                                           parent=self.gridLayoutWidget,
                                                           font=font, size_policy=datePolicy,
                                                           command=self.add_folder)
        self.gridLayout.addWidget(self.pushButton,  11, 2, 1, 1)

        self.pushButton_2 = self.construct_class.make_button(text='Удалить выбранную\n папку',
                                                             parent=self.gridLayoutWidget,
                                                             font=font, size_policy=datePolicy,
                                                             command=self.delete_folder
                                                             )
        self.gridLayout.addWidget(self.pushButton_2, 12, 2, 1, 1)

        self.pushButton_3 = self.construct_class.make_button(text='Сбросить настройки',
                                                             parent=self.gridLayoutWidget,
                                                             font=font, command=self.set_default_settings)
        self.gridLayout.addWidget(self.pushButton_3, 13, 0, 1, 3)

        QtCore.QMetaObject.connectSlotsByName(Form)

    def get_settings(self):
        try:
            file_dir = os.path.dirname(os.path.abspath(__file__))
            if os.stat(file_dir+r'\settings.json').st_size > 0:
                data = open('settings.json', encoding='utf-8-sig')
        except OSError:
            miss_file = self.construct_class.error('Файл settings.txt \n отсутсвует')
            miss_file.exec_()
            return
        try:
            obj = json.load(data)
        except json.decoder.JSONDecodeError as e:
            self.construct_class.error('В Файл settings.txt \n присутсвуют ошибки \n синтаксиса json')
            return
        data.close()
        return obj

    def apply_settings(self):
        settings = ('except_folders_list', 'constructor_list', 'checker_list',
                    'add_default_watermark', 'watermark_path', 'watermark_position', 'sort_files')
        settings_methods = [(self.construct_class.fill_list, (), {'widget_list': self.listWidget,
                                                                  'draw_list': self.except_folders_list}),
                            (self.construct_class.fill_combo_box, (self.constructor_list, self.comboBox), {}),
                            (self.construct_class.fill_combo_box, (self.checker_list, self.comboBox_2), {}),
                            ('check_watermark', lambda: self.checkBox_4.setChecked(self.add_default_watermark), {}),
                            ('activate_watermark', lambda: self.select_watermark(), {}),
                            ('activate_sorting', lambda: self.checkBox_5.setChecked(self.sort_files), {}),
                            ]
        for key in settings:
            if type(self.__dict__[key]) == list:
                self.__dict__[key][:] = self.load_settings.get(key, None)
            else:
                self.__dict__[key] = self.load_settings.get(key, None)
        for method, args, kwargs in settings_methods:
            try:
                if type(method) == str:
                    args()
                else:
                    method(*args, **kwargs)
            except Exception as e:
                pass

    def select_date(self):
        self.dateEdit.setEnabled(self.checkBox.isChecked())
        self.dateEdit_2.setEnabled(self.checkBox.isChecked())

    def select_constructor(self):
        self.comboBox.setEnabled(self.checkBox_2.isChecked())
        self.radio_button.setEnabled(self.checkBox_2.isChecked())
        self.radio_button_2.setEnabled(self.checkBox_2.isChecked())
        self.lineEdit.setEnabled(False)
        if self.checkBox_2.isChecked():
            self.btngroup.setExclusive(True)
            self.radio_button.setChecked(self.checkBox_2.isChecked())
        else:
            self.btngroup.setExclusive(False)
            self.radio_button.setChecked(self.checkBox_2.isChecked())
            self.radio_button_2.setChecked(self.checkBox_2.isChecked())

    def constructor_name_radio_option(self):
        choosed_combo_box = self.sender().text() == 'Фамилия из списка'
        self.comboBox.setEnabled(choosed_combo_box)
        self.lineEdit.setEnabled(not choosed_combo_box)

    def select_checker(self):
        self.comboBox_2.setEnabled(self.checkBox_3.isChecked())
        self.radio_button_3.setEnabled(self.checkBox_3.isChecked())
        self.radio_button_4.setEnabled(self.checkBox_3.isChecked())
        self.lineEdit_2.setEnabled(False)
        if self.checkBox_3.isChecked():
            self.btngroup_2.setExclusive(True)
            self.radio_button_3.setChecked(self.checkBox_3.isChecked())
        else:
            self.btngroup_2.setExclusive(False)
            self.radio_button_3.setChecked(self.checkBox_3.isChecked())
            self.radio_button_4.setChecked(self.checkBox_3.isChecked())

    def checker_name_radio_option(self):
        choosed_combo_box_2 = self.sender().text() == 'Фамилия из списка'
        self.comboBox_2.setEnabled(choosed_combo_box_2)
        self.lineEdit_2.setEnabled(not choosed_combo_box_2)

    def select_watermark(self):
        self.radio_button_5.setEnabled(self.checkBox_4.isChecked())
        self.radio_button_6.setEnabled(self.checkBox_4.isChecked())
        self.lineEdit_3.setEnabled(False)
        if self.checkBox_4.isChecked():
            self.btngroup_3.setExclusive(True)
            self.radio_button_5.setChecked(self.checkBox_4.isChecked())
            if self.watermark_path:
                path = os.path.abspath(self.watermark_path)
                if os.path.exists(path):
                    self.watermark_path = path
                    self.lineEdit_3.setText(self.watermark_path)
                else:
                    self.watermark_path = 'Стандартный путь из настроек не существует'
                    self.lineEdit_3.setText(self.watermark_path)
        else:
            self.btngroup_3.setExclusive(False)
            self.radio_button_5.setChecked(self.checkBox_4.isChecked())
            self.radio_button_6.setChecked(self.checkBox_4.isChecked())

    def watermark_option(self):
        choosed_default_watermark = self.sender().text() == 'Стандартный'
        self.lineEdit_3.setEnabled(not choosed_default_watermark)
        if choosed_default_watermark:
            self.lineEdit_3.clear()
            self.lineEdit_3.setText(self.watermark_path)
        else:
            self.lineEdit_3.clear()
            filename = QtWidgets.QFileDialog.getOpenFileName(self, "Выбрать файл", ".",
                                                             "jpg(*.jpg);;"
                                                             "png(*.png);;"
                                                             "bmp(*.bmp);;")[0]
            if filename:
                self.lineEdit_3.setText(filename)


    def add_folder(self):
        folder_name, ok = QtWidgets.QInputDialog.getText(self, "Дилог ввода текста", "Введите название папки")
        if ok:
            self.construct_class.fill_list(draw_list=[folder_name], widget_list=self.listWidget)

    def delete_folder(self):
        self.construct_class.remove_item(widget_list=self.listWidget)


    def clear_settings(self):
        self.checkBox.setChecked(False)
        self.select_date()
        self.construct_class.set_date(self.date_today, self.dateEdit)
        self.construct_class.set_date(self.date_today, self.dateEdit_2)
        self.checkBox_2.setChecked(False)
        self.comboBox.clear()
        self.lineEdit.clear()
        self.select_constructor()
        self.checkBox_3.setChecked(False)
        self.comboBox_2.clear()
        self.lineEdit_2.clear()
        self.select_checker()
        self.checkBox_4.setChecked(False)
        self.select_watermark()
        self.checkBox_5.setChecked(False)
        self.listWidget.clear()

    def set_default_settings(self):
        self.clear_settings()
        self.load_settings = self.get_settings()
        if self.load_settings:
            self.apply_settings()
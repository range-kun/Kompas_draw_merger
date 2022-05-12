from __future__ import annotations

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
        self.auto_save_folder = False
        self.watermark_path = ''
        self.watermark_position = []
        self.sort_files = False
        self.load_settings = self.get_settings()
        if self.load_settings:
            self.apply_settings()

    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(675, 461)

        self.gridLayoutWidget = QtWidgets.QWidget(Form)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(10, 10, 660, 441))
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

        self.checkBox_5 = self.construct_class.make_checkbox(text='Разбить на файлы по размерам',
                                                             font=font, parent=self.gridLayoutWidget)
        self.gridLayout.addWidget(self.checkBox_5, 9, 0, 1, 1)

        self.lineEdit_3 = self.construct_class.make_line_edit(parent=self.gridLayoutWidget, font=font_1)
        self.gridLayout.addWidget(self.lineEdit_3, 8, 1, 1, 2)
        self.lineEdit_3.setEnabled(False)

        self.radio_button_7 = self.construct_class.make_radio_button(text='Указать папку сохранения вручную', font=font_2,
                                                                   parent=self.gridLayoutWidget)
        self.gridLayout.addWidget(self.radio_button_7, 10, 0, 1, 1)
        self.radio_button_7.setChecked(True)

        self.radio_button_8 = self.construct_class.make_radio_button(text='Выбрать папку автоматически', font=font_2,
                                                                     parent=self.gridLayoutWidget)
        self.gridLayout.addWidget(self.radio_button_8, 10, 1, 1, 3)

        self.btngroup_4 = QtWidgets.QButtonGroup()
        self.btngroup_4.addButton(self.radio_button_7)
        self.btngroup_4.addButton(self.radio_button_8)

        self.label = self.construct_class.make_label(text='Исключить следующие папки:', font=font,
                                                     parent=self.gridLayoutWidget)
        self.gridLayout.addWidget(self.label, 11, 0, 1, 3)


        self.listWidget = QtWidgets.QListWidget(self.gridLayoutWidget)
        self.listWidget.setFont(font)
        self.gridLayout.addWidget(self.listWidget, 12, 0, 2, 2)


        self.pushButton = self.construct_class.make_button(text='Добавить папку',
                                                           parent=self.gridLayoutWidget,
                                                           font=font, size_policy=datePolicy,
                                                           command=self.add_folder)
        self.gridLayout.addWidget(self.pushButton,  12, 2, 1, 1)

        self.pushButton_2 = self.construct_class.make_button(text='Удалить выбранную\n папку',
                                                             parent=self.gridLayoutWidget,
                                                             font=font, size_policy=datePolicy,
                                                             command=self.delete_folder
                                                             )
        self.gridLayout.addWidget(self.pushButton_2, 13, 2, 1, 1)

        self.pushButton_3 = self.construct_class.make_button(text='Сбросить настройки',
                                                             parent=self.gridLayoutWidget,
                                                             font=font, command=self.set_default_settings)
        self.gridLayout.addWidget(self.pushButton_3, 14, 1, 1, 3)

        self.pushButton_4 = self.construct_class.make_button(text='Ок', parent=self.gridLayoutWidget,
                                                             font=font, command=self.close)
        self.gridLayout.addWidget(self.pushButton_4, 14, 0, 1, 1)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def get_settings(self):
        try:
            file_dir = os.path.dirname(os.path.abspath(__file__))
            if os.stat(file_dir+r'\settings.json').st_size > 0:
                data = open('settings.json', encoding='utf-8-sig')
        except OSError:
            self.construct_class.error('Файл settings.txt \n отсутсвует')
            return
        try:
            obj = json.load(data)
        except json.decoder.JSONDecodeError as e:
            self.construct_class.error('В Файл settings.txt \n присутсвуют ошибки \n синтаксиса json')
            data.close()
            return
        data.close()
        return obj

    def apply_settings(self):
        settings = ('except_folders_list', 'constructor_list', 'checker_list',
                    'add_default_watermark', 'watermark_path',
                    'watermark_position', 'sort_files', 'auto_save_folder')
        settings_methods = [(self.construct_class.fill_list, (), {'widget_list': self.listWidget,
                                                                  'draw_list': self.except_folders_list}),
                            (self.construct_class.fill_combo_box, (self.constructor_list, self.comboBox), {}),
                            (self.construct_class.fill_combo_box, (self.checker_list, self.comboBox_2), {}),
                            ('check_watermark', lambda: self.checkBox_4.setChecked(self.add_default_watermark), {}),
                            ('activate_watermark', lambda: self.select_watermark(), {}),
                            ('activate_sorting', lambda: self.checkBox_5.setChecked(self.sort_files), {}),
                            ('auto_save_folder', lambda: self.select_auto_folder(), {})
                            ]
        for key in settings:
            if type(self.__dict__[key]) == list:
                try:
                    self.__dict__[key][:] = self.load_settings.get(key, None)
                except TypeError:
                    self.__dict__[key][:] = []
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
                current_dir_path = os.path.dirname(os.path.abspath(__file__))
                path = os.path.abspath('bdt_stamp.png')
                if not os.path.exists(path):
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

    def select_auto_folder(self):
        if self.auto_save_folder:
            self.radio_button_8.setChecked(True)
        else:
            self.radio_button_7.setChecked(True)

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


class RadioButtonsWindow(QtWidgets.QDialog):

    def __init__(self, executions: list[str]):
        QtWidgets.QDialog.__init__(self)
        self.construct_class = MakeWidgets()
        self.executions = executions
        self.layout = QtWidgets.QGridLayout()
        self.setLayout(self.layout)
        self.setWindowTitle("Вы указали групповую спецификацию")

        self.font = QtGui.QFont()
        self.font.setFamily("Arial")
        self.font.setPointSize(12)
        self.set_label()
        self.set_options()
        self.add_buttons()

        self.radio_state = None

    def on_radio_clicked(self):
        radio_button = self.sender()
        if radio_button.isChecked():
            self.radio_state = radio_button.option

    def on_button_clicked(self):
        source = self.sender()

        if source.text() == 'Отмена':
            self.radio_state = None
        self.close()

    def set_label(self):
        plain_text = QtWidgets.QLabel('Выберите исполнение для слияния:')
        plain_text.setFont(self.font)
        self.layout.addWidget(plain_text, 0, 0, 1, 3)

    def set_options(self):
        for index, option in enumerate(self.executions):
            if type(option) != str:
                continue
            radiobutton = QtWidgets.QRadioButton(option)
            radiobutton.option = option
            radiobutton.toggled.connect(self.on_radio_clicked)
            radiobutton.setFont(self.font)
            self.layout.addWidget(radiobutton, 2, index)

    def add_buttons(self):
        button_layout = QtWidgets.QHBoxLayout()

        ok_button = QtWidgets.QPushButton('OК')
        ok_button.setFont(self.font)
        ok_button.clicked.connect(self.on_button_clicked)
        button_layout.addWidget(ok_button)

        cancel_button = QtWidgets.QPushButton('Отмена')
        cancel_button.setFont(self.font)
        cancel_button.clicked.connect(self.on_button_clicked)
        button_layout.addWidget(cancel_button)

        self.layout.addLayout(button_layout, 3, 0, 1, 4)


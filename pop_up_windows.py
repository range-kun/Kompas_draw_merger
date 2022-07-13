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

        self.grid_layout_widget = QtWidgets.QWidget(Form)
        self.grid_layout_widget.setGeometry(QtCore.QRect(10, 10, 660, 441))
        self.grid_layout_widget.setObjectName("gridLayoutWidget")

        self.grid_layout = QtWidgets.QGridLayout(self.grid_layout_widget)
        self.grid_layout.setContentsMargins(0, 0, 0, 0)

        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)

        font_1 = QtGui.QFont()
        font_1.setFamily("Arial")
        font_1.setPointSize(11)

        font_2 = QtGui.QFont()
        font_2.setFamily("MS Shell Dlg 2")
        font_2.setPointSize(12)

        date_policy = \
            QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.MinimumExpanding)
        size_policy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Ignored, QtWidgets.QSizePolicy.Ignored)
        check_box_policy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Fixed)

        # FILTER BY DATE SECTION
        self.filter_by_date_check_box = self.construct_class.make_checkbox(
            font=font,
            text='C датой только за указанный период',
            command=self.select_date,
            parent=self.grid_layout_widget
        )
        self.filter_by_date_check_box.setSizePolicy(check_box_policy)
        self.grid_layout.addWidget(self.filter_by_date_check_box, 1, 0, 1, 1)

        self.first_date_input = self.construct_class.make_date(
            date_par=self.date_today,
            font=font,
            parent=self.grid_layout_widget
        )
        self.first_date_input.setEnabled(False)
        self.grid_layout.addWidget(self.first_date_input, 1, 2, 1, 1)

        self.last_date_input = self.construct_class.make_date(
            date_par=self.date_today,
            font=font,
            parent=self.grid_layout_widget
        )
        self.last_date_input.setEnabled(False)
        self.last_date_input.setSizePolicy(date_policy)
        self.grid_layout.addWidget(self.last_date_input, 1, 1, 1, 1)

        # FILTER BY CONSTRUCTOR NAME SECTION
        self.filter_by_draw_designer = self.construct_class.make_checkbox(
            font=font,
            text='С указанной фамилией разработчика',
            command=self.select_constructor,
            parent=self.grid_layout_widget
        )
        self.grid_layout.addWidget(self.filter_by_draw_designer,  2, 0, 3, 1)

        self.constructor_from_combobox_radio_button = self.construct_class.make_radio_button(
            text='Фамилия из списка', font=font_1,
            parent=self.grid_layout_widget,
            command=self.constructor_name_radio_option
        )
        self.constructor_from_combobox_radio_button.setEnabled(False)
        self.grid_layout.addWidget(self.constructor_from_combobox_radio_button, 2, 1, 1, 1)

        self.input_random_constructor_radio_button = self.construct_class.make_radio_button(
            text='Другая', font=font_1,
            parent=self.grid_layout_widget,
            command=self.constructor_name_radio_option
        )
        self.input_random_constructor_radio_button.setEnabled(False)
        self.grid_layout.addWidget(self.input_random_constructor_radio_button, 2, 2, 1, 1)

        self.select_constructor_group = QtWidgets.QButtonGroup()
        self.select_constructor_group.addButton(self.constructor_from_combobox_radio_button)
        self.select_constructor_group.addButton(self.input_random_constructor_radio_button)

        self.constructor_combo_box = self.construct_class.make_combobox(font=font_2, parent=self.grid_layout_widget)
        self.constructor_combo_box.setEnabled(False)
        self.grid_layout.addWidget(self.constructor_combo_box, 3, 1, 1, 1)

        self.random_constructor_line_edit = self.construct_class.make_line_edit(
            parent=self.grid_layout_widget,
            font=font
        )
        self.random_constructor_line_edit.setSizePolicy(size_policy)
        self.random_constructor_line_edit.setEnabled(False)
        self.grid_layout.addWidget(self.random_constructor_line_edit, 3, 2, 1, 1)

        # FILTER BY CHECKER NAME SECTION
        self.select_checker_name_radio_button = self.construct_class.make_checkbox(
            font=font,
            text='С указанной фамилией проверяющего',
            command=self.select_checker,
            parent=self.grid_layout_widget
        )
        self.grid_layout.addWidget(self.select_checker_name_radio_button, 5, 0, 2, 1)

        self.checker_from_combobox_radio_button = self.construct_class.make_radio_button(
            text='Фамилия из списка', font=font_1,
            parent=self.grid_layout_widget,
            command=self.checker_name_radio_option
        )
        self.checker_from_combobox_radio_button.setEnabled(False)
        self.grid_layout.addWidget(self.checker_from_combobox_radio_button, 5, 1, 1, 1)

        self.checker_from_line_radio_button = self.construct_class.make_radio_button(
            text='Другая', font=font_1,
            parent=self.grid_layout_widget,
            command=self.checker_name_radio_option
        )
        self.checker_from_line_radio_button.setEnabled(False)
        self.grid_layout.addWidget(self.checker_from_line_radio_button, 5, 2, 1, 1)

        self.checker_btn_group = QtWidgets.QButtonGroup()
        self.checker_btn_group.addButton(self.checker_from_combobox_radio_button)
        self.checker_btn_group.addButton(self.checker_from_line_radio_button)

        self.checker_combo_box = self.construct_class.make_combobox(font=font_2, parent=self.grid_layout_widget)
        self.checker_combo_box.setEnabled(False)
        self.grid_layout.addWidget(self.checker_combo_box, 6, 1, 1, 1)

        self.random_checker_line_input = self.construct_class.make_line_edit(parent=self.grid_layout_widget, font=font_1)
        self.random_checker_line_input.setSizePolicy(size_policy)
        self.random_checker_line_input.setEnabled(False)
        self.grid_layout.addWidget(self.random_checker_line_input, 6, 2, 1, 1)

        # ADD WATERMARK SECTION
        self.add_water_mark_check_box = self.construct_class.make_checkbox(
            text='Добавить водяной знак',
            font=font, parent=self.grid_layout_widget,
            command=self.select_watermark
        )
        self.grid_layout.addWidget(self.add_water_mark_check_box, 7, 0, 2, 1)

        self.default_watermark_path_radio_button = self.construct_class.make_radio_button(
            text='Стандартный', font=font,
            parent=self.grid_layout_widget,
            command=self.watermark_option
        )
        self.grid_layout.addWidget(self.default_watermark_path_radio_button, 7, 1, 1, 1)
        self.default_watermark_path_radio_button.setChecked(True)

        self.custom_watermark_path_radio_button = self.construct_class.make_radio_button(
            text='Свое изображение', font=font,
            parent=self.grid_layout_widget,
            command=self.watermark_option
        )

        self.custom_watermark_path_edit_line = self.construct_class.make_line_edit(
            parent=self.grid_layout_widget,
            font=font_1
        )
        self.grid_layout.addWidget(self.custom_watermark_path_edit_line, 8, 1, 1, 2)
        self.custom_watermark_path_edit_line.setEnabled(False)
        self.grid_layout.addWidget(self.custom_watermark_path_radio_button, 7, 2, 1, 1)

        self.watermark_path_btn_group = QtWidgets.QButtonGroup()
        self.watermark_path_btn_group.addButton(self.default_watermark_path_radio_button)
        self.watermark_path_btn_group.addButton(self.custom_watermark_path_radio_button)

        # SPLIT FILE BY SIZE SECTION
        self.split_files_by_size_checkbox = self.construct_class.make_checkbox(
            text='Разбить на файлы по размерам',
            font=font,
            parent=self.grid_layout_widget
        )
        self.grid_layout.addWidget(self.split_files_by_size_checkbox, 9, 0, 1, 1)

        # CHOOSE DEFAULT SAVE FILE PATH SECTION
        self.manually_choose_save_folder_radio_button = self.construct_class.make_radio_button(
            text='Указать папку сохранения вручную',
            font=font_2,
            parent=self.grid_layout_widget
        )
        self.grid_layout.addWidget(self.manually_choose_save_folder_radio_button, 10, 0, 1, 1)
        self.manually_choose_save_folder_radio_button.setChecked(True)

        self.auto_choose_save_folder_radio_button = self.construct_class.make_radio_button(
            text='Выбрать папку автоматически',
            font=font_2,
            parent=self.grid_layout_widget
        )
        self.grid_layout.addWidget(self.auto_choose_save_folder_radio_button, 10, 1, 1, 3)

        self.save_folder_btngroup = QtWidgets.QButtonGroup()
        self.save_folder_btngroup.addButton(self.manually_choose_save_folder_radio_button)
        self.save_folder_btngroup.addButton(self.auto_choose_save_folder_radio_button)

        # EXCLUDE FOLDERS FROM SEARCH SETTING
        self.exclude_folder_label = self.construct_class.make_label(
            text='Исключить следующие папки:',
            font=font,
            parent=self.grid_layout_widget
        )
        self.grid_layout.addWidget(self.exclude_folder_label, 11, 0, 1, 3)

        self.exclude_folder_list_widget = QtWidgets.QListWidget(self.grid_layout_widget)
        self.exclude_folder_list_widget.setFont(font)
        self.grid_layout.addWidget(self.exclude_folder_list_widget, 12, 0, 2, 2)

        self.add_exclude_folder_button = self.construct_class.make_button(
            text='Добавить папку',
            parent=self.grid_layout_widget,
            font=font, size_policy=date_policy,
            command=self.add_folder
        )
        self.grid_layout.addWidget(self.add_exclude_folder_button,  12, 2, 1, 1)

        self.delete_exclude_folder_button = self.construct_class.make_button(
            text='Удалить выбранную\n папку',
            parent=self.grid_layout_widget,
            font=font, size_policy=date_policy,
            command=self.delete_folder
        )
        self.grid_layout.addWidget(self.delete_exclude_folder_button, 13, 2, 1, 1)

        # SETUP WINDOW SETTINGS
        self.reset_settings_button = self.construct_class.make_button(
            text='Сбросить настройки',
            parent=self.grid_layout_widget,
            font=font,
            command=self.set_default_settings
        )
        self.grid_layout.addWidget(self.reset_settings_button, 14, 1, 1, 3)

        self.close_window_button = self.construct_class.make_button(
            text='Ок',
            parent=self.grid_layout_widget,
            font=font,
            command=self.close
        )
        self.grid_layout.addWidget(self.close_window_button, 14, 0, 1, 1)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def get_settings(self):
        try:
            file_dir = os.path.dirname(os.path.abspath(__file__))
            if os.stat(file_dir+r'\settings.json').st_size > 0:
                data = open('settings.json', encoding='utf-8-sig')
        except OSError:
            self.construct_class.send_error('Файл settings.txt \n отсутсвует')
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
        settings_methods = [
            (self.construct_class.fill_list, (),
             {'widget_list': self.exclude_folder_list_widget, 'draw_list': self.except_folders_list
              }),
            (self.construct_class.fill_combo_box, (self.constructor_list, self.constructor_combo_box), {}),
            (self.construct_class.fill_combo_box, (self.checker_list, self.checker_combo_box), {}),
            ('check_watermark', lambda: self.add_water_mark_check_box.setChecked(self.add_default_watermark), {}),
            ('activate_watermark', lambda: self.select_watermark(), {}),
            ('activate_sorting', lambda: self.split_files_by_size_checkbox.setChecked(self.sort_files), {}),
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
                print(f"Ошибка чтения настроек {e}")
                pass

    def select_date(self):
        self.first_date_input.setEnabled(self.filter_by_date_check_box.isChecked())
        self.last_date_input.setEnabled(self.filter_by_date_check_box.isChecked())

    def select_constructor(self):
        self.constructor_combo_box.setEnabled(self.filter_by_draw_designer.isChecked())
        self.constructor_from_combobox_radio_button.setEnabled(self.filter_by_draw_designer.isChecked())
        self.input_random_constructor_radio_button.setEnabled(self.filter_by_draw_designer.isChecked())
        self.random_constructor_line_edit.setEnabled(False)
        if self.filter_by_draw_designer.isChecked():
            self.select_constructor_group.setExclusive(True)
            self.constructor_from_combobox_radio_button.setChecked(self.filter_by_draw_designer.isChecked())
        else:
            self.select_constructor_group.setExclusive(False)
            self.constructor_from_combobox_radio_button.setChecked(self.filter_by_draw_designer.isChecked())
            self.input_random_constructor_radio_button.setChecked(self.filter_by_draw_designer.isChecked())

    def constructor_name_radio_option(self):
        choosed_combo_box = self.sender().text() == 'Фамилия из списка'
        self.constructor_combo_box.setEnabled(choosed_combo_box)
        self.random_constructor_line_edit.setEnabled(not choosed_combo_box)

    def select_checker(self):
        self.checker_combo_box.setEnabled(self.select_checker_name_radio_button.isChecked())
        self.checker_from_combobox_radio_button.setEnabled(self.select_checker_name_radio_button.isChecked())
        self.checker_from_line_radio_button.setEnabled(self.select_checker_name_radio_button.isChecked())
        self.random_checker_line_input.setEnabled(False)
        if self.select_checker_name_radio_button.isChecked():
            self.checker_btn_group.setExclusive(True)
            self.checker_from_combobox_radio_button.setChecked(self.select_checker_name_radio_button.isChecked())
        else:
            self.checker_btn_group.setExclusive(False)
            self.checker_from_combobox_radio_button.setChecked(self.select_checker_name_radio_button.isChecked())
            self.checker_from_line_radio_button.setChecked(self.select_checker_name_radio_button.isChecked())

    def checker_name_radio_option(self):
        choosed_combo_box_2 = self.sender().text() == 'Фамилия из списка'
        self.checker_combo_box.setEnabled(choosed_combo_box_2)
        self.random_checker_line_input.setEnabled(not choosed_combo_box_2)

    def select_watermark(self):
        self.default_watermark_path_radio_button.setEnabled(self.add_water_mark_check_box.isChecked())
        self.custom_watermark_path_radio_button.setEnabled(self.add_water_mark_check_box.isChecked())
        self.custom_watermark_path_edit_line.setEnabled(False)
        if self.add_water_mark_check_box.isChecked():
            self.watermark_path_btn_group.setExclusive(True)
            self.default_watermark_path_radio_button.setChecked(self.add_water_mark_check_box.isChecked())

            if self.watermark_path:
                current_dir_path = os.path.dirname(os.path.abspath(__file__))
                path = os.path.abspath('bdt_stamp.png')
                if not os.path.exists(path):
                    path = os.path.abspath(self.watermark_path)

                if os.path.exists(path):
                    self.watermark_path = path
                    self.custom_watermark_path_edit_line.setText(self.watermark_path)
                else:
                    self.watermark_path = 'Стандартный путь из настроек не существует'
                    self.custom_watermark_path_edit_line.setText(self.watermark_path)
        else:
            self.watermark_path_btn_group.setExclusive(False)
            self.default_watermark_path_radio_button.setChecked(self.add_water_mark_check_box.isChecked())
            self.custom_watermark_path_radio_button.setChecked(self.add_water_mark_check_box.isChecked())

    def select_auto_folder(self):
        if self.auto_save_folder:
            self.auto_choose_save_folder_radio_button.setChecked(True)
        else:
            self.manually_choose_save_folder_radio_button.setChecked(True)

    def watermark_option(self):
        choosed_default_watermark = self.sender().text() == 'Стандартный'
        self.custom_watermark_path_edit_line.setEnabled(not choosed_default_watermark)
        if choosed_default_watermark:
            self.custom_watermark_path_edit_line.clear()
            self.custom_watermark_path_edit_line.setText(self.watermark_path)
        else:
            self.custom_watermark_path_edit_line.clear()
            filename = QtWidgets.QFileDialog.getOpenFileName(
             self, "Выбрать файл", ".",
             "jpg(*.jpg);;"
             "png(*.png);;"
             "bmp(*.bmp);;"
            )[0]
            if filename:
                self.custom_watermark_path_edit_line.setText(filename)

    def add_folder(self):
        folder_name, ok = QtWidgets.QInputDialog.getText(self, "Дилог ввода текста", "Введите название папки")
        if ok:
            self.construct_class.fill_list(draw_list=[folder_name], widget_list=self.exclude_folder_list_widget)

    def delete_folder(self):
        self.construct_class.remove_item(widget_list=self.exclude_folder_list_widget)

    def clear_settings(self):
        self.filter_by_date_check_box.setChecked(False)
        self.select_date()
        self.construct_class.set_date(self.date_today, self.first_date_input)
        self.construct_class.set_date(self.date_today, self.last_date_input)
        self.filter_by_draw_designer.setChecked(False)
        self.constructor_combo_box.clear()
        self.random_constructor_line_edit.clear()
        self.select_constructor()
        self.select_checker_name_radio_button.setChecked(False)
        self.checker_combo_box.clear()
        self.random_checker_line_input.clear()
        self.select_checker()
        self.add_water_mark_check_box.setChecked(False)
        self.select_watermark()
        self.split_files_by_size_checkbox.setChecked(False)
        self.exclude_folder_list_widget.clear()

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


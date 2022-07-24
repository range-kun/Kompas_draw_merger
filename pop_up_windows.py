from __future__ import annotations

from typing import NewType
import json
import os
from _datetime import datetime

from pydantic import BaseModel
from PyQt5 import QtCore, QtGui, QtWidgets

from Widgets_class import MakeWidgets, ExcludeFolderListWidget

DEFAULT_WATERMARK_PATH = 'bdt_stamp.png'
FILE_NOT_EXISTS_MESSAGE = "Путь к файлу из настроек не существует"
FilePath = NewType('FilePath', str)


class FilterWidgetPositions(BaseModel):
    check_box_position: list[int]
    combobox_position: list[int]
    input_line_position: list[int]
    combobox_radio_button_position: list[int]
    input_radio_button_position: list[int]


class UserSettings(BaseModel):
    except_folders_list: list[str]
    constructor_list: list[str]
    checker_list: list[str]
    sortament_list: list[str]
    watermark_position: list[int]
    add_default_watermark: bool = True
    split_file_by_size: bool = False
    auto_save_folder: bool = False
    watermark_path: FilePath


class SettingsWindow(QtWidgets.QDialog):
    def __init__(self):
        QtWidgets.QDialog.__init__(self)

        self.font = QtGui.QFont()
        self.font.setFamily("Arial")
        self.font.setPointSize(12)

        self.font_1 = QtGui.QFont()
        self.font_1.setFamily("Arial")
        self.font_1.setPointSize(11)

        self.font_2 = QtGui.QFont()
        self.font_2.setFamily("MS Shell Dlg 2")
        self.font_2.setPointSize(12)

        self.date_policy = \
            QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.MinimumExpanding)
        self.size_policy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Ignored, QtWidgets.QSizePolicy.Ignored)
        self.check_box_policy = \
            QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Fixed)

        self.construct_class = MakeWidgets()
        self.date_today = [int(i) for i in str(datetime.date(datetime.now())).split('-')]
        self.setup_ui()

        self.watermark_position = []
        self._watermark_path = ''
        self.apply_user_settings()

    def setup_ui(self):
        self.setObjectName("Form")
        self.resize(675, 520)
        self.setWindowTitle("Настройки")

        self.grid_layout_widget = QtWidgets.QWidget(self)
        self.grid_layout_widget.setGeometry(QtCore.QRect(10, 10, 660, 500))
        self.grid_layout_widget.setObjectName("gridLayoutWidget")

        self.grid_layout = QtWidgets.QGridLayout(self.grid_layout_widget)
        self.grid_layout.setContentsMargins(0, 0, 0, 0)

        self.setup_filter_by_date_section()

        self.setup_filter_by_constructor_section()
        self.setup_filter_by_checker_section()

        self.setup_filter_by_gauge_section()
        self.setup_watermark_section()
        self.setup_split_by_size_section()
        self.setup_save_file_path_section()
        self.setup_exclude_folders_section()
        self.setup_window_settings()

    def setup_filter_by_date_section(self):
        self.filter_by_date_check_box = self.construct_class.make_checkbox(
            font=self.font,
            text='C датой только за указанный период',
            command=self.switch_date_input_filter,
            parent=self.grid_layout_widget
        )
        self.filter_by_date_check_box.setSizePolicy(self.check_box_policy)
        self.grid_layout.addWidget(self.filter_by_date_check_box, 1, 0, 1, 1)

        self.first_date_input = self.construct_class.make_date(
            font=self.font,
            parent=self.grid_layout_widget
        )
        self.grid_layout.addWidget(self.first_date_input, 1, 2, 1, 1)

        self.last_date_input = self.construct_class.make_date(
            font=self.font,
            parent=self.grid_layout_widget
        )
        self.last_date_input.setSizePolicy(self.date_policy)
        self.grid_layout.addWidget(self.last_date_input, 1, 1, 1, 1)

    def setup_filter_by_constructor_section(self):
        constructor_positions = FilterWidgetPositions(
            check_box_position=[2, 0, 3, 1],
            combobox_position=[3, 1, 1, 1],
            input_line_position=[3, 2, 1, 1],
            combobox_radio_button_position=[2, 1, 1, 1],
            input_radio_button_position=[2, 2, 1, 1]
        )
        self.constructor_filter = FilterSection(
            'С указанной фамилией разработчиков',
            'Фамилии из списка',
            constructor_positions,
            self.grid_layout,
            self.grid_layout_widget,
            self.font,
            self.font_1,
            self.font_2,
            self.size_policy,
            self.construct_class
        )

    def setup_filter_by_checker_section(self):
        checker_positions = FilterWidgetPositions(
            check_box_position=[5, 0, 2, 1],
            combobox_position=[6, 1, 1, 1],
            input_line_position=[6, 2, 1, 1],
            combobox_radio_button_position=[5, 1, 1, 1],
            input_radio_button_position=[5, 2, 1, 1]
        )
        self.checker_filter = FilterSection(
            'С указанной фамилией проверяющего',
            'Фамилии из списка',
            checker_positions,
            self.grid_layout,
            self.grid_layout_widget,
            self.font,
            self.font_1,
            self.font_2,
            self.size_policy,
            self.construct_class
        )

    def setup_filter_by_gauge_section(self):
        gauge_positions = FilterWidgetPositions(
            check_box_position=[7, 0, 2, 1],
            combobox_position=[8, 1, 1, 1],
            input_line_position=[8, 2, 1, 1],
            combobox_radio_button_position=[7, 1, 1, 1],
            input_radio_button_position=[7, 2, 1, 1]
        )
        self.gauger_filter = FilterSection(
            'С указанным сортаментом',
            'Сортамент из списка',
            gauge_positions,
            self.grid_layout,
            self.grid_layout_widget,
            self.font,
            self.font_1,
            self.font_2,
            self.size_policy,
            self.construct_class
        )

    def setup_watermark_section(self):
        self.add_water_mark_check_box = self.construct_class.make_checkbox(
            text='Добавить водяной знак',
            font=self.font, parent=self.grid_layout_widget,
            command=self.switch_watermark_group
        )
        self.grid_layout.addWidget(self.add_water_mark_check_box, 9, 0, 2, 1)

        self.default_watermark_path_radio_button = self.construct_class.make_radio_button(
            text='Стандартный', font=self.font,
            parent=self.grid_layout_widget,
            command=self.watermark_option
        )
        self.grid_layout.addWidget(self.default_watermark_path_radio_button, 9, 1, 1, 1)

        self.custom_watermark_path_radio_button = self.construct_class.make_radio_button(
            text='Свое изображение', font=self.font,
            parent=self.grid_layout_widget,
            command=self.watermark_option
        )
        self.grid_layout.addWidget(self.custom_watermark_path_radio_button, 9, 2, 1, 1)

        self.custom_watermark_path_edit_line = self.construct_class.make_line_edit(
            parent=self.grid_layout_widget,
            font=self.font_1
        )
        self.grid_layout.addWidget(self.custom_watermark_path_edit_line, 10, 1, 1, 2)

        self.watermark_path_btn_group = QtWidgets.QButtonGroup()
        self.watermark_path_btn_group.addButton(self.default_watermark_path_radio_button)
        self.watermark_path_btn_group.addButton(self.custom_watermark_path_radio_button)

    def setup_split_by_size_section(self):
        self.split_files_by_size_checkbox = self.construct_class.make_checkbox(
            text='Разбить на файлы по размерам',
            font=self.font,
            parent=self.grid_layout_widget
        )
        self.grid_layout.addWidget(self.split_files_by_size_checkbox, 11, 0, 1, 1)

    def setup_save_file_path_section(self):
        self.manually_choose_save_folder_radio_button = self.construct_class.make_radio_button(
            text='Указать папку сохранения вручную',
            font=self.font_2,
            parent=self.grid_layout_widget
        )
        self.grid_layout.addWidget(self.manually_choose_save_folder_radio_button, 12, 0, 1, 1)

        self.auto_choose_save_folder_radio_button = self.construct_class.make_radio_button(
            text='Выбрать папку автоматически',
            font=self.font_2,
            parent=self.grid_layout_widget
        )
        self.grid_layout.addWidget(self.auto_choose_save_folder_radio_button, 12, 1, 1, 3)

        self.save_folder_btngroup = QtWidgets.QButtonGroup()
        self.save_folder_btngroup.addButton(self.manually_choose_save_folder_radio_button)
        self.save_folder_btngroup.addButton(self.auto_choose_save_folder_radio_button)

    def setup_exclude_folders_section(self):
        self.exclude_folder_label = self.construct_class.make_label(
            text='Исключить следующие папки:',
            font=self.font,
            parent=self.grid_layout_widget
        )
        self.grid_layout.addWidget(self.exclude_folder_label, 13, 0, 1, 3)

        self.exclude_folder_list_widget = ExcludeFolderListWidget(self.grid_layout_widget)
        self.exclude_folder_list_widget.setFont(self.font)
        self.grid_layout.addWidget(self.exclude_folder_list_widget, 14, 0, 2, 2)

        self.add_exclude_folder_button = self.construct_class.make_button(
            text='Добавить папку',
            parent=self.grid_layout_widget,
            font=self.font,
            size_policy=self.date_policy,
            command=self.exclude_folder_list_widget.add_folder
        )
        self.grid_layout.addWidget(self.add_exclude_folder_button,  14, 2, 1, 1)

        self.delete_exclude_folder_button = self.construct_class.make_button(
            text='Удалить выбранную\n папку',
            parent=self.grid_layout_widget,
            font=self.font,
            size_policy=self.date_policy,
            command=self.exclude_folder_list_widget.remove_item
        )
        self.grid_layout.addWidget(self.delete_exclude_folder_button, 15, 2, 1, 1)

    def setup_window_settings(self):
        self.reset_settings_button = self.construct_class.make_button(
            text='Сбросить настройки',
            parent=self.grid_layout_widget,
            font=self.font,
            command=self.set_default_settings
        )
        self.grid_layout.addWidget(self.reset_settings_button, 16, 1, 1, 3)

        self.close_window_button = self.construct_class.make_button(
            text='Ок',
            parent=self.grid_layout_widget,
            font=self.font,
            command=self.close
        )
        self.grid_layout.addWidget(self.close_window_button, 16, 0, 1, 1)
        QtCore.QMetaObject.connectSlotsByName(self)

    def apply_user_settings(self):
        loaded_user_settings = self.get_settings()
        self._watermark_path = loaded_user_settings.watermark_path
        if loaded_user_settings:
            self.fill_widgets_with_settings(loaded_user_settings)

    def get_settings(self) -> UserSettings | None:
        try:
            file_dir = os.path.dirname(os.path.abspath(__file__))
            if os.stat(file_dir+r'\settings.json').st_size > 0:
                with open('settings.json', encoding='utf-8-sig') as data:
                    obj = json.load(data)
                user_settings = UserSettings.parse_obj(obj)
        except OSError:
            self.construct_class.send_error('Файл settings.txt \n отсутсвует')
            return
        except json.decoder.JSONDecodeError:
            self.construct_class.error('В Файл settings.txt \n присутсвуют ошибки \n синтаксиса json')
            return

        return user_settings

    def fill_widgets_with_settings(self, user_settings: UserSettings):
        self.construct_class.set_date(self.date_today, self.first_date_input)
        self.construct_class.set_date(self.date_today, self.last_date_input)
        self.switch_date_input_filter()

        self.exclude_folder_list_widget.fill_list(draw_list=user_settings.except_folders_list)

        self.construct_class.fill_combo_box(user_settings.constructor_list, self.constructor_filter.data_combo_box)
        self.constructor_filter.switch_filter_input()

        self.construct_class.fill_combo_box(user_settings.checker_list, self.checker_filter.data_combo_box)
        self.checker_filter.switch_filter_input()

        self.construct_class.fill_combo_box(user_settings.sortament_list, self.gauger_filter.data_combo_box)
        self.gauger_filter.switch_filter_input()

        self.add_water_mark_check_box.setChecked(user_settings.add_default_watermark)
        self.switch_watermark_group()
        self.set_user_watermark_path()

        self.split_files_by_size_checkbox.setChecked(user_settings.split_file_by_size)
        self.select_auto_folder(user_settings.auto_save_folder)

    def switch_date_input_filter(self):
        self.first_date_input.setEnabled(self.filter_by_date_check_box.isChecked())
        self.last_date_input.setEnabled(self.filter_by_date_check_box.isChecked())

    def switch_watermark_group(self):
        self.default_watermark_path_radio_button.setEnabled(self.add_water_mark_check_box.isChecked())
        self.custom_watermark_path_radio_button.setEnabled(self.add_water_mark_check_box.isChecked())
        self.custom_watermark_path_edit_line.setEnabled(False)

        if self.add_water_mark_check_box.isChecked():
            self.watermark_path_btn_group.setExclusive(True)
            self.default_watermark_path_radio_button.setChecked(self.add_water_mark_check_box.isChecked())
        else:
            self.watermark_path_btn_group.setExclusive(False)
            self.default_watermark_path_radio_button.setChecked(self.add_water_mark_check_box.isChecked())
            self.custom_watermark_path_radio_button.setChecked(self.add_water_mark_check_box.isChecked())

    def get_user_watermark_path(self) -> str:
        path = os.path.abspath(DEFAULT_WATERMARK_PATH)
        if not os.path.exists(path):
            path = os.path.abspath(self._watermark_path)
        if not os.path.exists(path):
            raise OSError("Path not exists")
        return path

    def set_user_watermark_path(self):
        try:
            self.custom_watermark_path_edit_line.setText(self.get_user_watermark_path())
        except OSError:
            self.custom_watermark_path_edit_line.setText(FILE_NOT_EXISTS_MESSAGE)

    def watermark_option(self):
        choosed_default_watermark = self.sender().text() == 'Стандартный'
        self.custom_watermark_path_edit_line.setEnabled(not choosed_default_watermark)
        if choosed_default_watermark:
            self.custom_watermark_path_edit_line.clear()
            self.set_user_watermark_path()
        else:
            self.custom_watermark_path_edit_line.clear()
            filename = QtWidgets.QFileDialog.getOpenFileName(
                self,
                "Выбрать файл",
                ".",
                "png(*.png);;"
                "jpg(*.jpg);;"
                "bmp(*.bmp);;"
            )[0]
            if filename:
                self.custom_watermark_path_edit_line.setText(filename)

    def select_auto_folder(self, auto_save_folder):
        if auto_save_folder:
            self.auto_choose_save_folder_radio_button.setChecked(True)
        else:
            self.manually_choose_save_folder_radio_button.setChecked(True)

    def clear_settings(self):
        self.filter_by_date_check_box.setChecked(False)

        for filter_line in (self.constructor_filter, self.checker_filter, self.gauger_filter):
            filter_line.select_filter_checkbox.setChecked(False)
            filter_line.data_combo_box.clear()
            filter_line.random_data_line_input.clear()

        self.add_water_mark_check_box.setChecked(False)
        self.split_files_by_size_checkbox.setChecked(False)

        self.exclude_folder_list_widget.clear()

    def set_default_settings(self):
        self.clear_settings()
        self.apply_user_settings()


class FilterSection(QtWidgets.QDialog):
    def __init__(
            self, check_box_label: str, combobox_info: str, widget_positions: FilterWidgetPositions,
            parent, parent_widget, font, font_1, font_2, size_policy, constructor_class
    ):
        super(FilterSection, self).__init__()
        self.check_box_label = check_box_label
        self.combobox_info = combobox_info
        self.construct_class = constructor_class
        self.parent = parent
        self.parent_widget = parent_widget
        self.font = font
        self.font_1 = font_1
        self.font_2 = font_2
        self.size_policy = size_policy
        self.widget_positions = widget_positions

        self.select_filter_checkbox = self.construct_class.make_checkbox(
            font=self.font,
            text=self.check_box_label,
            command=self.switch_filter_input,
            parent=self.parent_widget,
        )
        self.parent.addWidget(self.select_filter_checkbox, *self.widget_positions.check_box_position)

        self.data_from_combobox_radio_button = self.construct_class.make_radio_button(
            text=self.combobox_info,
            font=self.font_1,
            parent=self.parent_widget,
            command=self.choose_data_radio_option
        )
        self.parent.addWidget(
            self.data_from_combobox_radio_button,
            *self.widget_positions.combobox_radio_button_position
        )

        self.data_from_line_radio_button = self.construct_class.make_radio_button(
            text='Другая',
            font=self.font_1,
            parent=self.parent_widget,
            command=self.choose_data_radio_option
        )
        self.parent.addWidget(
            self.data_from_line_radio_button,
            *self.widget_positions.input_radio_button_position
        )

        self.filter_btn_group = QtWidgets.QButtonGroup()
        self.filter_btn_group.addButton(self.data_from_combobox_radio_button)
        self.filter_btn_group.addButton(self.data_from_line_radio_button)

        self.data_combo_box = self.construct_class.create_checkable_combobox(
            font=self.font_2,
            parent=self.parent_widget
        )
        self.parent.addWidget(self.data_combo_box, *self.widget_positions.combobox_position)

        self.random_data_line_input = self.construct_class.make_line_edit(
            parent=self.parent_widget,
            font=self.font_1
        )
        self.random_data_line_input.setSizePolicy(self.size_policy)
        self.parent.addWidget(self.random_data_line_input, *self.widget_positions.input_line_position)

    def switch_filter_input(self):
        self.data_from_combobox_radio_button.setEnabled(self.select_filter_checkbox.isChecked())
        self.data_from_line_radio_button.setEnabled(self.select_filter_checkbox.isChecked())

        self.data_combo_box.setEnabled(self.select_filter_checkbox.isChecked())
        self.random_data_line_input.setEnabled(False)

        if self.select_filter_checkbox.isChecked():
            self.filter_btn_group.setExclusive(True)
            self.data_from_combobox_radio_button.setChecked(self.select_filter_checkbox.isChecked())
        else:
            self.filter_btn_group.setExclusive(False)
            self.data_from_combobox_radio_button.setChecked(self.select_filter_checkbox.isChecked())
            self.data_from_line_radio_button.setChecked(self.select_filter_checkbox.isChecked())

    def choose_data_radio_option(self):
        chosen_combo_box = self.sender().text() == self.combobox_info
        self.data_combo_box.setEnabled(chosen_combo_box)
        self.random_data_line_input.setEnabled(not chosen_combo_box)


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

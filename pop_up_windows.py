from __future__ import annotations

import json
import os

from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QGridLayout, QWidget

from widgets_tools import WidgetBuilder, ExcludeFolderListWidget, WidgetStyles
from schemas import SettingsData, Filters, FilterWidgetPositions, UserSettings, FilePath, SaveType
from utils import date_today_by_int, FILE_NOT_EXISTS_MESSAGE


DEFAULT_WATERMARK_PATH = 'bdt_stamp.png'
DEFAULT_WATERMARK_LABEL = 'Стандартный'


class SettingsWindow(QtWidgets.QDialog):
    def __init__(self):
        QtWidgets.QDialog.__init__(self)

        self.watermark_position = []
        self._watermark_path = ''

        self.construct_class = WidgetBuilder()
        self.style_class = WidgetStyles()

        self._arial_12_font = self.style_class.arial_12_font
        self._arial_11_font = self.style_class.arial_11_font
        self._arial_ms_shell_12_font = self.style_class.ms_shell_12_font

        self._date_policy = self.style_class.date_policy
        self._filter_policy = self.style_class.filter_policy
        self._check_box_policy = self.style_class.size_policy_button_2

        self._setup_ui()
        self._apply_user_settings()

    def collect_settings_window_info(self) -> SettingsData:
        date_range = self._collect_date_range_info()
        constructor_list = self._constructor_filter.collect_filter_data()
        checker_list = self._checker_filter.collect_filter_data()
        sortament_list = self._gauge_filter.collect_filter_data()

        filters = None
        if any([date_range, constructor_list, checker_list, sortament_list]):
            filters = Filters(
                date_range=date_range,
                constructor_list=constructor_list,
                checker_list=checker_list,
                sortament_list=sortament_list,
            )
        return SettingsData(
            filters=filters,
            watermark_path=self._get_watermark_path(),
            watermark_position=self.watermark_position,
            split_file_by_size=self._split_files_by_size_checkbox.isChecked(),
            save_type=self._get_save_type(),
            except_folders_list=self._exclude_folder_list_widget.get_items_text_data(),
        )

    def _collect_date_range_info(self) -> list[int, int] | None:
        if not self._filter_by_date_check_box.isChecked():
            return None
        date_1 = self._first_date_input.dateTime().toSecsSinceEpoch()
        date_2 = self._last_date_input.dateTime().toSecsSinceEpoch()
        date_range = sorted([date_1, date_2])
        return date_range

    def _get_watermark_path(self) -> FilePath | None:
        if not self._add_water_mark_check_box.isChecked():
            return None
        path = self._custom_watermark_path_edit_line.text()
        if path == FILE_NOT_EXISTS_MESSAGE:
            return None
        return path

    def _get_save_type(self):
        if self._manually_choose_save_folder_radio_button.isChecked():
            return SaveType.MANUALLY_SAVE_FOLDER
        return SaveType.AUTO_SAVE_FOLDER

    def _setup_ui(self):
        self.setObjectName("Form")
        self.resize(675, 520)
        self.setWindowTitle("Настройки")

        self.grid_layout_widget = QWidget(self)
        self.grid_layout_widget.setGeometry(QtCore.QRect(10, 10, 660, 500))
        self.grid_layout_widget.setObjectName("gridLayoutWidget")

        self.grid_layout = QGridLayout(self.grid_layout_widget)
        self.grid_layout.setContentsMargins(0, 0, 0, 0)

        self._setup_filter_by_date_section()
        self._setup_filter_by_constructor_section()
        self._setup_filter_by_checker_section()
        self._setup_filter_by_gauge_section()
        self._setup_watermark_section()
        self._setup_split_by_size_section()
        self._setup_save_file_path_section()
        self._setup_exclude_folders_section()
        self._setup_window_settings()

    def _setup_filter_by_date_section(self):
        self._filter_by_date_check_box = self.construct_class.make_checkbox(
            font=self._arial_12_font,
            text='C датой только за указанный период',
            command=self._switch_date_input_filter,
            parent=self.grid_layout_widget
        )
        self._filter_by_date_check_box.setSizePolicy(self._check_box_policy)
        self.grid_layout.addWidget(self._filter_by_date_check_box, 1, 0, 1, 1)

        self._first_date_input = self.construct_class.make_date(
            font=self._arial_12_font,
            parent=self.grid_layout_widget
        )
        self.grid_layout.addWidget(self._first_date_input, 1, 2, 1, 1)

        self._last_date_input = self.construct_class.make_date(
            font=self._arial_12_font,
            parent=self.grid_layout_widget
        )
        self.grid_layout.addWidget(self._last_date_input, 1, 1, 1, 1)

    def _setup_filter_by_constructor_section(self):
        constructor_positions = FilterWidgetPositions(
            check_box_position=[2, 0, 3, 1],
            combobox_position=[3, 1, 1, 1],
            input_line_position=[3, 2, 1, 1],
            combobox_radio_button_position=[2, 1, 1, 1],
            input_radio_button_position=[2, 2, 1, 1]
        )
        self._constructor_filter = FilterSection(
            'С указанной фамилией разработчиков',
            'Фамилии из списка',
            constructor_positions,
            self.grid_layout,
            self.grid_layout_widget,
            self.style_class,
            self.construct_class
        )

    def _setup_filter_by_checker_section(self):
        checker_positions = FilterWidgetPositions(
            check_box_position=[5, 0, 2, 1],
            combobox_position=[6, 1, 1, 1],
            input_line_position=[6, 2, 1, 1],
            combobox_radio_button_position=[5, 1, 1, 1],
            input_radio_button_position=[5, 2, 1, 1]
        )
        self._checker_filter = FilterSection(
            'С указанной фамилией проверяющего',
            'Фамилии из списка',
            checker_positions,
            self.grid_layout,
            self.grid_layout_widget,
            self.style_class,
            self.construct_class
        )

    def _setup_filter_by_gauge_section(self):
        gauge_positions = FilterWidgetPositions(
            check_box_position=[7, 0, 2, 1],
            combobox_position=[8, 1, 1, 1],
            input_line_position=[8, 2, 1, 1],
            combobox_radio_button_position=[7, 1, 1, 1],
            input_radio_button_position=[7, 2, 1, 1]
        )
        self._gauge_filter = FilterSection(
            'С указанным сортаментом',
            'Сортамент из списка',
            gauge_positions,
            self.grid_layout,
            self.grid_layout_widget,
            self.style_class,
            self.construct_class
        )

    def _setup_watermark_section(self):
        self._add_water_mark_check_box = self.construct_class.make_checkbox(
            text='Добавить водяной знак',
            font=self._arial_12_font, parent=self.grid_layout_widget,
            command=self._switch_watermark_group
        )
        self.grid_layout.addWidget(self._add_water_mark_check_box, 9, 0, 2, 1)

        self._default_watermark_path_radio_button = self.construct_class.make_radio_button(
            text=DEFAULT_WATERMARK_LABEL,
            font=self._arial_12_font,
            parent=self.grid_layout_widget,
            command=self._watermark_path_radio_option
        )
        self.grid_layout.addWidget(self._default_watermark_path_radio_button, 9, 1, 1, 1)

        self._custom_watermark_path_radio_button = self.construct_class.make_radio_button(
            text='Свое изображение', font=self._arial_12_font,
            parent=self.grid_layout_widget,
            command=self._watermark_path_radio_option
        )
        self.grid_layout.addWidget(self._custom_watermark_path_radio_button, 9, 2, 1, 1)

        self._custom_watermark_path_edit_line = self.construct_class.make_line_edit(
            parent=self.grid_layout_widget,
            font=self._arial_11_font
        )
        self.grid_layout.addWidget(self._custom_watermark_path_edit_line, 10, 1, 1, 2)

        self._watermark_path_btn_group = QtWidgets.QButtonGroup()
        self._watermark_path_btn_group.addButton(self._default_watermark_path_radio_button)
        self._watermark_path_btn_group.addButton(self._custom_watermark_path_radio_button)

    def _setup_split_by_size_section(self):
        self._split_files_by_size_checkbox = self.construct_class.make_checkbox(
            text='Разбить на файлы по размерам',
            font=self._arial_12_font,
            parent=self.grid_layout_widget
        )
        self.grid_layout.addWidget(self._split_files_by_size_checkbox, 11, 0, 1, 1)

    def _setup_save_file_path_section(self):
        self._manually_choose_save_folder_radio_button = self.construct_class.make_radio_button(
            text='Указать папку сохранения вручную',
            font=self._arial_ms_shell_12_font,
            parent=self.grid_layout_widget
        )
        self.grid_layout.addWidget(self._manually_choose_save_folder_radio_button, 12, 0, 1, 1)

        self._auto_choose_save_folder_radio_button = self.construct_class.make_radio_button(
            text='Выбрать папку автоматически',
            font=self._arial_ms_shell_12_font,
            parent=self.grid_layout_widget
        )
        self.grid_layout.addWidget(self._auto_choose_save_folder_radio_button, 12, 1, 1, 3)

        save_folder_btngroup = QtWidgets.QButtonGroup()
        save_folder_btngroup.addButton(self._manually_choose_save_folder_radio_button)
        save_folder_btngroup.addButton(self._auto_choose_save_folder_radio_button)

    def _setup_exclude_folders_section(self):
        exclude_folder_label = self.construct_class.make_label(
            text='Исключить следующие папки:',
            font=self._arial_12_font,
            parent=self.grid_layout_widget
        )
        self.grid_layout.addWidget(exclude_folder_label, 13, 0, 1, 3)

        self._exclude_folder_list_widget = ExcludeFolderListWidget(self.grid_layout_widget)
        self._exclude_folder_list_widget.setFont(self._arial_12_font)
        self.grid_layout.addWidget(self._exclude_folder_list_widget, 14, 0, 2, 2)

        add_exclude_folder_button = self.construct_class.make_button(
            text='Добавить папку',
            parent=self.grid_layout_widget,
            font=self._arial_12_font,
            size_policy=self._date_policy,
            command=self._exclude_folder_list_widget.add_folder
        )
        self.grid_layout.addWidget(add_exclude_folder_button,  14, 2, 1, 1)

        delete_exclude_folder_button = self.construct_class.make_button(
            text='Удалить выбранную\n папку',
            parent=self.grid_layout_widget,
            font=self._arial_12_font,
            size_policy=self._date_policy,
            command=self._exclude_folder_list_widget.remove_item
        )
        self.grid_layout.addWidget(delete_exclude_folder_button, 15, 2, 1, 1)

    def _setup_window_settings(self):
        reset_settings_button = self.construct_class.make_button(
            text='Сбросить настройки',
            parent=self.grid_layout_widget,
            font=self._arial_12_font,
            command=self._set_default_settings
        )
        self.grid_layout.addWidget(reset_settings_button, 16, 1, 1, 3)

        self._close_window_button = self.construct_class.make_button(
            text='Ок',
            parent=self.grid_layout_widget,
            font=self._arial_12_font,
            command=self.close
        )
        self.grid_layout.addWidget(self._close_window_button, 16, 0, 1, 1)
        QtCore.QMetaObject.connectSlotsByName(self)

    def _apply_user_settings(self):
        loaded_user_settings = self._get_settings_from_file()
        self._watermark_path = loaded_user_settings.watermark_path
        self.watermark_position = loaded_user_settings.watermark_position
        if loaded_user_settings:
            self._fill_widgets_with_settings(loaded_user_settings)

    def _get_settings_from_file(self) -> UserSettings | None:
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
            self.construct_class.send_error('В Файл settings.txt \n присутсвуют ошибки \n синтаксиса json')
            return

        return user_settings

    def _fill_widgets_with_settings(self, user_settings: UserSettings):
        self.construct_class.set_date(date_today_by_int(), self._first_date_input)
        self.construct_class.set_date(date_today_by_int(), self._last_date_input)
        self._switch_date_input_filter()

        if user_settings.except_folders_list:
            self._exclude_folder_list_widget.fill_list(draw_list=user_settings.except_folders_list)

        self.construct_class.fill_combo_box(user_settings.constructor_list, self._constructor_filter.data_combo_box)
        self._constructor_filter.switch_filter_input()

        self.construct_class.fill_combo_box(user_settings.checker_list, self._checker_filter.data_combo_box)
        self._checker_filter.switch_filter_input()

        self.construct_class.fill_combo_box(user_settings.sortament_list, self._gauge_filter.data_combo_box)
        self._gauge_filter.switch_filter_input()

        self._add_water_mark_check_box.setChecked(user_settings.add_default_watermark)
        self._switch_watermark_group()
        self._set_user_watermark_path()

        self._split_files_by_size_checkbox.setChecked(user_settings.split_file_by_size)
        self._select_auto_folder(user_settings.auto_save_folder)

    def _switch_date_input_filter(self):
        self._first_date_input.setEnabled(self._filter_by_date_check_box.isChecked())
        self._last_date_input.setEnabled(self._filter_by_date_check_box.isChecked())

    def _switch_watermark_group(self):
        self._default_watermark_path_radio_button.setEnabled(self._add_water_mark_check_box.isChecked())
        self._custom_watermark_path_radio_button.setEnabled(self._add_water_mark_check_box.isChecked())
        self._custom_watermark_path_edit_line.setEnabled(False)

        if self._add_water_mark_check_box.isChecked():
            self._watermark_path_btn_group.setExclusive(True)
            self._default_watermark_path_radio_button.setChecked(self._add_water_mark_check_box.isChecked())
        else:
            self._watermark_path_btn_group.setExclusive(False)
            self._default_watermark_path_radio_button.setChecked(self._add_water_mark_check_box.isChecked())
            self._custom_watermark_path_radio_button.setChecked(self._add_water_mark_check_box.isChecked())

    def _get_user_watermark_path(self) -> str:
        path = os.path.abspath(DEFAULT_WATERMARK_PATH)
        if not os.path.exists(path):
            path = os.path.abspath(self._watermark_path)
        if not os.path.exists(path):
            raise OSError("Path not exists")
        return path

    def _set_user_watermark_path(self):
        try:
            self._custom_watermark_path_edit_line.setText(self._get_user_watermark_path())
        except OSError:
            self._custom_watermark_path_edit_line.setText(FILE_NOT_EXISTS_MESSAGE)

    def _watermark_path_radio_option(self):
        def request_user_file() -> str | None:
            user_filename = QtWidgets.QFileDialog.getOpenFileName(
                self,
                "Выбрать файл",
                ".",
                "png(*.png);;"
                "jpg(*.jpg);;"
                "bmp(*.bmp);;"
            )[0]
            return user_filename

        choosed_default_watermark = self.sender().text() == DEFAULT_WATERMARK_LABEL
        self._custom_watermark_path_edit_line.setEnabled(not choosed_default_watermark)
        self._custom_watermark_path_edit_line.clear()

        if choosed_default_watermark:
            self._set_user_watermark_path()
        else:
            if filename := request_user_file():
                self._custom_watermark_path_edit_line.setText(filename)

    def _select_auto_folder(self, auto_save_folder):
        if auto_save_folder:
            self._auto_choose_save_folder_radio_button.setChecked(True)
        else:
            self._manually_choose_save_folder_radio_button.setChecked(True)

    def _clear_settings(self):
        self._filter_by_date_check_box.setChecked(False)

        for filter_line in (self._constructor_filter, self._checker_filter, self._gauge_filter):
            filter_line.select_filter_checkbox.setChecked(False)
            filter_line.data_combo_box.clear()
            filter_line.random_data_line_input.clear()

        self._exclude_folder_list_widget.clear()

    def _set_default_settings(self):
        self._clear_settings()
        self._apply_user_settings()


class FilterSection(QtWidgets.QDialog):
    def __init__(
            self, check_box_label: str, combobox_info: str, widget_positions: FilterWidgetPositions,
            parent: QGridLayout, parent_widget: QWidget, style_class: WidgetStyles, constructor_class: WidgetBuilder
    ):
        super(FilterSection, self).__init__()
        self.check_box_label = check_box_label
        self.combobox_info = combobox_info
        self.construct_class = constructor_class
        self.parent = parent
        self.parent_widget = parent_widget

        self.arial_12_font = style_class.arial_12_font
        self.arial_11_font = style_class.arial_11_font
        self.arial_ms_shell_12_font = style_class.ms_shell_12_font

        self.size_policy = style_class.filter_policy
        self.widget_positions = widget_positions

        self.select_filter_checkbox = self.construct_class.make_checkbox(
            font=self.arial_12_font,
            text=self.check_box_label,
            command=self.switch_filter_input,
            parent=self.parent_widget,
        )
        self.parent.addWidget(self.select_filter_checkbox, *self.widget_positions.check_box_position)

        self.data_from_combobox_radio_button = self.construct_class.make_radio_button(
            text=self.combobox_info,
            font=self.arial_11_font,
            parent=self.parent_widget,
            command=self._choose_data_radio_option
        )
        self.parent.addWidget(
            self.data_from_combobox_radio_button,
            *self.widget_positions.combobox_radio_button_position
        )

        self.data_from_line_radio_button = self.construct_class.make_radio_button(
            text='Другая',
            font=self.arial_11_font,
            parent=self.parent_widget,
            command=self._choose_data_radio_option
        )
        self.parent.addWidget(
            self.data_from_line_radio_button,
            *self.widget_positions.input_radio_button_position
        )

        self.filter_btn_group = QtWidgets.QButtonGroup()
        self.filter_btn_group.addButton(self.data_from_combobox_radio_button)
        self.filter_btn_group.addButton(self.data_from_line_radio_button)

        self.data_combo_box = self.construct_class.create_checkable_combobox(
            font=self.arial_ms_shell_12_font,
            parent=self.parent_widget
        )
        self.parent.addWidget(self.data_combo_box, *self.widget_positions.combobox_position)

        self.random_data_line_input = self.construct_class.make_line_edit(
            parent=self.parent_widget,
            font=self.arial_11_font
        )
        self.random_data_line_input.setSizePolicy(self.size_policy)
        self.parent.addWidget(self.random_data_line_input, *self.widget_positions.input_line_position)

    def collect_filter_data(self) -> list[str] | None:
        if not self.select_filter_checkbox.isChecked():
            return None

        if self.data_from_combobox_radio_button.isChecked():
            constructors_list = self.data_combo_box.collect_checked_items()
        else:
            constructors_list = [self.random_data_line_input.text()]
        return constructors_list

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

    def _choose_data_radio_option(self):
        chosen_combo_box = self.sender().text() == self.combobox_info
        self.data_combo_box.setEnabled(chosen_combo_box)
        self.random_data_line_input.setEnabled(not chosen_combo_box)


class RadioButtonsWindow(QtWidgets.QDialog):

    def __init__(self, executions: list[str]):
        QtWidgets.QDialog.__init__(self)
        self.construct_class = WidgetBuilder()
        style_class = WidgetStyles()

        self.executions = executions
        self.layout = QtWidgets.QGridLayout()
        self.setLayout(self.layout)
        self.setWindowTitle("Вы указали групповую спецификацию")

        self._arial_12_font = style_class.arial_12_font

        self._set_label()
        self._set_options()
        self._add_buttons()

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

    def _set_label(self):
        plain_text = QtWidgets.QLabel('Выберите исполнение для слияния:')
        plain_text.setFont(self._arial_12_font)
        self.layout.addWidget(plain_text, 0, 0, 1, 3)

    def _set_options(self):
        for index, option in enumerate(self.executions):
            if type(option) != str:
                continue
            radiobutton = QtWidgets.QRadioButton(option)
            radiobutton.option = option
            radiobutton.toggled.connect(self.on_radio_clicked)
            radiobutton.setFont(self._arial_12_font)
            self.layout.addWidget(radiobutton, 2, index)

    def _add_buttons(self):
        button_layout = QtWidgets.QHBoxLayout()

        ok_button = QtWidgets.QPushButton('OК')
        ok_button.setFont(self._arial_12_font)
        ok_button.clicked.connect(self.on_button_clicked)
        button_layout.addWidget(ok_button)

        cancel_button = QtWidgets.QPushButton('Отмена')
        cancel_button.setFont(self._arial_12_font)
        cancel_button.clicked.connect(self.on_button_clicked)
        button_layout.addWidget(cancel_button)

        self.layout.addLayout(button_layout, 3, 0, 1, 4)

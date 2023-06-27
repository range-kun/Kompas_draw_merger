import webbrowser
from pathlib import Path
from typing import Callable

from PyQt5 import QtCore
from PyQt5 import QtGui
from PyQt5 import QtWidgets

from programm.pyqt_windows.pop_up_windows import SettingsWindow
from programm.widgets_tools import MainListWidget
from programm.widgets_tools import WidgetBuilder
from programm.widgets_tools import WidgetStyles


class MainWindow(WidgetBuilder):
    IMAGE_PATH = Path(__file__).parent.parent.parent.resolve() / "img"

    def __init__(self):
        WidgetBuilder.__init__(self, parent=None)
        self.setFixedSize(930, 646)
        self.setWindowTitle("Конвертер")

        self.settings_window = SettingsWindow()
        self.settings_window_data = self.settings_window.collect_settings_window_info()
        self.style_class = WidgetStyles()
        self.setup_ui()

    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_Delete:
            self.change_list_widget_state(self.list_widget.remove_selected)

    def setup_ui(self):
        self.setup_styles()
        self.setFont(self.style_class.arial_12_font)
        self.grid_widget.setGeometry(QtCore.QRect(10, 10, 906, 611))

        self.setup_look_up_section()
        self.setup_spec_section()
        self.setup_look_up_parameters_section()
        self.setup_upper_list_buttons()
        self.setup_list_widget_section()
        self.setup_lower_list_buttons()
        self.setup_bottom_section()

    def setup_styles(self):
        self.central_widget = QtWidgets.QWidget(self)
        self.grid_widget = QtWidgets.QWidget(self.central_widget)
        self.setCentralWidget(self.central_widget)

        self.grid_layout = QtWidgets.QGridLayout(self.grid_widget)
        self.grid_layout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.grid_layout.setContentsMargins(0, 0, 0, 0)

    def setup_look_up_section(self):
        self.source_of_draws_field = self.make_text_edit(
            font=self.style_class.arial_11_font,
            placeholder="Выберите папку с файлами в формате .spw или .cdw",
            size_policy=self.style_class.line_edit_size_policy,
        )
        self.grid_layout.addWidget(self.source_of_draws_field, 1, 0, 1, 2)

        self.choose_folder_button = self.make_button(
            text="Выбор папки \n с чертежами для поиска",
            font=self.style_class.arial_12_font,
        )
        self.grid_layout.addWidget(self.choose_folder_button, 1, 2, 1, 1)
        self.choose_folder_button.setToolTip(
            "Добавить все чертежи в формате .cdw и .spw указанной папке в список ниже"
        )

        self.choose_data_base_button = self.make_button(
            text="Выбор файла\n с базой чертежей",
            font=self.style_class.arial_12_font,
            enabled=False,
        )
        self.choose_data_base_button.setToolTip(
            "Выберите имеющейся файл в формате .json, в котором будут содержаться пути до чертежей"
        )
        self.grid_layout.addWidget(self.choose_data_base_button, 1, 3, 1, 1)

    def setup_spec_section(self):
        self.choose_specification_button = self.make_button(
            text="Выбор \nспецификации",
            font=self.style_class.arial_12_font,
            enabled=False,
        )
        self.grid_layout.addWidget(self.choose_specification_button, 2, 2, 1, 1)

        self.save_data_base_file_button = self.make_button(
            text="Сохранить \n базу чертежей",
            font=self.style_class.arial_12_font,
            enabled=False,
        )
        self.grid_layout.addWidget(self.save_data_base_file_button, 2, 3, 1, 1)

        self.path_to_spec_field = self.make_text_edit(
            font=self.style_class.arial_11_font,
            placeholder="Укажите путь до файла со спецификацией .spw",
            size_policy=self.style_class.line_edit_size_policy,
        )
        self.path_to_spec_field.setEnabled(False)
        self.grid_layout.addWidget(self.path_to_spec_field, 2, 0, 1, 2)

    def setup_look_up_parameters_section(self):
        self.search_in_folder_radio_button = self.make_radio_button(
            text="Поиск по папке",
            font=self.style_class.arial_12_font,
        )
        self.search_in_folder_radio_button.setChecked(True)
        self.grid_layout.addWidget(self.search_in_folder_radio_button, 3, 2, 1, 1)

        self.search_by_spec_radio_button = self.make_radio_button(
            text="Поиск по спецификации",
            font=self.style_class.arial_12_font,
        )
        self.grid_layout.addWidget(self.search_by_spec_radio_button, 3, 0, 1, 1)

        self.bypassing_folders_inside_checkbox = self.make_checkbox(
            font=self.style_class.arial_12_font,
            text="С обходом всех папок внутри",
            activate=True,
        )
        self.grid_layout.addWidget(self.bypassing_folders_inside_checkbox, 3, 3, 1, 1)

        self.bypassing_sub_assemblies_chekcbox = self.make_checkbox(
            font=self.style_class.arial_12_font,
            text="С поиском по подсборкам",
        )
        self.bypassing_sub_assemblies_chekcbox.setEnabled(False)
        self.grid_layout.addWidget(self.bypassing_sub_assemblies_chekcbox, 3, 1, 1, 1)

    def setup_upper_list_buttons(self):
        upper_items_list_layout = QtWidgets.QHBoxLayout()
        self.grid_layout.addLayout(upper_items_list_layout, 4, 0, 1, 4)

        self.clear_draw_list_button = self.make_button(
            text="Очистить список и выбор папки для поиска",
            font=self.style_class.arial_12_font,
        )
        upper_items_list_layout.addWidget(self.clear_draw_list_button)

        self.refresh_draw_list_button = self.make_button(
            text="Обновить файлы для склеивания",
            font=self.style_class.arial_12_font,
        )
        self.refresh_draw_list_button.setToolTip(
            "Обновит список чертежей снизу, применив "
            "фильтры поиска, обновить папку с чертежами и.т.д"
        )
        upper_items_list_layout.addWidget(self.refresh_draw_list_button)

        self.save_items_list = self.make_button(
            text="Скопировать выбранные файлы",
            font=self.style_class.arial_12_font,
        )
        upper_items_list_layout.addWidget(self.save_items_list)
        self.save_items_list.setToolTip(
            "Скопирует все выбранные файлы из списка ниже в указанную папку."
        )

    def setup_list_widget_section(self):
        horizontal_layout = QtWidgets.QHBoxLayout()
        self.grid_layout.addLayout(horizontal_layout, 8, 0, 1, 4)

        self.list_widget = MainListWidget(self.grid_widget)
        self.list_widget.setToolTip(
            "Для выбора нескольких чертежей используйте ctrl или shift, для удаления del"
        )
        horizontal_layout.addWidget(self.list_widget)

        vertical_layout = QtWidgets.QVBoxLayout()
        horizontal_layout.addLayout(vertical_layout)

        self.move_line_up_button = self.make_button(
            text="\n\n",
            size_policy=self.style_class.size_policy_button_2,
            command=self.list_widget.move_item_up,
        )
        self.move_line_up_button.setIcon(QtGui.QIcon(str(self.IMAGE_PATH / "arrow_up.png")))
        self.move_line_up_button.setIconSize(QtCore.QSize(50, 50))
        self.move_line_up_button.setToolTip("Переместить выбранный чертеж вверх")

        self.move_line_down_button = self.make_button(
            text="\n\n",
            size_policy=self.style_class.size_policy_button_2,
            command=self.list_widget.move_item_down,
        )
        self.move_line_down_button.setIcon(QtGui.QIcon(str(self.IMAGE_PATH / "arrow_down.png")))
        self.move_line_down_button.setIconSize(QtCore.QSize(50, 50))
        self.move_line_down_button.setToolTip("Переместить выбранный чертеж вниз")

        self.delete_list_widget_item = self.make_button(
            text="\n\n",
            size_policy=self.style_class.size_policy_button_2,
            command=lambda: self.change_list_widget_state(self.list_widget.remove_selected),
        )
        self.delete_list_widget_item.setIcon(QtGui.QIcon(str(self.IMAGE_PATH / "red_cross.png")))
        self.delete_list_widget_item.setIconSize(QtCore.QSize(50, 50))
        self.delete_list_widget_item.setToolTip("Удалить выбранный чертеж/чертежи из списка")

        self.help_button = self.make_button(
            text="\n\n",
            size_policy=self.style_class.size_policy_button_2,
            command=lambda: webbrowser.open('https://youtu.be/L_o0YrXBaFo'),
        )
        self.help_button.setIcon(QtGui.QIcon(str(self.IMAGE_PATH / "info.png")))
        self.help_button.setIconSize(QtCore.QSize(50, 50))
        self.help_button.setToolTip("Ссылка на справочное видео")

        vertical_layout.addWidget(self.move_line_up_button)
        vertical_layout.addWidget(self.move_line_down_button)
        vertical_layout.addWidget(self.delete_list_widget_item)
        vertical_layout.addWidget(self.help_button)

    def setup_lower_list_buttons(self):
        self.select_all_button = self.make_button(
            text="Выделить все",
            font=self.style_class.arial_12_font,
            command=self.list_widget.select_all,
            size_policy=self.style_class.size_policy_button,
        )
        self.grid_layout.addWidget(self.select_all_button, 10, 0, 1, 1)

        self.remove_selection_button = self.make_button(
            text="Снять выделение",
            font=self.style_class.arial_12_font,
            command=self.list_widget.unselect_all,
            size_policy=self.style_class.size_policy_button,
        )
        self.grid_layout.addWidget(self.remove_selection_button, 10, 1, 1, 1)

        self.add_file_to_list_button = self.make_button(
            text="Добавить файл в список",
            font=self.style_class.arial_12_font,
        )
        self.grid_layout.addWidget(self.add_file_to_list_button, 10, 2, 1, 1)

        self.add_folder_to_list_button = self.make_button(
            text="Добавить папку в список",
            font=self.style_class.arial_12_font,
        )
        self.add_folder_to_list_button.setToolTip(
            "Добавить чертежи в формате .cdw и .spw в указанной папке в "
            "конец списка (файлы из вложенных папок не добавляются)"
        )
        self.grid_layout.addWidget(self.add_folder_to_list_button, 10, 3, 1, 1)
        self.switch_select_unselect_buttons(False)

    def setup_bottom_section(self):
        horizontal_layout = QtWidgets.QHBoxLayout()
        self.grid_layout.addLayout(horizontal_layout, 11, 0, 1, 4)

        self.additional_settings_button = self.make_button(
            text="Дополнительные настройки",
            font=self.style_class.arial_12_font,
            command=self.show_settings,
        )
        horizontal_layout.addWidget(self.additional_settings_button)

        self.delete_single_draws_after_merge_checkbox = self.make_checkbox(
            font=self.style_class.arial_12_font,
            text="Удалить отдельные PDF-чертежи по окончанию",
            activate=True,
        )
        self.delete_single_draws_after_merge_checkbox.setToolTip(
            "В процессе сливания чертежей программа "
            "конвертирует их в отдельные PDF файлы и хранит на диске"
        )
        horizontal_layout.addWidget(self.delete_single_draws_after_merge_checkbox)

        self.open_file_after_merge_chekbox = self.make_checkbox(
            text="Открывать папку и файл по окончанию",
            font=self.style_class.arial_12_font,
            activate=True,
        )
        horizontal_layout.addWidget(self.open_file_after_merge_chekbox)

        self.merge_files_button = self.make_button(
            text="Склеить файлы",
            font=self.style_class.arial_12_font_bold,
        )
        self.grid_layout.addWidget(self.merge_files_button, 13, 0, 1, 4)

        self.progress_bar = QtWidgets.QProgressBar(self.grid_widget)
        self.progress_bar.setTextVisible(False)
        self.grid_layout.addWidget(self.progress_bar, 15, 0, 1, 4)

        self.status_bar = QtWidgets.QStatusBar(self)
        self.status_bar.setObjectName("statusbar")
        self.setStatusBar(self.status_bar)
        QtCore.QMetaObject.connectSlotsByName(self)

    def change_list_widget_state(self, method: Callable, *args, **kwargs):
        method(*args, **kwargs)
        self.switch_select_unselect_buttons(self.list_widget.count() > 0)

    def switch_select_unselect_buttons(self, status: bool):
        self.select_all_button.setEnabled(status)
        self.remove_selection_button.setEnabled(status)
        self.save_items_list.setEnabled(status)
        self.delete_list_widget_item.setEnabled(status)
        self.move_line_up_button.setEnabled(status)
        self.move_line_down_button.setEnabled(status)

    def increase_step(self, current_progres):
        self.progress_bar.setValue(int(current_progres))

    def switch_button_group(self, switch=None):
        if not switch:
            switch = False if self.merge_files_button.isEnabled() else True

        self.merge_files_button.setEnabled(switch)
        self.choose_folder_button.setEnabled(switch)
        self.additional_settings_button.setEnabled(switch)
        self.refresh_draw_list_button.setEnabled(switch)
        self.choose_data_base_button.setEnabled(switch)
        self.choose_specification_button.setEnabled(switch)

    def show_settings(self):
        self.settings_window.exec_()
        self.settings_window_data = self.settings_window.collect_settings_window_info()

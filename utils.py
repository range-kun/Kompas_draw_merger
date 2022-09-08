import itertools
import os
import re
import time
from datetime import datetime
from operator import itemgetter
from pathlib import Path

from kompas_api import fetch_obozn_and_execution
from schemas import FilePath, FileName, DrawObozn, DrawExecution, ErrorType
from PyQt5 import QtWidgets

FILE_NOT_EXISTS_MESSAGE = "Путь к файлу из настроек не существует"


class FileNotSpec(Exception):
    pass


def date_today_by_int() -> list[int]:
    return [int(i) for i in str(datetime.date(datetime.now())).split('-')]


def get_today_date():
    return time.strftime("%d.%m.%Y")


def check_specification(spec_path: FilePath):
    if not spec_path.endswith('.spw') or not os.path.isfile(spec_path):
        raise FileNotSpec('Указанный файл не является спецификацией или не существует')
    if not os.path.isfile(spec_path):
        raise FileExistsError("Указанный файл не существует")


def date_to_seconds(date_string):
    if len(date_string.split('.')[-1]) == 2:
        new_date = date_string.split('.')
        new_date[-1] = '20' + new_date[-1]
        date_string = '.'.join(new_date)
    struct_date = time.strptime(date_string, "%d.%m.%Y")
    return time.mktime(struct_date)


def sort_files_by_name(directory_with_draws):
    return sorted(os.listdir(directory_with_draws), key=lambda fil: int(fil.split()[0]))


class MergerFolderData:
    def __init__(
            self,
            path_to_search: FilePath,
            need_to_be_split: bool,
            draw_file_paths: list[FilePath],
            merger_class,
            directory_to_save: str | None = None
    ):
        self._directory_to_save = directory_to_save
        self._path_to_search = path_to_search
        self._draw_file_paths = draw_file_paths
        self._need_to_be_split = need_to_be_split
        self._ui_merger = merger_class
        self._single_draw_dir_name = 'Однодетальные'

        self._today_date = get_today_date()
        self._core_dir = rf'{directory_to_save or self._path_to_search}\pdf'

        self._main_draw_name = self._fetch_main_draw_name()
        self.single_draw_dir = self._fetch_single_draw_dir()
        self.pdf_file_paths = self._create_single_detail_pdf_paths()

    def create_main_pdf_file_path(self, size: list[float] = None) -> FilePath:
        if size:  # need to split
            format_name = self._fetch_format(size)
            return FilePath(os.path.join(
                os.path.dirname(self.single_draw_dir), f'{format_name}-{self._main_draw_name}.pdf'
            ))
        else:
            pdf_file_name = f"{os.path.basename(self.single_draw_dir[:-len(self._single_draw_dir_name)])}.pdf"
            return FilePath(os.path.join(os.path.dirname(self.single_draw_dir), pdf_file_name))

    @staticmethod
    def _fetch_format(value: list[float]) -> str:

        format_sizes = {(595, 842): "A4", (842, 1190): "A3", (1190, 1684): "A2", (1684, 2384): "A1", (2384, 3370): "A0"}
        min_file_size, max_file_size = sorted(map(int, value))

        format_name = format_sizes.get((min_file_size, max_file_size))
        if format_name:
            return format_name

        for min_format_size, max_format_size in format_sizes.keys():
            if min_format_size - 2 <= min_file_size <= min_format_size + 2 and \
                    max_format_size - 2 <= max_file_size <= max_format_size + 2:
                return format_name

        return f"{min_file_size}x{max_file_size}"

    def _create_single_detail_pdf_paths(self) -> list[FilePath]:
        single_draw_file_template = self.single_draw_dir + r"\\{0} {1}.pdf"
        file_paths = [
            FilePath(single_draw_file_template.format(number + 1, os.path.basename(file)))
            for number, file in enumerate(self._draw_file_paths)
        ]
        return file_paths

    def _fetch_main_draw_name(self) -> FileName:
        if self._ui_merger.search_by_spec_radio_button.isChecked():
            main_name = Path(self._ui_merger.specification_path).stem
        else:
            main_name = os.path.basename(self._path_to_search)
        return FileName(main_name)

    def _get_current_file_prefix_number(self) -> str:
        # next code check if folder or file with same name exists if so:
        # get maximum number of files and folders and incriminate +1 to name of new file and folder
        try:
            string_of_files = ' '.join(os.listdir(self._core_dir))
        except FileNotFoundError:
            return '01'

        cur_max_prefix_number = \
            max(map(int, re.findall(rf'{self._main_draw_name} - (\d\d)(?= {self._today_date})', string_of_files)),
                default=0)
        return str(cur_max_prefix_number + 1) if cur_max_prefix_number > 8 else '0' + str(cur_max_prefix_number + 1)

    def _fetch_single_draw_dir(self) -> FilePath:
        prefix_number = self._get_current_file_prefix_number()
        _single_draw_dir = fr'{self._core_dir}\{self._main_draw_name} - {prefix_number} {self._today_date}'
        if self._need_to_be_split:
            _single_draw_dir += rf'\{self._single_draw_dir_name}'
        else:
            _single_draw_dir += f' {self._single_draw_dir_name}'
        return FilePath(_single_draw_dir)


class DrawOboznCreation:
    def __init__(self, draw_obozn: DrawObozn):
        self.draw_obozn = draw_obozn
        self.need_to_verify_path = False
        self.execution: DrawExecution = DrawExecution("")

    @property
    def draw_obozn_list(self) -> list[DrawObozn]:
        draw_info = fetch_obozn_and_execution(self.draw_obozn)
        if not draw_info:
            return self._create_duplicate_obozn_list()
        self.need_to_verify_path = True
        _, self.execution, _ = draw_info
        return self._create_obozn_with_different_executions(*draw_info)

    def _create_duplicate_obozn_list(self) -> list[DrawObozn]:
        """
            Иногда в штампе указываектся обозначение для двух исполнений одной сборки xxx.00.00.00 и xxx.00.00.00-01
            при формирование базы данных, ключ для такого чертежа будет записан следующим образом:
            xxx.00.00.00xxx.00.00.01-01 данный цикл проверяет имеется ли такой ключ в базе данных
        """
        spec_obozn_list = []
        _modification_symbol = ""
        if not self.draw_obozn[-1].isdigit():
            # if draw have been modified xxx.00.00.01 -> xxx.00.00.01A it's gonna be A modification symbol
            _modification_symbol = self.draw_obozn[-1]

        db_obozn = self.draw_obozn
        for num in range(1, 4):  # обычно максимальное количество исполнений до -03
            db_obozn += DrawObozn(self.draw_obozn + f"-0{num}{_modification_symbol}")
            spec_obozn_list.append(db_obozn)
        return spec_obozn_list

    @staticmethod
    def _create_obozn_with_different_executions(
            spec_obozn: DrawObozn,
            execution: DrawExecution,
            modification_symbol: str | None
    ) -> list[DrawObozn]:
        """
            При формировании спецификации в штампе указывается только одна сборка
            а различные исполнения указываются
            только на чертеже при этом состав и количество деталей не изменно.
            Или может быть групповая спецефикация с одной спекой
        """
        if modification_symbol is None:
            modification_symbol = ""
        obozn_list = [
            DrawObozn(f"{spec_obozn}-0{execution_number}" + modification_symbol)
            for execution_number in range(int(execution), 0, -1)
        ]
        obozn_list.append(DrawObozn(f"{spec_obozn}{modification_symbol}"))
        return obozn_list


class ErrorsPrinter:
    def __init__(self, missing_list: list, error_type: ErrorType):
        self._missing_list = missing_list
        self._error_type = error_type
        self._message_for_file: str | None = None
        self._title_for_file: str | None = None

    def create_error_message(self) -> tuple[str, str]:
        if len(self._missing_list) > 15:
            return "Печать ошибок",\
                   "Для выводва на экран количество ошибок слишько велико, хотите сохрнаить отчет в текстовый файл?"
        return self.proceed_errors_list()

    @property
    def message_for_file(self) -> str:
        if self._message_for_file is None:
            _, self._message_for_file = self.proceed_errors_list()
        return self._message_for_file

    def proceed_errors_list(self) -> tuple[str, str]:
        one_line_messages, grouped_messages = self._sort_messages_by_type()
        missing_message = self._group_missing_files_info(grouped_messages)
        if self._error_type == ErrorType.FILE_MISSING:
            title, message = self._create_missing_files_message(missing_message, one_line_messages)
        elif self._error_type == ErrorType.FILE_NOT_OPENED:
            title, message = self._create_file_erorrs_message(missing_message, one_line_messages)
        else:
            raise NotImplementedError
        self._title_for_file, self._message_for_file = title, message

        return title, message

    def _sort_messages_by_type(self) -> tuple[list[str], list[tuple[FileName, DrawObozn]]]:
        one_line_messages = []
        grouped_messages: list[tuple[FileName, DrawObozn]] = []
        for error in self._missing_list:
            if type(error) == str:
                one_line_messages.append(error)
            else:
                grouped_messages.append(error)

        return one_line_messages, grouped_messages

    @staticmethod
    def _group_missing_files_info(grouped_messages: list[tuple[FileName, DrawObozn]]) -> str:
        grouped_list = itertools.groupby(grouped_messages, itemgetter(0))
        grouped_list_message = [key + ':\n' + '\n'.join(['----' + v for k, v in value]) for key, value in grouped_list]
        return '\n'.join(grouped_list_message)

    @staticmethod
    def _create_missing_files_message(missing_message: str, one_line_messages: list[str]) -> tuple[str, str]:
        window_title = "Отсуствующие чертежи"
        error_message = f"Не были найдены следующи чертежи:\n{missing_message} {''.join(one_line_messages)}" \
                        f"\nСохранить список?"
        return window_title, error_message

    @staticmethod
    def _create_file_erorrs_message(missing_message: str, one_line_messages: list[str]) -> tuple[str, str]:
        window_title = "Ошибки при обработке файлов"
        error_message = f"Были получены следующие ошибки при " \
                        f"обработке файлов:\n{missing_message} {''.join(one_line_messages)}" \
                        f"\nСохранить список?"
        return window_title, error_message

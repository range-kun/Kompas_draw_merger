from __future__ import annotations

import os
import queue
import shutil
import time
from collections import defaultdict
from typing import BinaryIO

import fitz
import pythoncom
from PyPDF2 import PdfFileMerger
from PyPDF2 import PdfMerger
from PyPDF2 import PdfReader
from PyPDF2 import PdfWriter
from PyQt5 import QtWidgets
from PyQt5.QtCore import pyqtSignal
from PyQt5.QtCore import QThread

from programm import schemas
from programm import utils
from programm.errors import FolderNotSelectedError
from programm.kompas_api import Converter
from programm.schemas import FILE_NOT_CHOSEN_MESSAGE
from programm.schemas import FilePath
from programm.schemas import SaveType
from programm.schemas import ThreadKompasAPI
from programm.utils import FILE_NOT_EXISTS_MESSAGE


class MergeThread(QThread):
    buttons_enable = pyqtSignal(bool)
    send_errors = pyqtSignal(str)
    status = pyqtSignal(str)
    kill_thread = pyqtSignal()
    increase_step = pyqtSignal(bool)
    progress_bar = pyqtSignal(int)
    choose_folder = pyqtSignal(bool)

    def __init__(
        self,
        *,
        files: list[FilePath],
        directory: FilePath,
        data_queue: queue.Queue,
        kompas_thread_api: ThreadKompasAPI,
        settings_window_data: schemas.SettingsData,
        merger_data: schemas.MergerData,
    ):
        self._files_path = files
        self._search_path = directory
        self.data_queue = data_queue
        self._kompas_thread_api = kompas_thread_api

        self._constructor_class = QtWidgets.QFileDialog()
        self._settings_window_data = settings_window_data
        self._merger_data = merger_data
        self._need_to_split_file = self._settings_window_data.split_file_by_size
        self._need_to_close_files: list[BinaryIO] = []

        QThread.__init__(self)

    def run(self):
        try:
            directory_to_save = self.select_save_folder()
        except FolderNotSelectedError:
            self._kill_thread()
            return

        self._file_paths_creator = self._initiate_file_paths_creator(directory_to_save)

        single_draw_dir = self._file_paths_creator.single_draw_dir
        os.makedirs(single_draw_dir)

        single_pdf_file_paths = self._file_paths_creator.pdf_file_paths
        self._convert_single_files_to_pdf(single_pdf_file_paths)

        merge_data = self._create_merger_data(single_pdf_file_paths)
        self._merge_pdf_files(merge_data)
        self._close_file_objects()

        if self._settings_window_data.watermark_path:
            self._add_watermark(list(merge_data.keys()))

        if self._merger_data.delete_single_draws_after_merge_checkbox:
            shutil.rmtree(single_draw_dir)

        if not self._settings_window_data.split_file_by_size:
            pdf_file = self._file_paths_creator.create_main_pdf_file_path()
            os.startfile(pdf_file)

        os.system(
            f"explorer {(os.path.normpath(os.path.dirname(single_draw_dir))).replace('//', '/')}"
        )

        self.buttons_enable.emit(True)
        self.progress_bar.emit(int(0))
        self.status.emit("Слитие успешно завершено")

    def _kill_thread(self):
        self.send_errors.emit("Запись прервана, папка не была найдена")
        self.buttons_enable.emit(True)
        self.progress_bar.emit(0)
        self.kill_thread.emit()

    def select_save_folder(self) -> FilePath | None:
        # If request folder from this thread later when trying to retrieve kompas api
        # Exception will be raised, that's why, folder is requested from main UiMerger class
        if self._settings_window_data.save_type == SaveType.AUTO_SAVE_FOLDER:
            return None
        self.choose_folder.emit(True)
        while True:
            time.sleep(0.1)
            try:
                directory_to_save = self.data_queue.get(block=False)
            except queue.Empty:
                pass
            else:
                break
        if directory_to_save == FILE_NOT_CHOSEN_MESSAGE:
            raise FolderNotSelectedError
        return directory_to_save

    def _initiate_file_paths_creator(
        self, directory_to_save: FilePath | None = None
    ) -> utils.MergerFolderData:
        return utils.MergerFolderData(
            self._search_path,
            self._need_to_split_file,
            self._files_path,
            self._merger_data,
            directory_to_save,
        )

    def _convert_single_files_to_pdf(self, pdf_file_paths: list[FilePath]):
        pythoncom.CoInitialize()
        _converter = Converter(self._kompas_thread_api)
        for file_path, pdf_file_path in zip(self._files_path, pdf_file_paths):
            self.increase_step.emit(True)
            self.status.emit(f"Конвертация {file_path}")
            _converter.convert_draw_to_pdf(file_path, pdf_file_path)

    @staticmethod
    def _merge_pdf_files(merge_data: dict[FilePath, PdfWriter | PdfFileMerger]):
        for pdf_file_path, merger_instance in merge_data.items():
            with open(pdf_file_path, "wb") as pdf:
                merger_instance.write(pdf)

    def _create_merger_data(
        self, pdf_file_paths: list[FilePath]
    ) -> dict[FilePath, PdfWriter | PdfFileMerger]:
        if self._need_to_split_file:
            merger_instance = self._create_split_merger_data(pdf_file_paths)
        else:
            merger_instance = self._create_single_merger_data(pdf_file_paths)

        return merger_instance

    def _create_split_merger_data(self, pdf_file_paths) -> dict[FilePath, PdfWriter]:
        # using different classes because PyPDF2.PdfWriter can add single page unlike PdfFileMerger
        merger_instance: defaultdict = defaultdict(PdfWriter)

        for pdf_file_path in pdf_file_paths:
            file = self._get_file_obj(pdf_file_path)
            merger_pdf_reader = PdfReader(file)

            for page in merger_pdf_reader.pages:
                size = page.mediabox[2:]
                file_path = self._file_paths_creator.create_main_pdf_file_path(size)
                merger_instance[file_path].add_page(page)

            self.increase_step.emit(True)
            self.status.emit(f"Сливание {pdf_file_path}")

        return merger_instance

    def _create_single_merger_data(
        self, pdf_file_paths: list[FilePath]
    ) -> dict[FilePath, PdfMerger]:
        # using different classes because PyPDF2.PdfWriter can add single page unlike PdfFileMerger
        merger_instance = PdfMerger()

        for single_pdf_file_path in pdf_file_paths:
            file = self._get_file_obj(single_pdf_file_path)
            merger_instance.append(fileobj=file)

            self.increase_step.emit(True)
            self.status.emit(f"Сливание {single_pdf_file_path}")

        file_path = self._file_paths_creator.create_main_pdf_file_path()
        return {file_path: merger_instance}

    def _get_file_obj(self, file_path: FilePath) -> BinaryIO:
        file = open(
            file_path, "rb"
        )  # files will be closed later, using with would lead to blank pdf lists
        self._need_to_close_files.append(file)
        return file

    def _close_file_objects(self):
        for file_obj in self._need_to_close_files:
            file_obj.close()
        self._need_to_close_files = []

    def _add_watermark(self, pdf_file_paths: list[FilePath]):
        def add_watermark_to_file(_pdf_file_path: FilePath):
            pdf_doc = fitz.open(_pdf_file_path)  # open the PDF
            rect = fitz.Rect(watermark_position)  # where to put image: use upper left corner
            for page in pdf_doc:
                if not page.is_wrapped:
                    page.wrap_contents()
                try:
                    page.insert_image(rect, filename=watermark_path, overlay=False)
                except ValueError:
                    self.send_errors.emit(
                        (
                            "Заданы неверные координаты, "
                            "размещения картинки, водяной знак не был добавлен"
                        )
                    )
                    return
            pdf_doc.saveIncr()  # do an incremental save

        watermark_path = self._settings_window_data.watermark_path
        watermark_position = self._settings_window_data.watermark_position

        if not watermark_position:
            return
        if not os.path.exists(watermark_path) or watermark_path == FILE_NOT_EXISTS_MESSAGE:
            self.send_errors.emit("Путь к файлу с картинкой не существует")
            return

        for pdf_file_path in pdf_file_paths:
            add_watermark_to_file(pdf_file_path)

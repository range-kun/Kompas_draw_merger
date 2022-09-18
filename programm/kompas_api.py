# -*- coding: utf-8 -*-
from __future__ import annotations

import enum
import os
import re
from contextlib import contextmanager
from typing import TypeVar, Type

import pythoncom
from win32com.client import Dispatch, gencache


from programm.schemas import DrawType, SpecSectionData, DrawData, FilePath, DrawObozn, DrawExecution, DrawName, \
    ThreadKompasAPI

DIFFERENCE_ON_DRAW_MESSAGES = ["различияисполненийпосборочномучертежу", "отличияисполненийпосборочномучертежу"]
WITHOUT_EXECUTION = 1000

T = TypeVar("T")


class ObjectType(enum.IntEnum):
    OBOZN_ISP = 4
    REGULAR_LINE = 1


class SpecType(enum.IntEnum):
    REGULAR_SPEC = 17
    GROUP_SPEC = 51


class StampCell(enum.IntEnum):
    GAUGE_CELL = 3
    CONSTRUCTOR_NAME_CELL = 110
    CHECKER_NAME_CELL = 115
    CONSTRUCTOR_DATE_CELL = 130


class DocNotOpened(Exception):
    pass


class NoExecutions(Exception):
    pass


class NotSupportedSpecType(Exception):
    pass


class CoreKompass:
    """
    Класс запуска, создания API и выхода из компаса
    """
    def __init__(self):
        self.kompas_api7_module = None
        self.application = None
        self.app = None
        self.const = None

    def set_kompas_api(self):
        pythoncom.CoInitialize()
        self.kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        self.application = self.kompas_api7_module.IApplication(
            Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(
                self.kompas_api7_module.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch
            )
        )
        self.app = self.application.Application
        self.const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
        self.app.HideMessage = self.const.ksHideMessageYes

    @property
    def application_stream(self):
        # нужно отдавать всегда новый стрим т.к в разных потоках не может работать конвертированный до этого com объект
        return pythoncom.CoMarshalInterThreadInterfaceInStream(
            pythoncom.IID_IDispatch, self.application)

    def exit_kompas(self):
        if self.is_kompas_open():
            self.app.HideMessage = self.const.ksHideMessageNo
            if not self.app.Visible:  # если компас в невидимом режиме
                self.app.Quit()  # закрываем компас

    def is_kompas_open(self):
        value = object.__getattribute__(self, "application")
        if value is None:
            return False
        return True

    def collect_thread_api(self, thread_api: Type[T]) -> T:
        dict_of_stream_objects = {}
        for field in thread_api._fields:
            dict_of_stream_objects[field] = getattr(self, field)
        return thread_api(**dict_of_stream_objects)

    @classmethod
    def convert_stream_to_com(cls, stream_obj):
        return Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(stream_obj, pythoncom.IID_IDispatch))

    def __getattribute__(self, attr):
        if object.__getattribute__(self, attr) is None:
            if attr in ["kompas_api7_module", "application", "const", "app"]:
                # создание происходит не сразу а лишь при вызове,
                # так как открытие компаса занимает некоторое время
                self.set_kompas_api()
        return object.__getattribute__(self, attr)


class KompasAPI:
    def __init__(self, kompas_api_data: ThreadKompasAPI):
        self.kompas_api7_module = kompas_api_data.kompas_api7_module
        application = CoreKompass.convert_stream_to_com(kompas_api_data.application_stream)
        self.app = application.Application
        self.docs = self.app.Documents
        self.const = kompas_api_data.const

    def get_document_api(self, file_path: FilePath):
        doc = self.docs.Open(file_path, False, False)  # открываем документ, в невидимом режиме для записи
        if doc is None:
            raise DocNotOpened
        if os.path.splitext(file_path)[1] == ".cdw":  # если чертёж, то используем интерфейс для чертежа
            doc2d = self.kompas_api7_module.IKompasDocument2D(
                doc._oleobj_.QueryInterface(self.kompas_api7_module.IKompasDocument2D.CLSID, pythoncom.IID_IDispatch))
        else:  # если спецификация, то используем интерфейс для спецификации
            doc2d = self.kompas_api7_module.ISpecificationDocument(
                doc._oleobj_.QueryInterface
                (self.kompas_api7_module.ISpecificationDocument.CLSID, pythoncom.IID_IDispatch))
        return doc, doc2d

    @contextmanager
    def get_kompas_file_data(self, file_path: FilePath):
        doc = None
        try:
            doc, doc_2d = self.get_document_api(file_path)
            yield doc_2d
        except DocNotOpened:
            _, filename = os.path.split(file_path)
            yield f"Не удалось открыть файл {filename} возможно файл создан в более новой версии \n"
        finally:
            if doc is not None:
                doc.Close(self.const.kdDoNotSaveChanges)

    @contextmanager
    def get_draw_stamp(self, file_path: FilePath):
        with self.get_kompas_file_data(file_path) as doc_2d:
            draw_stamp = doc_2d.LayoutSheets.Item(0).Stamp  # массив листов документа
            yield draw_stamp

    @staticmethod
    def get_line_section(line) -> int:
        return line.Section

    @staticmethod
    def get_spc_description(spc_descriptions):
        return spc_descriptions.Item(0).Objects

    @staticmethod
    def _get_oformlenie(i_layout_sheet):
        return i_layout_sheet.LayoutStyleNumber

    def create_spc_object(self, doc_2d, doc):
        oformlenie = self._get_oformlenie(doc_2d.LayoutSheets.Item(0))  # считываем номер оформления документа
        spc_descriptions = doc.SpecificationDescriptions
        spc_description = self.get_spc_description(spc_descriptions)
        return oformlenie, spc_description


class Converter:
    def __init__(self, kompas_api_data: ThreadKompasAPI):
        kompas_api_data = CoreKompass.convert_stream_to_com(kompas_api_data.application_stream)
        self.application = kompas_api_data.Application
        self.kompas_api5_module = Dispatch("Kompas.Application.5")
        self.converter = self.set_converter()

    def set_converter(self):
        converter = self.application.Converter(self.kompas_api5_module.ksSystemPath(5) + r"\Pdf2d.dll")
        converter_parameters_module = gencache.EnsureModule("{31EBF650-BD38-43EC-892B-1F8AC6C14430}", 0, 1, 0)
        converter_parameters = converter_parameters_module. \
            IPdf2dParam(converter.ConverterParameters(0)._oleobj_.
                        QueryInterface(converter_parameters_module.IPdf2dParam.CLSID, pythoncom.IID_IDispatch))

        converter_parameters.CutByFormat = True  # обрезать по формату
        converter_parameters.EmbedFonts = True  # встроить шрифты
        converter_parameters.GrayScale = True  # оттенки серого
        converter_parameters.MultiPageOutput = True  # сохранять все страницы
        converter_parameters.MultipleFormat = 1
        converter_parameters.Resolution = 300  # разрешение ( на векторные пдф ваще не влияет никак)
        converter_parameters.Scale = 1.0  # масшта
        return converter

    def convert_draw_to_pdf(self, draw_path: FilePath, pdf_file_path: FilePath):
        self.converter.Convert(draw_path, pdf_file_path, 0, False)


class OboznSearcher:
    """
        Класс поиска обозначений из спецификации групового и обычного типа
    """
    MAXIMUM_COLUMN_NUMBER = 10
    SIZE_COLUMN = (1, 1, 0)
    OBOZN_COLUMN = (4, 1, 0)
    NAME_COLUMN = (5, 1, 0)

    def __init__(
            self,
            spec_path: FilePath,
            kompas_api: KompasAPI,
            spec_obozn: DrawObozn = None,
            without_sub_assembles: bool = False,
    ):
        self.spec_path = spec_path
        self.without_sub_assembles = without_sub_assembles
        self.spec_obozn = spec_obozn
        self.kompas_api = kompas_api

        self.doc, self.doc_2d = self.kompas_api.get_document_api(self.spec_path)
        self.oformlenie, self.spc_description = self.kompas_api.create_spc_object(self.doc, self.doc_2d)
        self._check_oformlenie()

        self.assembly_draws: list[DrawData] = []
        self.spec_draws: list[DrawData] = []
        self.detail_draws: list[DrawData] = []
        self.errors: list[str] = []

    def _check_oformlenie(self):
        if self.oformlenie not in [SpecType.REGULAR_SPEC, SpecType.GROUP_SPEC]:  # работаем только в простой и групповой
            self.doc.Close(self.kompas_api.const.kdDoNotSaveChanges)
            raise NotSupportedSpecType(
                f"\n{os.path.basename(self.spec_path)} - указан не поддерживаемый тип спефецикации"
            )

    def need_to_select_executions(self):
        return self.oformlenie == SpecType.GROUP_SPEC

    def get_all_spec_executions(self) -> dict[DrawExecution | str, int]:
        executions = {}
        correct_spc_lines = self._get_all_lines_with_correct_type([ObjectType.OBOZN_ISP])
        for spc_line in correct_spc_lines:
            try:
                for column_number in range(1, self.MAXIMUM_COLUMN_NUMBER):
                    execution_in_draw = self._get_cell_data(spc_line, (6, column_number, 0))
                    if execution_in_draw == "-":
                        execution_in_draw = "Базовое исполнение"

                    if self._verify_column_not_empty(column_number):
                        executions[DrawExecution(execution_in_draw)] = column_number
            except AttributeError:
                break

        if not executions:
            raise NoExecutions

        executions["Все исполнения"] = WITHOUT_EXECUTION

        return executions

    def _get_all_lines_with_correct_type(self, correct_line_type: list[int]) -> list:
        return [line for line in self.spc_description if line.ObjectType in correct_line_type]

    @staticmethod
    def _get_cell_data(line, cell_coordinates: tuple[int, int, int]) -> str:
        return line.Columns.Column(*cell_coordinates).Text.Str.strip()

    def _verify_column_not_empty(self, column_number: int) -> bool:
        for spc_line in self.spc_description:
            if spc_line.ObjectType in [ObjectType.REGULAR_LINE, 2] \
                    and self._get_cell_data(spc_line, self.OBOZN_COLUMN) \
                    and self._get_cell_data(spc_line, (6, column_number, 0)):
                return True
        return False

    def get_obozn_from_specification(self, column_numbers: list[int] = None) -> tuple[list[SpecSectionData], list[str]]:
        if self.oformlenie == SpecType.REGULAR_SPEC:
            response = self._get_obozn_from_simple_specification()
        else:  # group sdec
            if not column_numbers:
                column_numbers = self._get_column_numbers()
            response = self._get_obozn_from_group_spec(column_numbers)
        spec_data, errors = response
        return spec_data, errors

    def _get_obozn_from_simple_specification(self) -> tuple[list[SpecSectionData], list[str]]:
        for line in self.spc_description:
            draw_obozn, draw_name, size = self._get_line_obozn_name_size(line)
            if not draw_obozn:
                continue

            line_section = self.kompas_api.get_line_section(line)
            if line_section == DrawType.ASSEMBLY_DRAW:
                self.assembly_draws.append(DrawData(draw_obozn=draw_obozn, draw_name=draw_name))
                if self.without_sub_assembles:
                    break
            else:
                self._parse_lines_for_detail_and_spec(draw_obozn, draw_name, size, line_section)

        return self._create_spec_output(), self.errors

    def _parse_lines_for_detail_and_spec(self, draw_obozn: DrawObozn, draw_name: DrawName, size, line_section):

        if line_section == DrawType.SPEC_DRAW:
            if self.is_detail(draw_obozn) is True:
                message = f"\nВозможно указана деталь в качестве спецификации " \
                          f"-> {self.spec_path} ||| {draw_obozn} \n"
                self.errors.append(message)
                return
            self.spec_draws.append(DrawData(draw_obozn=draw_obozn, draw_name=draw_name))

        elif line_section == DrawType.DETAIL and size != "б/ч" and size != "бч":
            if draw_obozn.lower() in DIFFERENCE_ON_DRAW_MESSAGES or "гост" in draw_name or "гост" in draw_obozn:
                return
            self.detail_draws.append(DrawData(draw_obozn=draw_obozn, draw_name=draw_name))

    @staticmethod
    def is_detail(obozn: str) -> bool:
        if re.match(r".+\.\d[1-9][А-ЯA-Z]?(?:-0[1-9])?[А-ЯA-Z]?$", obozn.strip(), re.I):
            return True
        return False

    def _get_column_numbers(self) -> list[int]:
        def get_obozn_execution() -> DrawExecution:
            spec_obozn = self.spec_obozn if self.spec_obozn else DrawObozn("")

            draw_info = fetch_obozn_and_execution(spec_obozn)
            if draw_info is None:
                _execution = DrawExecution("-")
            else:
                _, _execution, _ = draw_info
            return _execution
        execution = get_obozn_execution()
        # нужно возвращать лист так как по нему итерируемся позже
        column_numbers = [self._get_column_number_by_execution(execution)]
        return column_numbers

    def _get_column_number_by_execution(self, execution: DrawExecution) -> int:
        def rid_of_extra_chars(_execution: DrawExecution) -> DrawExecution:
            if len(_execution) > 1:
                _execution = DrawExecution(_execution.lstrip("-").lstrip("0").strip())
            return _execution

        execution = rid_of_extra_chars(execution)
        correct_spc_lines = self._get_all_lines_with_correct_type([ObjectType.OBOZN_ISP])

        for spc_line in correct_spc_lines:
            try:
                for column_number in range(1, self.MAXIMUM_COLUMN_NUMBER):
                    execution_in_draw = rid_of_extra_chars(
                        DrawExecution(self._get_cell_data(spc_line, (6, column_number, 0)))
                    )
                    if execution_in_draw == execution and self._verify_column_not_empty(column_number):
                        return column_number
            except AttributeError:
                raise NoExecutions
            break
        raise NoExecutions

    def _get_obozn_from_group_spec(
            self,
            column_numbers: list[int],
    ) -> tuple[list[SpecSectionData], list[str]]:
        registered_obozn: list[DrawObozn] = []
        correct_lines = self._get_all_lines_with_correct_type([ObjectType.REGULAR_LINE, 2])

        for line in correct_lines:
            draw_obozn, draw_name, size = self._get_line_obozn_name_size(line)
            if not draw_obozn or draw_obozn in registered_obozn:
                continue

            line_section = self.kompas_api.get_line_section(line)
            if line_section == DrawType.ASSEMBLY_DRAW:
                self.assembly_draws.append(DrawData(draw_obozn=draw_obozn, draw_name=draw_name))
                if self.without_sub_assembles:
                    break
                registered_obozn.append(draw_obozn)

            for column_number in column_numbers:
                if not self._get_cell_data(line, (6, column_number, 0)):
                    continue
                registered_obozn.append(draw_obozn)
                self._parse_lines_for_detail_and_spec(draw_obozn, draw_name, size, line_section)

        return self._create_spec_output(), self.errors

    def _create_spec_output(self) -> list[SpecSectionData]:
        spec_data = []
        for draw_type, draw_names in [
                (DrawType.ASSEMBLY_DRAW, self.assembly_draws),
                (DrawType.DETAIL, self.detail_draws),
                (DrawType.SPEC_DRAW, self.spec_draws)]:
            if not draw_names:
                continue
            spec_data.append(SpecSectionData(draw_type=draw_type, draw_names=draw_names))

        return spec_data

    def _get_line_obozn_name_size(self, line) -> tuple[DrawObozn | None, DrawName | None, str | None]:
        try:
            draw_obozn = DrawObozn(self._get_cell_data(line, self.OBOZN_COLUMN).lower().replace(" ", ""))
            draw_name = DrawName(self._get_cell_data(line, self.NAME_COLUMN).lower().replace(" ", ""))
            size = self._get_cell_data(line, self.SIZE_COLUMN).lower().replace(" ", "")
        except AttributeError:
            return None, None, None
        return draw_obozn, draw_name, size


class SpecPathChecker(OboznSearcher):
    def __init__(self, spec_path: FilePath, kompas_api: KompasAPI, execution: DrawExecution):
        super(SpecPathChecker, self).__init__(spec_path, kompas_api)
        self.execution = execution

    def verify_its_correct_spec_path(self) -> bool:
        def _look_for_line_with_str() -> bool:
            for line in self.spc_description.Objects:
                if line.ObjectType in [1, 2] and self._get_cell_data(line, self.OBOZN_COLUMN):
                    obozn = self._get_cell_data(line, self.OBOZN_COLUMN).strip().lower().replace(" ", "")
                    if obozn in DIFFERENCE_ON_DRAW_MESSAGES:
                        return True
            return False

        if self.oformlenie == SpecType.REGULAR_SPEC:
            response = _look_for_line_with_str()
            self.doc.Close(self.kompas_api.const.kdDoNotSaveChanges)
            return response

        try:
            column_number = self._get_column_number_by_execution(self.execution)
        finally:
            self.doc.Close(self.kompas_api.const.kdDoNotSaveChanges)

        if column_number:
            return True
        return False


def fetch_obozn_and_execution(draw_obozn: DrawObozn) -> tuple[DrawObozn, DrawExecution, str | None] | None:
    draw_info = re.search(r"(.+)(?:-)(0[13579][а-яёa]?$)", draw_obozn, re.I)
    if not draw_info:
        return None

    obozn, execution = draw_info.groups()
    modification_symbol = ""
    if not execution[-1].isdigit():
        execution, modification_symbol, = execution[:-1], execution[-1]
    return obozn, execution, modification_symbol

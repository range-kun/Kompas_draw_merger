# -*- coding: utf-8 -*-
from __future__ import annotations

import enum
import os
import re
from contextlib import contextmanager
from typing import TypeVar, Generic

import pythoncom
from win32com.client import Dispatch, gencache


from schemas import DrawType, SpecSectionData, DrawData, FilePath, DrawObozn, DrawExecution, DrawName, ThreadKompasAPI

DIFFERENCE_ON_DRAW_MESSAGES = ["различияисполненийпосборочномучертежу", "отличияисполненийпосборочномучертежу"]

SIZE_COLUMN = (1, 1, 0)
OBOZN_COLUMN = (4, 1, 0)
NAME_COLUMN = (5, 1, 0)
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

    def collect_thread_api(self, thread_api: Generic[T]) -> T:
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
                # создание происходит не сразу а лишь при вызове, так как открытие компаса занимает некоторое время
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
        if os.path.splitext(file_path)[1] == '.cdw':  # если чертёж, то используем интерфейс для чертежа
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
            yield f'Не удалось открыть файл {filename} возможно файл создан в более новой версии \n'
        finally:
            if doc is not None:
                doc.Close(self.const.kdDoNotSaveChanges)

    @contextmanager
    def get_draw_stamp(self, file_path: FilePath):
        with self.get_kompas_file_data(file_path) as doc_2d:
            draw_stamp = doc_2d.LayoutSheets.Item(0).Stamp  # массив листов документа
            yield draw_stamp

    @staticmethod
    def is_detail(obozn: str) -> bool:
        detail_group = re.match(r".+\.\d[1-9][А-Я]?(?:-0[1-9])?[А-Я]?$", obozn.strip(), re.I)
        if detail_group:
            return True


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
    def __init__(
            self,
            spec_path: FilePath,
            kompas_api: KompasAPI,
            spec_obozn: DrawObozn = None,
            only_document_list: bool = False,
    ):
        self.spec_path = spec_path
        self.only_document_list = only_document_list
        self.spec_obozn = spec_obozn
        self.kompas_api = kompas_api

        self.doc, self.doc_2d = self.kompas_api.get_document_api(self.spec_path)
        self.oformlenie, self.spc_description = self._create_spc_object()
        self._check_oformlenie()

        self.assembly_draws: list[DrawData] = []
        self.spec_draws: list[DrawData] = []
        self.detail_draws: list[DrawData] = []
        self.errors: list[str] = []

    def need_to_select_executions(self):
        return self.oformlenie == SpecType.GROUP_SPEC

    def get_all_spec_executions(self) -> dict[DrawExecution, int]:
        executions = {}

        for spc_line in self.spc_description.Objects:
            if spc_line.ObjectType != ObjectType.OBOZN_ISP:
                continue
            try:
                for column_number in range(1, 10):
                    execution_in_draw = spc_line.Columns.Column(6, column_number, 0).Text.Str.strip()
                    if execution_in_draw == '-':
                        execution_in_draw = 'Базовое исполнение'
                    if self._verify_column_not_empty(column_number):
                        executions[DrawExecution(execution_in_draw)] = column_number
            except AttributeError:
                break
        if not executions:
            raise NoExecutions

        executions['Все исполнения'] = WITHOUT_EXECUTION
        return executions

    def _verify_column_not_empty(self, column_number: int) -> bool | None:
        for spc_line in self.spc_description.Objects:
            if spc_line.ObjectType in [1, 2] \
                    and spc_line.Columns.Column(*OBOZN_COLUMN) \
                    and spc_line.Columns.Column(6, column_number, 0).Text.Str.strip():
                return True

    def _create_spc_object(self):
        i_layout_sheet = self.doc_2d.LayoutSheets.Item(0)
        oformlenie = i_layout_sheet.LayoutStyleNumber  # считываем номер оформления документа
        spc_descriptions = self.doc.SpecificationDescriptions
        spc_description = spc_descriptions.Item(0)
        return oformlenie, spc_description

    def _check_oformlenie(self):
        if self.oformlenie not in [SpecType.REGULAR_SPEC, SpecType.GROUP_SPEC]:  # работаем только в простой и групповой
            self.doc.Close(self.kompas_api.const.kdDoNotSaveChanges)
            raise NotSupportedSpecType(
                f"\n{os.path.basename(self.spec_path)} - указан не поддерживаемый тип спефецикации"
            )

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
        for line in self.spc_description.Objects:
            draw_obozn, draw_name, size = self._get_line_obozn_name_size(line)
            if not draw_obozn and not draw_name:
                continue

            if line.Section == DrawType.ASSEMBLY_DRAW:
                self.assembly_draws.append(DrawData(draw_obozn=draw_obozn, draw_name=draw_name))
                if self.only_document_list:
                    return self._create_spec_output(), self.errors
            self._parse_lines_for_detail_and_spec(line)

        return self._create_spec_output(), self.errors

    def _parse_lines_for_detail_and_spec(self, line):
        draw_obozn, draw_name, size = self._get_line_obozn_name_size(line)
        if line.Section == DrawType.SPEC_DRAW:
            if self.kompas_api.is_detail(draw_obozn):
                message = f"\nВозможно указана деталь в качестве спецификации " \
                          f"-> {self.spec_path} ||| {draw_obozn} \n"
                self.errors.append(message)
                return
            self.spec_draws.append(DrawData(draw_obozn=draw_obozn, draw_name=draw_name))

        elif line.Section == DrawType.DETAIL and size != "б/ч" and size != "бч":
            if draw_obozn.lower() in DIFFERENCE_ON_DRAW_MESSAGES or "гост" in draw_name or "гост" in draw_obozn:
                return
            self.detail_draws.append(DrawData(draw_obozn=draw_obozn, draw_name=draw_name))

    def _get_column_numbers(self) -> list[int]:
        def get_obozn_execution() -> DrawExecution:
            draw_info = fetch_obozn_and_execution(self.spec_obozn)

            if draw_info is None:
                _execution = "-"
            else:
                _, _execution, _ = draw_info
            return _execution

        execution = get_obozn_execution()
        column_numbers = [self._get_column_number_by_execution(execution)]
        return column_numbers

    def _get_column_number_by_execution(self, execution: DrawExecution) -> int | None:
        def rid_of_extra_chars(_execution: DrawExecution) -> DrawExecution:
            if len(_execution) > 1:
                _execution = _execution.lstrip('-').lstrip('0').strip()
            return _execution

        execution = rid_of_extra_chars(execution)

        for spc_line in self.spc_description.Objects:
            if spc_line.ObjectType != ObjectType.OBOZN_ISP:
                continue
            try:
                for column_number in range(1, 10):
                    execution_in_draw = rid_of_extra_chars(
                        spc_line.Columns.Column(6, column_number, 0).Text.Str.strip()
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

        for line in self.spc_description.Objects:
            draw_obozn, draw_name, size = self._get_line_obozn_name_size(line)
            if not draw_obozn or draw_obozn in registered_obozn:
                continue

            if line.Section == DrawType.ASSEMBLY_DRAW:
                self.assembly_draws.append(DrawData(draw_obozn=draw_obozn, draw_name=draw_name))
                if self.only_document_list:
                    return self._create_spec_output(), self.errors
                registered_obozn.append(draw_obozn)

            for column_number in column_numbers:
                if line.ObjectType not in [1, 2] or not line.Columns.Column(6, column_number, 0).Text.Str:
                    continue
                registered_obozn.append(draw_obozn)
                self._parse_lines_for_detail_and_spec(line)

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

    @staticmethod
    def _get_line_obozn_name_size(line) -> tuple[DrawObozn | None, DrawName | None, str | None]:
        try:
            draw_obozn = line.Columns.Column(*OBOZN_COLUMN).Text.Str.strip().lower().replace(' ', '')
            draw_name = line.Columns.Column(*NAME_COLUMN).Text.Str.strip().lower()
            size = line.Columns.Column(*SIZE_COLUMN).Text.Str.strip().lower()
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
                if line.ObjectType in [1, 2] and line.Columns.Column(*OBOZN_COLUMN):
                    obozn = line.Columns.Column(*OBOZN_COLUMN).Text.Str.strip().lower().replace(' ', '')
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


def fetch_obozn_and_execution(draw_obozn: DrawObozn) -> tuple[DrawObozn, DrawExecution, str | None] | None:
    draw_info = re.search(r"(.+)(?:-)(0[13579][а-яёa]?$)", draw_obozn, re.I)
    if not draw_info:
        return

    obozn, execution = draw_info.groups()
    modification_symbol = ""
    if not execution[-1].isdigit():
        execution, modification_symbol, = execution[:-1], execution[-1]
    return obozn, execution, modification_symbol

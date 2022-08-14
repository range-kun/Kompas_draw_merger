# -*- coding: utf-8 -*-
from __future__ import annotations

import enum
import os
import re
from contextlib import contextmanager
from typing import Any

import pythoncom
from win32com.client import Dispatch, gencache

from schemas import DrawType, SpecSectionData, DrawData, FilePath, DrawObozn, DrawExecution, DrawName

DIFFERENCE_ON_DRAW_MESSAGES = ["различияисполненийпосборочномучертежу", "отличияисполненийпосборочномучертежу"]

SIZE_COLUMN = (1, 1, 0)
OBOZN_COLUMN = (4, 1, 0)
NAME_COLUMN = (5, 1, 0)
WITHOUT_EXECUTION = 1000


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


def get_kompas_api7():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    pythoncom.CoInitializeEx(0)
    api = module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface
                              (module.IKompasAPIObject.CLSID, pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0)
    return module, api, const.constants


def get_kompas_api5():
    module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
    api = module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface
                              (module.KompasObject.CLSID, pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0)
    return module, api, const.constants


def set_converter(app, kompas_object):
    iConverter = app.Converter(kompas_object.ksSystemPath(5) + r"\Pdf2d.dll")  # интерфейс для сохранения в PDF
    converter_parameters_module = gencache.EnsureModule("{31EBF650-BD38-43EC-892B-1F8AC6C14430}", 0, 1, 0)
    converter_parameters = converter_parameters_module.\
        IPdf2dParam(iConverter.ConverterParameters(0)._oleobj_.
                    QueryInterface(converter_parameters_module.IPdf2dParam.CLSID, pythoncom.IID_IDispatch))
    converter_parameters.CutByFormat = True  # обрезать по формату
    converter_parameters.EmbedFonts = True  # встроить шрифты
    converter_parameters.GrayScale = True  # оттенки серого
    converter_parameters.MultiPageOutput = True  # сохранять все страницы
    converter_parameters.MultipleFormat = 1
    converter_parameters.Resolution = 300  # разрешение ( на векторные пдф ваще не влияет никак)
    converter_parameters.Scale = 1.0  # масшта
    return iConverter


def get_kompas_settings(application, kompas_object):
    app = application.Application
    i_converter = set_converter(app, kompas_object)
    docs = app.Documents
    return app, i_converter, docs


def exit_kompas(app):
    if not app.Visible:  # если компас в невидимом режиме
        app.Quit()  # закрываем компас


def is_detail(obozn: str) -> bool:
    detail_group = re.match(r".+\.\d[1-9][А-Я]?(?:-0[1-9])?[А-Я]?$", obozn.strip(), re.I)
    if detail_group:
        return True


def get_all_spec_executions(spc_description) -> dict[DrawExecution, int]:
    executions = {}

    for spc_line in spc_description.Objects:
        if spc_line.ObjectType != ObjectType.OBOZN_ISP:
            continue
        try:
            for column_number in range(1, 10):
                execution_in_draw = spc_line.Columns.Column(6, column_number, 0).Text.Str.strip()
                if execution_in_draw == '-':
                    execution_in_draw = 'Базовое исполнение'
                if verify_column_not_empty(spc_description, column_number):
                    executions[DrawExecution(execution_in_draw)] = column_number
        except AttributeError:
            break
    if not executions:
        raise NoExecutions

    executions['Все исполнения'] = WITHOUT_EXECUTION
    return executions


def get_obozn_from_specification(
        spec_path: FilePath, *,
        only_document_list: bool =False,
        column_number: list[int] = None,
        initial_call: bool=False,
        spec_obozn: DrawObozn = None
) -> tuple[list[SpecSectionData], Any, list[str]]:
    kompas_api7_module, application, const = get_kompas_api7()
    app = application.Application
    app.HideMessage = const.ksHideMessageYes

    response = create_spc_object(spec_path, app, kompas_api7_module, const)
    oformlenie, spc_description, doc = response

    if oformlenie not in [SpecType.REGULAR_SPEC, SpecType.GROUP_SPEC]:  # работаем только в простой и групповой
        doc.Close(const.kdDoNotSaveChanges)
        app.HideMessage = const.ksHideMessageNo
        raise NotSupportedSpecType(f"\n{os.path.basename(spec_path)} - указан не поддерживаемый тип спефецикации")

    try:
        if oformlenie == SpecType.REGULAR_SPEC:
            response = get_obozn_from_simple_specification(
                spc_description,
                spec_path,
                only_document_list
            )
        else:  # group sdec
            if initial_call:  # выбрать пользователю какое исполнение он хочет слить
                column_number = yield get_all_spec_executions(spc_description)
            response = get_obozn_from_group_spec(
                spc_description,
                spec_path,
                only_document_list,
                column_number,
                spec_obozn
            )
    finally:
        doc.Close(const.kdDoNotSaveChanges)
        app.HideMessage = const.ksHideMessageNo

    spec_data, errors = response
    return spec_data, application, errors


def get_draw_obozn_name_size(line) -> tuple[DrawObozn | None, DrawName | None, str | None]:
    try:
        draw_obozn = line.Columns.Column(*OBOZN_COLUMN).Text.Str.strip().lower().replace(' ', '')
        draw_name = line.Columns.Column(*NAME_COLUMN).Text.Str.strip().lower()
        size = line.Columns.Column(*SIZE_COLUMN).Text.Str.strip().lower()
    except AttributeError:
        return None, None, None
    return draw_obozn, draw_name, size


def get_obozn_from_simple_specification(
        spc_description,
        spec_path: FilePath,
        only_document_list: bool
) -> tuple[list[SpecSectionData], list[str]]:

    assembly_draws: list[DrawData] = []
    spec_draws: list[DrawData] = []
    detail_draws: list[DrawData] = []
    errors: list[str] = []

    for line in spc_description.Objects:
        draw_obozn, draw_name, size = get_draw_obozn_name_size(line)
        if not draw_obozn and not draw_name:
            continue

        if line.Section == DrawType.ASSEMBLY_DRAW:
            assembly_draws.append(DrawData(draw_obozn=draw_obozn, draw_name=draw_name))
            if only_document_list:
                return create_spec_output(assembly_draws), errors

        elif line.Section == DrawType.SPEC_DRAW:
            if is_detail(draw_obozn):
                message = f"\nВозможно указана деталь в качестве спецификации " \
                          f"-> {spec_path} ||| {draw_obozn} \n"
                errors.append(message)
                continue
            spec_draws.append(DrawData(draw_obozn=draw_obozn, draw_name=draw_name))

        elif line.Section == DrawType.DETAIL and size != 'б/ч' and size != 'бч':
            if draw_obozn.lower() in DIFFERENCE_ON_DRAW_MESSAGES or "гост" in draw_name or "гост" in draw_obozn:
                continue
            detail_draws.append(DrawData(draw_obozn=draw_obozn, draw_name=draw_name))

    return create_spec_output(assembly_draws, spec_draws, detail_draws), errors


def get_obozn_from_group_spec(
        spc_description,
        spec_path: str,
        only_document_list: bool,
        column_numbers: list[int] = None,
        spec_obozn: DrawObozn = None
) -> tuple[list[SpecSectionData], list[str]]:

    assembly_draws: list[DrawData] = []
    spec_draws: list[DrawData] = []
    detail_draws: list[DrawData] = []
    errors: list[str] = []
    registered_obozn: list[DrawObozn] = []

    def get_execution() -> DrawExecution:
        draw_info = fetch_obozn_and_execution(spec_obozn)

        if draw_info is None:
            _execution = "-"
        else:
            _, _execution = draw_info
            if not _execution[:-1].isdigit():  # если есть буква в конце
                _execution = _execution[:-1]
        return _execution

    if not column_numbers:
        execution = get_execution()
        column_numbers = [get_column_number(spc_description, execution)]

    for line in spc_description.Objects:
        draw_obozn, draw_name, size = get_draw_obozn_name_size(line)
        if not draw_obozn or draw_obozn in registered_obozn:
            continue

        if line.Section == DrawType.ASSEMBLY_DRAW:
            assembly_draws.append(DrawData(draw_obozn=draw_obozn, draw_name=draw_name))
            registered_obozn.append(draw_obozn)
            if only_document_list:
                return create_spec_output(assembly_draws), errors

        for column_number in column_numbers:
            if line.ObjectType not in [1, 2] or not line.Columns.Column(6, column_number, 0).Text.Str:
                continue

            registered_obozn.append(draw_obozn)
            if line.Section == DrawType.SPEC_DRAW:
                if is_detail(draw_obozn):
                    message = f"\nВозможно указана деталь в качестве спецификации " \
                              f"-> {spec_path} ||| {draw_obozn} \n"
                    errors.append(message)
                    continue
                spec_draws.append(DrawData(draw_obozn=draw_obozn, draw_name=draw_name))

            elif line.Section == DrawType.DETAIL and size != "б/ч" and size != "бч":
                if draw_obozn.lower() in DIFFERENCE_ON_DRAW_MESSAGES or "гост" in draw_name or "гост" in draw_obozn:
                    detail_draws.append(DrawData(draw_obozn=draw_obozn, draw_name=draw_name))

    return create_spec_output(assembly_draws, spec_draws, detail_draws), errors


def create_spec_output(
        assembly_draws: list[DrawData],
        spec_draws: list[DrawData] = None,
        detail_draws: list[DrawData] = None) -> list[SpecSectionData]:
    spec_data = []
    for draw_type, draw_names in [
            (DrawType.ASSEMBLY_DRAW, assembly_draws),
            (DrawType.DETAIL, detail_draws),
            (DrawType.SPEC_DRAW, spec_draws)]:
        if not draw_names:
            continue
        spec_data.append(SpecSectionData(draw_type=draw_type, draw_names=draw_names))

    return spec_data


def fetch_obozn_and_execution(draw_obozn: DrawObozn) -> tuple[str | Any] | None:
    draw_info = re.search(r"(.+)(?:-)(0[13579][а-яёa]?$)", draw_obozn, re.I)
    if not draw_info:
        return
    return draw_info.groups()


def create_spc_object(spec_path: FilePath, app, kompas_api7_module, const):
    docs = app.Documents
    doc = docs.Open(spec_path, False, False)  # открываем документ, в невидимом режиме для записи

    try:
        doc2d = kompas_api7_module.ISpecificationDocument(
            doc._oleobj_.QueryInterface
            (kompas_api7_module.ISpecificationDocument.CLSID, pythoncom.IID_IDispatch)
        )
    except Exception:
        try:
            doc.Close(const.kdDoNotSaveChanges)
        except AttributeError:
            pass
        raise FileExistsError(f"\n{os.path.basename(spec_path)} "
                              f"- ошибка при открытии спецификации обновите базу данных")

    i_layout_sheet = doc2d.LayoutSheets.Item(0)
    oformlenie = i_layout_sheet.LayoutStyleNumber  # считываем номер оформления документа
    spc_descriptions = doc.SpecificationDescriptions
    spc_description = spc_descriptions.Item(0)
    return oformlenie, spc_description, doc


def verify_its_correct_spec_path(spec_path: FilePath, execution: DrawExecution):
    def look_for_line_with_str(list_of_messages: list[str]) -> bool:
        for line in spc_description.Objects:
            if line.ObjectType in [1, 2] and line.Columns.Column(*OBOZN_COLUMN):
                obozn = line.Columns.Column(*OBOZN_COLUMN).Text.Str.strip().lower().replace(' ', '')
                if obozn in list_of_messages:
                    return True

    kompas_api7_module, application, const = get_kompas_api7()
    app = application.Application
    app.HideMessage = const.ksHideMessageYes
    response = create_spc_object(spec_path, app, kompas_api7_module, const)

    oformlenie, spc_description, doc = response
    if oformlenie == SpecType.REGULAR_SPEC:
        response = look_for_line_with_str(DIFFERENCE_ON_DRAW_MESSAGES)
        doc.Close(const.kdDoNotSaveChanges)
        return response

    try:
        column_number = get_column_number(spc_description, execution)
    finally:
        doc.Close(const.kdDoNotSaveChanges)

    if column_number:
        return True


def get_column_number(spc_description, execution: DrawExecution) -> int | None:
    for spc_line in spc_description.Objects:
        if spc_line.ObjectType != ObjectType.OBOZN_ISP:
            continue
        try:
            for column_number in range(1, 10):
                execution_in_draw = spc_line.Columns.Column(6, column_number, 0).Text.Str.strip()
                execution_in_draw = execution_in_draw[1:] if len(execution_in_draw) > 1 \
                    else execution_in_draw  # исполнение м.б - или -01
                if execution_in_draw == execution and verify_column_not_empty(spc_description, column_number):
                    return column_number
        except AttributeError:
            raise NoExecutions
        break
    raise NoExecutions


def verify_column_not_empty(spc_description, column_number: int) -> bool | None:
    for spc_line in spc_description.Objects:
        if spc_line.ObjectType in [1, 2] \
                and spc_line.Columns.Column(*OBOZN_COLUMN) \
                and spc_line.Columns.Column(6, column_number, 0).Text.Str.strip():
            return True


def get_document_api(file_path: str, docs, kompas_api7_module):
    doc = docs.Open(file_path, False, False)  # открываем документ, в невидимом режиме для записи
    if doc is None:
        raise DocNotOpened
    if os.path.splitext(file_path)[1] == '.cdw':  # если чертёж, то используем интерфейс для чертежа
        doc2d = kompas_api7_module.IKompasDocument2D(
            doc._oleobj_.QueryInterface (kompas_api7_module.IKompasDocument2D.CLSID, pythoncom.IID_IDispatch))
    else:  # если спецификация, то используем интерфейс для спецификации
        doc2d = kompas_api7_module.ISpecificationDocument(
            doc._oleobj_.QueryInterface
            (kompas_api7_module.ISpecificationDocument.CLSID, pythoncom.IID_IDispatch))
    return doc, doc2d


@contextmanager
def get_kompas_file_data(file_path: str, docs, kompas_api7_module, const):
    try:
        doc, doc2d = get_document_api(file_path, docs, kompas_api7_module)
        yield doc, doc2d
    except DocNotOpened:
        doc = None
        _, filename = os.path.split(file_path)
        yield f'Не удалось открыть файл {filename} возможно файл создан в более новой версии \n'
    finally:
        if doc is not None:
            doc.Close(const.kdDoNotSaveChanges)

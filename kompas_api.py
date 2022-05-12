# -*- coding: utf-8 -*-
from __future__ import annotations

import enum
import os
import time
import re
from typing import Optional, Any, Union

import pythoncom
from win32com.client import Dispatch, gencache

OBOZN_COLUMN = (4, 1, 0)
NAME_COLUMN = (5, 1, 0)
WITHOUT_EXECUTION = 1000


class ObjectType(enum.IntEnum):
    OBOZN_ISP = 4
    REGULAR_LINE = 1


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


def get_right_api(file, docs, kompas_api7_module):
    doc = docs.Open(file, False, False)  # открываем документ, в невидимом режиме для записи
    if os.path.splitext(file)[1] == '.cdw':  # если чертёж, то используем интерфейс для чертежа
        doc2d = kompas_api7_module.IKompasDocument2D(doc._oleobj_.QueryInterface
                                                     (kompas_api7_module.IKompasDocument2D.CLSID,
                                                      pythoncom.IID_IDispatch))
    else:  # если спецификация, то используем интерфейс для спецификации
        doc2d = kompas_api7_module.ISpecificationDocument(
            doc._oleobj_.QueryInterface
            (kompas_api7_module.ISpecificationDocument.CLSID, pythoncom.IID_IDispatch))
    return doc, doc2d


def date_to_seconds(date_string):
    if len(date_string.split('.')[-1]) == 2:
        new_date = date_string.split('.')
        new_date[-1] = '20' + new_date[-1]
        date_string = '.'.join(new_date)
    struct_date = time.strptime(date_string, "%d.%m.%Y")
    return time.mktime(struct_date)


def get_draws_from_specification(spec_path: str, *,
                                 only_document_list=False,
                                 draw_obozn: str = None,
                                 by_button_click=False,
                                 column_number: list[int] = None):
    kompas_api7_module, application, const = get_kompas_api7()
    app = application.Application
    app.HideMessage = const.ksHideMessageYes
    file_name = os.path.basename(spec_path).split()[0]

    response = create_spc_object(spec_path, app, kompas_api7_module, const)
    if type(response) == str:
        return response
    oformlenie, spc_description, doc = response

    if oformlenie not in [17, 51]:  # 17, 51 - работаем только в простой и групповой
        return f"Указан не поддерживаемый тип спецификации {spec_path} \n"

    if oformlenie == 17:
        response = get_paths_from_simple_specification(
            spc_description,
            spec_path,
            only_document_list
        )
    else:
        if by_button_click:  # выбрать пользователю какое исполнение он хочет слить
            executions = get_all_spec_executions(spc_description)
            return executions
        response = get_paths_from_group_spec(
            spc_description,
            spec_path,
            only_document_list,
            draw_obozn,
            column_number
        )
        if type(response) == str:  # совпадение не найден
            doc.close()
            return f"\nИсполнение {draw_obozn} для групповой спецификации не найдено\n"

    if only_document_list:
        documentation_draws, errors = response
        return documentation_draws, application, errors
    documentation_draws, detail_draws, assembly_draws, errors = response

    files: dict[str: list[tuple[str, str]]] = {}
    if documentation_draws:
        files["Сборочные чертежи"] = documentation_draws
    if detail_draws:
        files['Детали'] = detail_draws
    if assembly_draws:
        files["Сборочные единицы"] = assembly_draws
    app.HideMessage = const.ksHideMessageNo
    doc.Close(const.kdDoNotSaveChanges)
    return files, application, errors


def get_paths_from_simple_specification(
        spc_description,
        spec_path: str,
        only_document_list: bool
):
    documentation_draws: list[tuple[str, str]] = []
    assembly_draws: list[tuple[str, str]] = []
    detail_draws: list[tuple[str, str]] = []
    errors: list[str] = []

    for i in spc_description.Objects:
        if (i.ObjectType == ObjectType.REGULAR_LINE or i.ObjectType == 2) \
                and i.Columns.Column(4, 1, 0):
            obozn = i.Columns.Column(*OBOZN_COLUMN).Text.Str.strip().lower().replace(' ', '')
            name = i.Columns.Column(*NAME_COLUMN).Text.Str.strip().lower()
            if not obozn:
                continue
            try:
                size = i.Columns.Column(1, 1, 0).Text.Str.strip().lower()
            except AttributeError:
                size = None

            if i.Section == 5:
                documentation_draws.append((obozn, name))
                if only_document_list:
                    return documentation_draws, errors

            elif i.Section == 15:
                if is_detail(obozn):
                    message = f"\nВозможно указана деталь в качестве спецификации " \
                              f"-> {spec_path} ||| {obozn} \n"
                    errors.append(message)
                    continue
                assembly_draws.append((obozn, name))

            elif i.Section == 20 and size != 'б/ч' and size != 'бч':
                if obozn == "различияисполненийпосборочномучертежу":
                    continue
                if "гост" in name or 'гост' in obozn:
                    continue
                detail_draws.append((obozn, name))
    return documentation_draws, detail_draws, assembly_draws, errors


def get_paths_from_group_spec(
        spc_description,
        spec_path: str,
        only_document_list: bool,
        draw_obozn: str,
        column_numbers: list[int] = None
):
    documentation_draws: list[tuple[str, str]] = []
    assembly_draws: list[tuple[str, str]] = []
    detail_draws: list[tuple[str, str]] = []
    errors: list[str] = []
    registered_obozn: list[str] = []

    if not column_numbers:
        draw_info = fetch_obozn_execution_and_name(draw_obozn)

        if draw_info is None:
            execution = "-"
        else:
            _, execution = draw_info
            if not execution[:-1].isdigit():  # если есть буква в конце
                execution = execution[:-1]
        response = get_column_number(spc_description, execution)
        if type(response) == str:
            return response
        column_numbers = [response]

    for i in spc_description.Objects:
        for column_number in column_numbers:
            try:
                obozn = i.Columns.Column(*OBOZN_COLUMN).Text.Str.strip().lower().replace(' ', '')
                name = i.Columns.Column(*NAME_COLUMN).Text.Str.strip().lower()
            except AttributeError:
                continue

            if i.Section == 5:
                if obozn in registered_obozn:
                    continue
                documentation_draws.append((obozn, name))
                registered_obozn.append(obozn)
                if only_document_list:
                    return documentation_draws, errors

            elif i.ObjectType in [1, 2] and obozn and i.Columns.Column(6, column_number, 0).Text.Str:
                if obozn in registered_obozn:
                    continue
                registered_obozn.append(obozn)
                try:
                    size = i.Columns.Column(1, 1, 0).Text.Str.strip().lower()
                except AttributeError:
                    size = None

                if i.Section == 15:
                    if is_detail(obozn):
                        message = f"\nВозможно указана деталь в качестве спецификации " \
                                  f"-> {spec_path} ||| {obozn} \n"
                        errors.append(message)
                        continue
                    assembly_draws.append((obozn, name))

                elif i.Section == 20 and size != 'б/ч' and size != 'бч':
                    detail_draws.append((obozn, name))

    return documentation_draws, detail_draws, assembly_draws, errors


def exit_kompas(app):
    if not app.Visible:  # если компас в невидимом режиме
        app.Quit()  # закрываем компас


def is_detail(obozn: str) -> bool:
    detail_group = re.match(r".+\.\d[1-9][А-Я]?(?:-0[1-9])?[А-Я]?$", obozn.strip(), re.I)
    if detail_group:
        return True


def fetch_obozn_execution_and_name(draw_obozn: str) -> Optional[tuple[str | Any]]:
    draw_info = re.search(r"(.+)(?:-)(0[13579][а-яёa]?$)", draw_obozn, re.I)
    if not draw_info:
        return
    return draw_info.groups()


def get_all_spec_executions(spc_description) -> Union[str, dict[str, int]]:
    executions: dict[str, int] = {}
    return_executions: dict[str, int] = {}

    for spc_line in spc_description.Objects:
        if spc_line.ObjectType == ObjectType.OBOZN_ISP:
            try:
                for column_number in range(1, 10):
                    execution_in_draw = spc_line.Columns.Column(6, column_number, 0).Text.Str.strip()
                    if execution_in_draw == '-':
                        execution_in_draw = 'Базовое исполнение'
                    executions[execution_in_draw] = column_number
            except AttributeError:
                break
    if not executions:
        return 'Ошибка открытия спецификации'

    for column_text, column_number in executions.items():
        is_this_column_not_emty = verify_column_not_empty(spc_description, column_number)
        if is_this_column_not_emty:
            return_executions[column_text] = column_number
    return_executions['Все исполнения'] = WITHOUT_EXECUTION
    return return_executions


def get_column_number(spc_description, execution: str) -> Union[str, int]:
    for spc_line in spc_description.Objects:
        if spc_line.ObjectType == ObjectType.OBOZN_ISP:
            try:
                for column_number in range(1, 10):
                    execution_in_draw = spc_line.Columns.Column(6, column_number, 0).Text.Str.strip()
                    execution_in_draw = execution_in_draw[1:] if len(execution_in_draw) > 1 \
                        else execution_in_draw  # исполнение м.б - или -01
                    if execution_in_draw == execution:
                        break
            except AttributeError:
                return "Совпадение не найдено"
            break

    is_this_column_not_emty = verify_column_not_empty(spc_description, column_number)
    if is_this_column_not_emty:
        return column_number
    return "Совпадение не найдено"


def verify_column_not_empty(spc_description, column_number):
    for spc_line in spc_description.Objects:
        if spc_line.ObjectType in [1, 2] and spc_line.Columns.Column(*OBOZN_COLUMN):

            if spc_line.Columns.Column(6, column_number, 0).Text.Str.strip():
                return True


def create_spc_object(spec_path: str, app, kompas_api7_module, const):
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
        return f"Ошибка при открытии спецификации {spec_path} \n"

    i_layout_sheet = doc2d.LayoutSheets.Item(0)
    oformlenie = i_layout_sheet.LayoutStyleNumber  # считываем номер оформления документа
    spc_descriptions = doc.SpecificationDescriptions
    spc_description = spc_descriptions.Item(0)
    return oformlenie, spc_description, doc


def verify_its_correct_spec_path(spec_path: str, execution: str):
    def look_for_line(line: str) -> bool:
        for i in spc_description.Objects:
            if i.ObjectType in [1, 2] and i.Columns.Column(*OBOZN_COLUMN):
                obozn = i.Columns.Column(*OBOZN_COLUMN).Text.Str.strip().lower().replace(' ', '')
                if line == obozn:
                    return True

    kompas_api7_module, application, const = get_kompas_api7()
    app = application.Application
    app.HideMessage = const.ksHideMessageYes
    response = create_spc_object(spec_path, app, kompas_api7_module, const)
    if type(response) == str:
        return

    oformlenie, spc_description, doc = response
    if oformlenie != 51:
        response = look_for_line('различияисполненийпосборочномучертежу')
        doc.Close(const.kdDoNotSaveChanges)
        return response
    response = get_column_number(spc_description, execution)
    if type(response) == str:
        doc.Close(const.kdDoNotSaveChanges)
        return

    doc.Close(const.kdDoNotSaveChanges)
    return True

# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import time
import re

import pythoncom
from win32com.client import Dispatch, gencache


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


def get_draws_from_specification(spec_path: str, only_document_list=False):
    files: dict[str: list[tuple[str, str]]] = {}
    documentation_draws: list[tuple[str, str]] = []
    assembly_draws: list[tuple[str, str]] = []
    detail_draws: list[tuple[str, str]] = []
    kompas_api7_module, application, const = get_kompas_api7()
    app = application.Application
    app.HideMessage = const.ksHideMessageYes
    docs = app.Documents
    doc = docs.Open(spec_path, False, False)  # открываем документ, в невидимом режиме для записи
    errors: list[str] = []

    try:
        doc2d = kompas_api7_module.ISpecificationDocument(
            doc._oleobj_.QueryInterface
            (kompas_api7_module.ISpecificationDocument.CLSID, pythoncom.IID_IDispatch))
    except Exception:
        return f"Ошибка при открытии спецификации {spec_path} \n"

    i_layout_sheet = doc2d.LayoutSheets.Item(0)
    oformlenie = i_layout_sheet.LayoutStyleNumber  # считываем номер оформления документа
    spc_descriptions = doc.SpecificationDescriptions

    if oformlenie == 17:  # 17 - работаем только в простой спецификации
        spc_description = spc_descriptions.Item(0)
        for i in spc_description.Objects:
            if i.ObjectType in [1, 2] and i.Columns.Column(4, 1, 0):
                obozn = i.Columns.Column(4, 1, 0).Text.Str.strip().lower().replace(' ', '')
                name = i.Columns.Column(5, 1, 0).Text.Str.strip().lower()
                if not obozn:
                    continue
                try:
                    size = i.Columns.Column(1, 1, 0).Text.Str.strip().lower()
                except AttributeError:
                    size = None

                if i.Section == 5:
                    documentation_draws.append((obozn, name))
                    if only_document_list:
                        return documentation_draws, application, errors

                elif i.Section == 15:
                    if is_detail(obozn):
                        message = f"\nВозможно указана деталь в качестве спецификации " \
                                  f"-> {spec_path} ||| {obozn} \n"
                        errors.append(message)
                        continue
                    assembly_draws.append((obozn, name))

                elif i.Section == 20 and size != 'б/ч' and size != 'бч':
                    detail_draws.append((obozn, name))
    if documentation_draws:
        files["Сборочные чертежи"] = documentation_draws
    if detail_draws:
        files['Детали'] = detail_draws
    if assembly_draws:
        files["Сборочные единицы"] = assembly_draws
    app.HideMessage = const.ksHideMessageNo
    return files, application, errors


def exit_kompas(app):
    if not app.Visible:  # если компас в невидимом режиме
        app.Quit()  # закрываем компас


def is_detail(obozn: str) -> bool:
    detail_group = re.match(r".+\.\d[1-9][А-Я]?(?:-0[1-9])?[А-Я]?$", obozn.strip(), re.I)
    if detail_group:
        return True

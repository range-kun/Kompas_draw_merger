# -*- coding: utf-8 -*-
#|Фамилия

import pythoncom
from win32com.client import Dispatch, gencache
import os
import time


def get_kompas_api7():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    api = module.IApplication(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0)
    return module, api, const.constants


def get_kompas_api5():
    module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
    api = module.KompasObject(
        Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(module.KompasObject.CLSID, pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0)
    return module, api, const.constants

kompas_api7_module, application, const = get_kompas_api7()
kompas6_api5_module, kompas_object, kompas6_constants = get_kompas_api5()
app = application.Application
docs = app.Documents
iConverter = application.Converter(kompas_object.ksSystemPath(5) + "\Pdf2d.dll")


def cdw_to_pdf(files, directory_pdf):
    number = 0
    for file in files:
        number += 1
        iConverter.Convert(file, directory_pdf + "\\" +
                        f'{number} ' + os.path.basename(file) + ".pdf", 0, False)


def filter_by_date(files, date_1, date_2):
    draw_list = []
    app.HideMessage = const.ksHideMessageNo  # отключаем отображение сообщений Компас, отвечая на всё "нет"
    for file in files:  # структура обработки для каждого документа
        doc = docs.Open(file, False, False)  # открываем документ, в невидимом режиме для записи
        if os.path.splitext(file)[1] == '.cdw':  # если чертёж, то используем интерфейс для чертежа
            doc2D = kompas_api7_module.IKompasDocument2D(doc._oleobj_.QueryInterface
                                               (kompas_api7_module.IKompasDocument2D.CLSID, pythoncom.IID_IDispatch))
        else:  # если спецификация, то используем интерфейс для спецификации
            doc2D = kompas_api7_module.ISpecificationDocument(
                doc._oleobj_.QueryInterface
                (kompas_api7_module.ISpecificationDocument.CLSID, pythoncom.IID_IDispatch))
        iStamp = doc2D.LayoutSheets.Item(0).Stamp  # массив листов документа
        date_in_stamp = iStamp.Text(130).Str
        if date_in_stamp:
            date_1_in_seconds, date_2_in_seconds = sorted([date_1, date_2])
            try:
                date_in_stamp = date_to_seconds(iStamp.Text(130).Str)
            except:
                continue
            if date_in_stamp in range(date_1_in_seconds, date_2_in_seconds+1):
                draw_list.append(file)
        doc.Close(const.kdDoNotSaveChanges)
    return draw_list

def date_to_seconds(date_string):
    if len(date_string.split('.')[-1]) == 2:
        new_date = date_string.split('.')
        new_date[-1] = '20' + new_date[-1]
        date_string = '.'.join(new_date)
        struct_date = time.strptime(date_string, "%d.%m.%Y")
    return time.mktime(struct_date)

def exit_kompas():
    if kompas_object.Visible == False:  # если компас в невидимом режиме
        kompas_object.Quit()  # закрываем компас





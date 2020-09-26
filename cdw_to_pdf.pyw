# -*- coding:  utf-8 -*-
#https://forum.ascon.ru/index.php/topic,32760.0.html

import os, PyPDF2, itertools
import get_kompas_api
import time

#  Подключим описание интерфейсов API7
kompas_api7_module, application = get_kompas_api.get_kompas_api7()
app = application.Application
docs = app.Documents
_, kompas_object, _ = get_kompas_api.get_kompas_api5()
iConverter = application.Converter(kompas_object.ksSystemPath(5) + "\Pdf2d.dll")


def pdf_to_cdw(file, directory_pdf):
	iConverter.Convert(file,
						directory_pdf + "\\" + os.path.basename(file) + ".pdf", 0, False)


def merge_pdf(directory):
	# Получаем список файлов в переменную files
	files = os.listdir(directory)
	spec_files = [i for i in files if os.path.splitext(i)[0].endswith('.spw')]
	cdw_files = [i for i in files if os.path.splitext(i)[0].endswith('.cdw')]
	files = spec_files + cdw_files
	merger = PyPDF2.PdfFileMerger()

	for filename in files:
		merger.append(fileobj=open(os.path.join(directory, filename), 'rb'))

	merger.write(open(os.path.join(directory, 'book.pdf'), 'wb'))


def create_pdf(directory):
	kompas_ext = ['.spw', '.cdw']
	directory_pdf = '%s\pdf %s' % (directory, time.strftime("%d.%m.%Y"))
	if not os.path.exists(directory_pdf):
		os.makedirs(directory_pdf)
	for (this_dir, _, files_here) in os.walk(directory):
		files = (os.path.join(this_dir, i) for i in files_here if os.path.splitext(i)[1] in kompas_ext)
		for file in files:
			pdf_to_cdw(file, directory_pdf)
	merge_pdf(directory_pdf)

create_pdf(r'C:\Users\Pitoohon-User\Desktop\01 Рама\02 Кронштейн')


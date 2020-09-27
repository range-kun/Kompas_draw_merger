# -*- coding:  utf-8 -*-
#https://forum.ascon.ru/index.php/topic,32760.0.html

import os, PyPDF2
import get_kompas_api
import time
import shutil

#  Подключим описание интерфейсов API7
kompas_api7_module, application = get_kompas_api.get_kompas_api7()
app = application.Application
docs = app.Documents
_, kompas_object, _ = get_kompas_api.get_kompas_api5()
iConverter = application.Converter(kompas_object.ksSystemPath(5) + "\Pdf2d.dll")


def pdf_to_cdw(file, directory_pdf, number):
	iConverter.Convert(file,
						directory_pdf + "\\" + f'{number} ' + os.path.basename(file) + ".pdf", 0, False)


def merge_pdf(directory):
	# Получаем список файлов в переменную files
	files = sorted(os.listdir(directory), key=lambda file: int(file.split()[0]))
	merger = PyPDF2.PdfFileMerger()
	for filename in files:
		merger.append(fileobj=open(os.path.join(directory, filename), 'rb'))
	merger.write(open(os.path.join(os.path.dirname(directory), f'{os.path.basename(directory)}.pdf'), 'wb'))
	for file in merger.inputs:
		file[0].close()


def create_folders(directory):
	main_name = [os.path.splitext(i)[0] for i in os.listdir(directory) if os.path.splitext(i)[1] == '.cdw'][0]
	base_pdf_dir = f'{directory}\pdf'
	directory_pdf = r'%s\pdf\%s - 01 %s' % (directory, main_name, time.strftime("%d.%m.%Y"))

	if not os.path.exists(base_pdf_dir):
		os.makedirs(base_pdf_dir)
	if os.path.exists(directory_pdf):
		today_update = len([i for i in os.listdir(base_pdf_dir) if i.endswith(f'{time.strftime("%d.%m.%Y")}')])
		today_update = str(today_update+1) if today_update > 8 else '0'+str(today_update+1)
		directory_pdf = r'%s\pdf\%s - %s %s' % (directory, main_name, today_update, time.strftime("%d.%m.%Y"))
		os.makedirs(directory_pdf)
	else:
		os.makedirs(directory_pdf)
	return directory_pdf


def create_pdf(directory):
	kompas_ext = ['.spw', '.cdw']
	number = 0
	directory_pdf = create_folders(directory)
	for (this_dir, _, files_here) in os.walk(directory):
		if not this_dir.endswith('Былое'):
			files = [os.path.join(this_dir, i) for i in files_here if os.path.splitext(i)[1] in kompas_ext]
			for file in sorted(files, key=lambda fil: os.path.splitext(fil)[1], reverse=True):
				number += 1
				pdf_to_cdw(file, directory_pdf, number)
	merge_pdf(directory_pdf)
	shutil.rmtree(directory_pdf)

create_pdf(r'C:\Users\Pitoohon-User\Desktop\01 Рама')

from docx import Document
from docx.shared import Pt
import os
import sys
import shutil
from datetime import datetime
from tqdm import tqdm
import time

file = os.path.join(os.getcwd(), 'template.docx')
temp_dir_path = os.path.join(os.getcwd(), 'temp')
output_file = os.path.join(temp_dir_path, 'output.docx')
year = datetime.strftime(datetime.now(), '%y')
docx_files = []

if not os.path.exists(file):
    sys.exit('Не обнаружен файл template.docx')

if not os.path.exists(temp_dir_path):
    os.mkdir(temp_dir_path)

# Удаление файлов из временной директории
for files in os.listdir(temp_dir_path):
    temp_file = os.path.join(temp_dir_path, files)
    try:
        shutil.rmtree(temp_file)
    except OSError:
        os.remove(temp_file)

if (len(sys.argv) == 3) and (int(sys.argv[2]) >= int(sys.argv[1])):
    start_num = int(sys.argv[1])
    end_num = int(sys.argv[2])
else:
    start_num = int(input('Введите начальный номер договора: '))
    end_num = int(input('Введите конечный номер договора: '))
    if start_num > end_num:
        exit('Начальный номер не может быть больше конечного')

doc = Document(file)

# Создание временных файлов
for num in tqdm(range(start_num, end_num + 1)):
    p = doc.paragraphs[0]
    p.text = ''
    run = p.add_run(f'ДОГОВОР ОК-{num:03}-{year}')
    run.font.bold = True
    run.font.name = 'Arial'
    run.font.size = Pt(10)

    temp_file = os.path.join(temp_dir_path, f'{num:04}.docx')
    doc.save(temp_file)
    time.sleep(0.1)

for filedocx in os.listdir(temp_dir_path):
    if filedocx.endswith('.docx'):
        docx_files.append(os.path.join(temp_dir_path, filedocx))

merged_document = Document(docx_files[0])

# Склеивание временных файлов в один итоговый
for index, file in enumerate(docx_files):
    sub_doc = Document(file)
    for element in sub_doc.element.body:
        merged_document.element.body.append(element)

    if index != 0:
        sub_doc = Document(file)
        for element in sub_doc.element.body:
            merged_document.element.body.append(element)


merged_document.save(output_file)
time.sleep(0.5)
os.startfile(output_file, 'print')

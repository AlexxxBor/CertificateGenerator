import os
from enum import Enum

import openpyxl
from docxtpl import DocxTemplate


class CertType(Enum):
    MAIN_CERT = "сертификат"
    DIST_CERT = "сертификат с отличием"


def get_template(cert_type: CertType) -> DocxTemplate:
    if cert_type == CertType.DIST_CERT:
        return DocxTemplate("templates/tpl_with_distinction.docx")
    return DocxTemplate("templates/tpl_certificate.docx")


def get_dir(directory: str):
    if not os.path.exists(directory):
        os.makedirs(directory)
        return directory
    return directory


def make_certificate(tpl_data: dict, cert_type: CertType, path: str) -> None:
    certificate = get_template(cert_type)
    file_name = f"{context["surname"]} {context["name"]} {context["patronymic"]}"
    certificate.render(tpl_data)
    certificate.save(f"{path}/{file_name}.docx")


WORKING_DIR = get_dir("сертификаты")
CERT_DATA_SHEET = "cert_data"

course_dir = WORKING_DIR
tpl_data_keys = ("surname", "name", "patronymic", "course", "mod", "hour", "cert", "number")

wb = openpyxl.load_workbook(filename="data/IT-куб.xlsx")

for sheet in wb.sheetnames:
    if sheet == CERT_DATA_SHEET:
        continue

    try:
        course_dir = get_dir(f"{WORKING_DIR}/{sheet}")
    except OSError as e:
        print(f"Не могу создать папку в {WORKING_DIR} для {sheet}. Возникла ошибка: {e}")
        break

    try:
        for row in wb[sheet].iter_rows(min_row=2):
            # TODO: обработка пустых ячеек
            tpl_data_values = tuple(cell.value for cell in row)
            context = {tpl_data_keys[i]: value for i, value in enumerate(tpl_data_values)}

            module = context.pop('mod')
            if module != "без модуля":
                context['course'] = f"{context['course']} ({module})"

            make_certificate(context, CertType(context["cert"]), course_dir)

    except Exception as e:
        print(f"При формировании набора данных возникла ошибка:\n{e}")

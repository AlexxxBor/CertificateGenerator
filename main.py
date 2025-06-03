import os
import openpyxl

from enum import Enum
from docxtpl import DocxTemplate
from docx2pdf import convert
from art import tprint
from openpyxl import Workbook


class CertType(Enum):
    MAIN_CERT = "сертификат"
    DIST_CERT = "сертификат с отличием"
    SUMMER_CERT = "летние смены"


def get_template(cert_type: CertType) -> DocxTemplate:
    if cert_type == CertType.DIST_CERT:
        return DocxTemplate("templates/tpl_with_distinction.docx")
    elif cert_type == CertType.SUMMER_CERT:
        return DocxTemplate("templates/tpl_certificate_it_summer.docx")
    return DocxTemplate("templates/tpl_certificate.docx")


def get_dir(directory: str):
    if not os.path.exists(directory):
        os.makedirs(directory)
        return directory
    return directory


def make_certificate(tpl_data: dict, cert_type: CertType, path: str) -> None:
    docx_dir = get_dir(f"{path}/docx")
    file_name = f"{tpl_data["surname"]} {tpl_data["name"]} {tpl_data["patronymic"]}"

    certificate = get_template(cert_type)
    certificate.render(tpl_data)
    certificate.save(f"{docx_dir}/{file_name}.docx")

    pdf_dir = get_dir(f"{path}/pdf")
    docx_file = f"{docx_dir}/{file_name}.docx"
    pdf_file = f"{pdf_dir}/{file_name}.pdf"
    convert(docx_file, pdf_file)


def pages_range(pages: str):
    result = []
    for item in pages.replace(' ', '').split(','):
        if '-' in item:
            start, end = map(int, item.split('-'))
            result.extend(range(start, end + 1))
        else:
            result.append(int(item))
    return result


def get_used_sheets(sheets_range: list[int], wb: Workbook):
    using_sheets_names = []
    for number, name in enumerate(wb.sheetnames, 1):
        if number in sheets_range:
            using_sheets_names.append(name)
    return using_sheets_names


WORKING_DIR = get_dir("сертификаты")
CERT_DATA_SHEET = "cert_data"

tpl_data_keys = ("surname", "name", "patronymic", "course", "mod", "hour", "cert", "number")
wb = openpyxl.load_workbook(filename="data/IT-куб.xlsx")


def main():
    for num, sheet in enumerate(wb.sheetnames, 1):
        if sheet == CERT_DATA_SHEET:
            continue
        print(f'[{num}]{sheet}; ', end='')

    sheets_range = pages_range(input("\nКакие листы использовать? "))
    work_sheets = get_used_sheets(sheets_range, wb)

    tprint('starting...')
    count = 0

    for sheet in wb.sheetnames:
        if sheet not in work_sheets:
            continue

        try:
            course_dir = get_dir(f"{WORKING_DIR}/{sheet}")
        except OSError as e:
            print(f"Не могу создать папку курса в папке '{WORKING_DIR}' для листа '{sheet}'.")
            print(f"Возникла ошибка: {e}")
            break

        try:
            for row in wb[sheet].iter_rows(min_row=2):
                tpl_data_values = tuple(cell.value for cell in row)

                if None in tpl_data_values:
                    break

                context = {tpl_data_keys[i]: value for i, value in enumerate(tpl_data_values)}

                edu_module = context.pop('mod')
                if edu_module != "без модуля":
                    context['course'] = f"{context['course']} ({edu_module})"

                try:
                    make_certificate(context, CertType(context["cert"]), course_dir)
                    count += 1
                except Exception as e:
                    print(f"При создании сертификата возникла ошибка: {e}")

        except Exception as e:
            print(f"При формировании набора данных возникла ошибка: {e}")

    tprint('done!')
    input("Нажмите любую кнопку, чтобы закрыть это окно.")


if __name__ == "__main__":
    main()

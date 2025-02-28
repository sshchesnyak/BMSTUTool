import datetime
import os
import tkinter
import pyzipper
from tkinter import filedialog

import PyPDF2 as ppdf
import pdf2docx as cv
import docx2pdf as ccv
import win32_setctime as wsc

source = tkinter.Tk()
source.withdraw()

def open_check_pdf():
    source.wm_deiconify()
    source_path = filedialog.askopenfilename(title="Выберите исходный pdf файл")
    source.withdraw()
    if source_path.endswith(".pdf"):
        return True, source_path
    else:
        print("Ошибка! Выбранный файл не имеет расширения pdf")
        return False, source_path

def open_check_docx():
    source.wm_deiconify()
    source_path = filedialog.askopenfilename(title="Выберите исходный docx файл")
    source.withdraw()
    if source_path.endswith(".docx"):
        return True, source_path
    else:
        print("Ошибка! Выбранный файл не имеет расширения docx")
        return False, source_path

def open_check_zip():
    source.wm_deiconify()
    source_path = filedialog.askopenfilename(title="Выберите исходный zip файл")
    source.withdraw()
    if source_path.endswith(".zip"):
        return True, source_path
    else:
        print("Ошибка! Выбранный файл не имеет расширения zip")
        return False, source_path

def open_check_pdfs():
    source.wm_deiconify()
    source_paths = filedialog.askopenfilenames(title="Выберите исходные pdf файлы для объединения")
    source.withdraw()
    for source_path in source_paths:
        if not source_path.endswith(".pdf"):
            print(f"Ошибка! Выбранный файл {source_path} не имеет расширения pdf")
            return False, source_paths
    return True, source_paths


def border_worker(border_string, total):
    ends = border_string.split("-")
    left_border = ends[0].strip()
    right_border = ends[1].strip()
    if left_border.isdigit() and right_border.isdigit():
        left_border = int(left_border)
        right_border = int(right_border)
        if (left_border <= right_border) and (left_border > 0) and (right_border <= total):
            return True, left_border, right_border
        else:
            print("Ошибка ввода! Страницы вне границ документа или не в возрастающем порядке")
            return False, left_border, right_border
    else:
        print("Ошибка ввода! Границы диапазона не являются числами")
        return False, left_border, right_border

def date_str_checker(date_str):
    if len(date_str) > 0 and date_str.find(",") != -1:
        datetime_array = date_str.split(",")
        if len(datetime_array) == 5:
            for dt in datetime_array:
                dt = dt.strip()
                if not dt.isdigit():
                    print("Ошибка ввода! Элемент даты создания не является числом")
                    return False, datetime_array
            year = int(datetime_array[0])
            month = int(datetime_array[1])
            day = int(datetime_array[2])
            hour = int(datetime_array[3])
            minute = int(datetime_array[4])
            new_datetime = datetime.datetime(year=year, month=month, day=day, hour=hour, minute=minute)
            timestamp = new_datetime.timestamp()
            return True, timestamp
        else:
            print("Ошибка ввода! Неверное число элементов даты")
            return False, date_str
    else:
        print("Ошибка ввода! Формат данных не соответствует запрошенному")
        return False, date_str

def ind_splitter():
    is_pdf, source_path = open_check_pdf()
    if is_pdf:
        document = ppdf.PdfReader(source_path)
        writer = ppdf.PdfWriter()
        source.wm_deiconify()
        terminal_path = filedialog.askdirectory(title="Выберите конечную директорию", mustexist=True)
        source.withdraw()
        total_count = len(document.pages)
        print(f"В работе страница 0/{total_count}")
        page_count = 1
        for i in range(0, total_count):
            print(f"В работе страница {page_count}/{total_count}")
            writer.add_page(document.pages[i])
            with open(terminal_path + "/" + str(page_count) + ".pdf", "wb") as f:
                writer.write(f)
            page_count += 1
            writer.close()
            writer = ppdf.PdfWriter()

def group_splitter():
    is_pdf, source_path = open_check_pdf()
    check = True
    if is_pdf:
        document = ppdf.PdfReader(source_path)
        total_count = len(document.pages)
        left_borders = []
        right_borders = []
        print("Выберите диапазоны страниц в формате: 1-2;7-21")
        page_pool_str = input()
        if len(page_pool_str) > 0:
            if page_pool_str.find(";") != -1:
                page_pool = page_pool_str.split(";")
                for rang in page_pool:
                    check, left_border, right_border = border_worker(rang, total_count)
                    if check:
                        left_borders.append(left_border)
                        right_borders.append(right_border)
            elif page_pool_str.find("-") != -1:
                check, left_border, right_border = border_worker(page_pool_str, total_count)
                if check:
                    left_borders.append(left_border)
                    right_borders.append(right_border)
            else:
                print("Ошибка ввода! Формат данных не соответствует запрошенному")
                check = False
        if check:
            source.wm_deiconify()
            terminal_path = filedialog.askdirectory(title="Выберите конечную директорию", mustexist=True)
            source.withdraw()
            for i in range(0,len(left_borders)):
                print(f"В работе диапазон {i+1}/{len(left_borders)}")
                new_document = ppdf.PdfWriter()
                left_border = left_borders[i]
                right_border = right_borders[i]
                for j in range(left_border, right_border+1):
                    print(f"В работе страница {j-left_border+1}/{right_border-left_border+1} диапазона {i+1}/{len(left_borders)}")
                    new_document.add_page(document.pages[j-1])
                    with open(terminal_path + "/" + str(left_border) + "-" + str(right_border) + ".pdf", "wb") as f:
                        new_document.write(f)
                new_document.close()

def merger():
    is_pdf, source_paths = open_check_pdfs()
    writer = ppdf.PdfWriter()
    check = True
    if is_pdf:
        print(f"Вы добавили следующие файлы: {source_paths}")
        print("Введите последовательность файлов в итоговом файле в виде: 1,3,2. Альтернативно нажмите на клавишу ENTER чтобы оставить порядок по умолчанию")
        order = input()
        if len(order) > 0:
            if order.find(",") != -1:
                order = order.split(",")
                for i in order:
                    i= i.strip()
                    if i.isdigit():
                        i = int(i)
                        writer.append(source_paths[i-1])
                    else:
                        print("Ошибка ввода! Границы диапазона не являются числами")
                        check = False
                        break
            else:
                print("Ошибка ввода! Формат данных не соответствует запрошенному")
                check = False
        else:
            for source_path in source_paths:
                writer.append(source_path)
        if check:
            source.wm_deiconify()
            terminal_path = filedialog.asksaveasfilename(title="Введите название итогового файла", defaultextension=".pdf")
            source.withdraw()
            with open(terminal_path, "wb") as f:
                writer.write(f)

def pdf2docx():
    is_pdf, source_path = open_check_pdf()
    if is_pdf:
        source.wm_deiconify()
        terminal_path = filedialog.asksaveasfilename(title="Введите название итогового файла", defaultextension=".docx")
        source.withdraw()
        conv = cv.Converter(source_path)
        conv.convert(terminal_path)
        conv.close()

def docx2pdf():
    is_docx, source_path = open_check_docx()
    if is_docx:
        source.wm_deiconify()
        terminal_path = filedialog.asksaveasfilename(title="Введите название итогового файла", defaultextension=".pdf")
        source.withdraw()
        ccv.convert(source_path, terminal_path)

def mod_created_dt():
    is_pdf, source_path = open_check_pdf()
    if is_pdf:
        print("Введите желаемую дату создания в следующем формате: ГГГГ,ММ,ДД,чч,мм")
        created_str = input()
        check_created, time_created = date_str_checker(created_str)
        if check_created:
            wsc.setctime(source_path, time_created)

def mod_mod_dt():
    is_pdf, source_path = open_check_pdf()
    if is_pdf:
        print("Введите желаемую дату изменения в следующем формате: ГГГГ,ММ,ДД,чч,мм")
        modified_str = input()
        check_modified, time_modified = date_str_checker(modified_str)
        if check_modified:
            os.utime(source_path, (time_modified, time_modified))

def create_protected_archive(default_pwd, default_dir, fold_id):
    source.wm_deiconify()
    source_folder = filedialog.askdirectory(title="Выберите архивируемую директорию", mustexist=True)
    if default_dir != "":
        terminal_folder = f"{default_dir}/{os.path.basename(source_folder)}{fold_id}.zip"
    else:
        terminal_folder = filedialog.asksaveasfilename(title="Введите название итогового файла", defaultextension=".zip")
    source.withdraw()
    if default_pwd != "":
        pwd = default_pwd
    else:
        print("Введите пароль для шифрования: ")
        pwd = input()
    if terminal_folder != "":
        if pwd != "":
            with pyzipper.AESZipFile(terminal_folder, "w", compression=pyzipper.ZIP_DEFLATED, encryption=pyzipper.WZ_AES) as zf:
                zf.pwd = pwd.encode("utf-8")
                for root, dirs, files in os.walk(source_folder):
                    for filename in files:
                        filepath = os.path.join(root, filename)
                        archive_path = os.path.relpath(filepath, source_folder)
                        zf.write(filepath, archive_path)
        else:
            print("Ошибка ввода! Пароль не может быть пустым")
    else:
        print("Ошибка ввода! Введенная директория не существует")

def create_protected_archives():
    terminate = False
    index = 0
    pwd = ""
    terminal_folder = ""
    print("Вы хотите, чтобы все созданные архивы были защищены одним паролем? (y/n)")
    one_pwd = input()
    if one_pwd == "y":
        print("Введите пароль для шифрования: ")
        pwd = input()
    print("Вы хотите, чтобы все созданные архивы были размещены в одной директории? (y/n)")
    one_dir = input()
    if one_dir == "y":
        source.wm_deiconify()
        terminal_folder = filedialog.askdirectory(title="Выберите местоположение итоговых файлов", mustexist=True)
        source.withdraw()
    while not terminate:
        create_protected_archive(pwd, terminal_folder, index)
        print("Желаете создать еще один архив? (y/n)")
        ans = input()
        if ans == "y":
            index += 1
        else:
            terminate = True

def open_protected_archive():
    is_zip, source_path = open_check_zip()
    if is_zip:
        source.wm_deiconify()
        terminal_folder = filedialog.askdirectory(title="Выберите местоположение итоговых файлов", mustexist=True)
        source.withdraw()
        pwd_correct = False
        while not pwd_correct:
            print("Введите пароль от архива: ")
            pwd = input()
            try:
                with pyzipper.AESZipFile(source_path) as zf:
                    zf.pwd = pwd.encode("utf-8")
                    zf.extractall(terminal_folder)
                    pwd_correct = True
            except:
                print("Пароль указан не верно. Хотите попробовать еще раз? (y/n)")
                choice = input()
                if choice == "y":
                    pwd_correct = False
                else:
                    pwd_correct = True


if __name__ == "__main__":
    exit_flag = False
    while not exit_flag:
        print("Добро пожаловать в BMSTUTool 1.0.")
        print("Выберите действие 1-4:")
        print("1. Разбить файл pdf на отдельные страницы")
        print("2. Разбить файл pdf по диапазонам")
        print("3. Объединить файлы pdf")
        print("4. Преобразовать из pdf в docx")
        print("5. Преобразовать docx в pdf")
        print("6. Изменить дату и время создания pdf файла")
        print("7. Изменить дату и время изменения pdf файла")
        print("8. Создание запороленного zip архива")
        print("9. Создание набора запороленных zip архивов")
        print("10. Извлечение данных из запороленного zip архива")
        print("11. Выйти из программы")
        option = input()
        match option:
            case "1":
                ind_splitter()
            case "2":
                group_splitter()
            case "3":
                merger()
            case "4":
                pdf2docx()
            case "5":
                docx2pdf()
            case "6":
                mod_created_dt()
            case "7":
                mod_mod_dt()
            case "8":
                create_protected_archive("", "", 0)
            case "9":
                create_protected_archives()
            case "10":
                open_protected_archive()
            case "11":
                exit_flag = True
                break
            case _:
                exit_flag = True
                break
        print("Желаете продолжить? (y/n)")
        option = input()
        if option == "y":
            os.system("cls")
        else:
            exit_flag = True





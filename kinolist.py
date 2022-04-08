from operator import truediv
import os
import glob
import re
from secrets import choice
import shutil  # сохранение файла
import sys
from copy import deepcopy  # копирование таблиц

import requests  # загрузка файла из сети
from docx import Document  # работа с docx
from docx.shared import Cm, Pt, RGBColor
from kinopoisk_unofficial.kinopoisk_api_client import KinopoiskApiClient
from kinopoisk_unofficial.request.films.film_request import FilmRequest
from kinopoisk_unofficial.request.staff.staff_request import StaffRequest
from kinopoisk.movie import Movie # api для поиска фильмов
from mutagen.mp4 import MP4, MP4Cover # работа тегами
from PIL import Image # работа с изображениями

import config

ver = '0.3.6'
api = config.api_key


# проверка авторизации
def isapiok(api):
    try:
        api_client = KinopoiskApiClient(api)
        request = FilmRequest(507)
        response = api_client.films.send_film_request(request)
    except:
        return False
    else:
        return True


# получение информации о фильме по kinopoisk id
def getFilminfo(film_code, api):
    api_client = KinopoiskApiClient(api)
    request_staff = StaffRequest(film_code)
    response_staff = api_client.staff.send_staff_request(request_staff)
    stafflist = []
    for i in range(0, 11):  # загружаем 11 персоналий (режиссер + 10 актеров)
        stafflist.append(response_staff.items[i].name_ru)
    request_film = FilmRequest(film_code)
    response_film = api_client.films.send_film_request(request_film)

    # с помощью регулярного выражения находим значение стран в кавычках ''
    countries = re.findall("'([^']*)'", str(response_film.film.countries))

    # имя файла
    filename = response_film.film.name_ru
    # очистка имени файла от запрещенных символов
    trtable = filename.maketrans("", "", '\/:*?"<>')
    filename = filename.translate(trtable)
    filmlist = [
        response_film.film.name_ru, response_film.film.year, response_film.film.rating_kinopoisk, countries,
        response_film.film.description, response_film.film.poster_url, filename
    ]
    return filmlist + stafflist


# заполнение таблицы в docx файле
def writeFilmtoTable(current_table, filminfo):
    paragraph = current_table.cell(0, 1).paragraphs[0]  # название фильма + рейтинг
    run = paragraph.add_run(str(filminfo[0]) + ' - ' + 'Кинопоиск ' + str(filminfo[2]))
    run.font.name = 'Arial'
    run.font.size = Pt(11)
    run.font.bold = True

    paragraph = current_table.cell(1, 1).add_paragraph()  # год
    run = paragraph.add_run(str(filminfo[1]))
    run.font.name = 'Arial'
    run.font.size = Pt(10)

    paragraph = current_table.cell(1, 1).add_paragraph()  # страна
    run = paragraph.add_run(', '.join(filminfo[3]))
    run.font.name = 'Arial'
    run.font.size = Pt(10)

    paragraph = current_table.cell(1, 1).add_paragraph()  # режиссер
    run = paragraph.add_run('Режиссер: ' + filminfo[7])
    run.font.name = 'Arial'
    run.font.size = Pt(10)

    paragraph = current_table.cell(1, 1).add_paragraph()

    paragraph = current_table.cell(1, 1).add_paragraph()  # в главных ролях
    run = paragraph.add_run('В главных ролях: ')
    run.font.color.rgb = RGBColor(255, 102, 0)
    run.font.name = 'Arial'
    run.font.size = Pt(10)
    run = paragraph.add_run(', '.join(filminfo[8:]))
    run.font.color.rgb = RGBColor(0, 0, 255)
    run.font.name = 'Arial'
    run.font.size = Pt(10)
    run.font.underline = True

    paragraph = current_table.cell(1, 1).add_paragraph()
    paragraph = current_table.cell(1, 1).add_paragraph()
    paragraph = current_table.cell(1, 1).add_paragraph()  # синопсис
    run = paragraph.add_run(filminfo[4])
    run.font.name = 'Arial'
    run.font.size = Pt(10)
    paragraph = current_table.cell(1, 1).add_paragraph()

    # загрузка постера
    image_url = filminfo[5]
    if not os.path.isdir("covers"):
        os.mkdir("covers")
    filename = str(filminfo[6] + '.jpg')
    file_path = './covers/' + str(filminfo[6] + '.jpg')
    resp = requests.get(image_url, stream=True)
    if resp.status_code == 200:
        resp.raw.decode_content = True
        with open(file_path, 'wb') as f:  # открываем файл для бинарной записи
            shutil.copyfileobj(resp.raw, f)
        # print('Постер загружен:', filename)
    else:
        print('Не удалось загрузить постер (' + image_url + ')')

    # изменение размера постера
    image = Image.open(file_path)
    width, height = image.size
    # обрезка до соотношения сторон 1x1.5
    if width > (height / 1.5):
        image = image.crop((((width - height / 1.5) / 2), 0, ((width - height / 1.5) / 2) + height / 1.5, height))
    image.thumbnail((360, 540))
    image.save(file_path)
    # запись постера в таблицу
    paragraph = current_table.cell(0, 0).paragraphs[1]
    run = paragraph.add_run()
    run.add_picture(file_path, width=Cm(7))


# копирование таблицы в указанный параграф
def copy_table_after(table, paragraph):
    tbl, p = table._tbl, paragraph._p
    new_tbl = deepcopy(tbl)
    p.addnext(new_tbl)


# клонирует первую таблицу в документе num раз
def cloneFirstTable(document:Document, num):
    template = document.tables[0]
    paragraph = document.paragraphs[0]
    for i in range(num):
        copy_table_after(template, paragraph)
        paragraph = document.add_paragraph()


# запись тегов в mp4 файл
def writeTagstoMp4(film):
    file_path = str(film[6] + '.mp4')
    if not os.path.isfile(file_path):
        print(f'Ошибка: Файл "{file_path}" не найден!')
        return
    video = MP4(file_path)
    video.delete()  # удаление всех тегов
    video["\xa9nam"] = film[0]  # title
    video["desc"] = film[4]  # description
    video["ldes"] = film[4]  # long description
    video["\xa9day"] = str(film[1])  #year
    cover = './covers/' + str(film[6] + '.jpg')
    with open(cover, "rb") as f:
        video["covr"] = [MP4Cover(f.read(), imageformat=MP4Cover.FORMAT_JPEG)]
    video.save()
    print(file_path + ' - тег записан')


# определение пути для запуска из автономного exe файла.
# pyinstaller cоздает временную папку, путь в _MEIPASS
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def inputkinopoiskid(choice):
    if choice == 1:
        filmsearch = []
        while True:
            search = input('Введите название фильма и год выпуска или enter чтобы продолжить: ')
            if search == '':
                return filmsearch
            movie_list = Movie.objects.search(search)   
            id = str(movie_list[0].id)
            print(movie_list[0])
            print(f"Kinopoisk_id: {id}")
            choice_1 = input('Варианты: Добавить в список (1), новый поиск (2), закончить и продолжить (enter): ')
            if choice_1 == '1':
                filmsearch.append(id)
                print(filmsearch)
            elif choice_1 == '2':
                continue
            elif choice_1 == '':
                print(filmsearch)
                return filmsearch
   
    elif choice == 2:
        inputstr = input('Введите через пробел идентификаторы фильмов (kinopoisk id): ')
        return  inputstr.split()


terminal_size = os.get_terminal_size().columns - 1
print('-' * (terminal_size))
print("Kinolist: Программа создания списка фильмов".center(terminal_size, " "))
print(ver.center(terminal_size, " "))
print('-' * (terminal_size))

if not isapiok(api):
    print('Ошибка API!')
    os.system('pause')
    sys.exit()

# считываем значения из файла list.txt
film_codes = []
if os.path.isfile('./list.txt'):
    file_list = open('./list.txt', 'r')
    lines = file_list.readlines()  # считываем все строки
    file_list.close()
    for line in lines:
        film_codes.append(line.strip())
    if len(film_codes) < 1:
        print('В списке 0 фильмов. Работа программы завершена.')
        os.system('pause')
        sys.exit()
    print('Найден файл "list.txt"' + ' (записей: ' + str(len(film_codes)) + ')')
else:
    print('Ошибка: Файл "list.txt" не найден!')
    while True:
        choice = input('Выберите режим: Поиск фильмов по названию (1); ручной ввод kinopoisk_id (2); enter чтобы выйти: ')
        if choice == "1":
            print('выбор 1')
            film_codes = inputkinopoiskid(1)
            print(film_codes)
            break
        elif choice == "2":
            print('выбор 2')
            film_codes = inputkinopoiskid(2)
            print(film_codes)
            break
        elif choice == "":
            print('')
            print('Работа программы завершена.')
            os.system('pause')
            sys.exit()

    if len(film_codes) < 1:
        print('В списке 0 фильмов. Работа программы завершена.')
        os.system('pause')
        sys.exit()
    else:
        with open('./list.txt', 'w') as f:
            f.write('\n'.join(film_codes))
        print('Файл "list.txt" сохранен.')

file_path = resource_path('template.docx')  # определяем путь до шаблона
try:
    doc = Document(file_path)  # открываем шаблон
except Exception:
    print('Ошибка! Не найден шаблон template.docx. Список не создан.')
    print('')
    print('Работа программы завершена.')
    os.system('pause')
    sys.exit()

if len(film_codes) > 1:
    cloneFirstTable(doc, len(film_codes) - 1)  # добавляем копии шаблонов таблиц

err = 0
tablenum = 0
fullfilmslist = []
for i in range(len(film_codes)):
    if i > 20:
        print('Ошибка! Достигнуто ограничение API - не больше 20 фильмов за раз.')
        break
    try:
        filminfo = getFilminfo(film_codes[i], api)
        fullfilmslist.append(filminfo)
    except:
        print(film_codes[i] + ' - ошибка')
        err += 1
    else:
        current_table = doc.tables[tablenum]
        writeFilmtoTable(current_table, filminfo)
        print(filminfo[0] + ' - ок')
        tablenum += 1

try:
    doc.save('./list.docx')
except PermissionError:
    print('Ошибка! Нет доступа к файлу list.docx. Список не создан.')
    print('')
    print('Работа программы завершена.')
    os.system('pause')
    sys.exit()

if err > 0:
    print('Выполнено с ошибками! (' + str(err) + ')')
else:
    print('')
    print('Список создан.')
    print('-' * terminal_size)

if len(film_codes) - 1 > i:
    print('Внимание! В файле списка присутствуют лишние пустые таблицы.')

mp4files = glob.glob('*.mp4')
if len(mp4files) < 1:
    print('Файлы mp4 не найдены.')
    print('')
    print('Работа программы завершена.')
    os.system('pause')
    sys.exit()

print('Найдены файлы mp4:')
for file in mp4files:
    print(file)

sys.exit()

print('')
ask = str(input('Начать запись тегов? (y/n) '))
if ask.lower() == "y":
    for film in fullfilmslist:
        writeTagstoMp4(film)
    print('')
    print('Запись тегов завершена.')
elif ask == '' or ask == 'n':
    print('Отмена. Работа программы завершена.')
os.system('pause')
sys.exit()

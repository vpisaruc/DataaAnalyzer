from bs4 import BeautifulSoup as bs
import urllib.request as ur
import datetime
from glob import glob
import re
import os
import win32com.client as win32
from win32com.client import constants

"""Класс парсера данных"""
class Pars:

    """Текущая дата"""
    now = datetime.datetime.now()

    """Функция запуска парсера"""
    def _start_parser(self):
        self.__parser_2018()
        self.__parser_2019()

    """Конвертер из .doc в .docx, т.к с .doc файлами невозможно работать"""
    def __save_as_docx(self, path):
        # Открыть MS Word
        word = win32.gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Open(path)
        doc.Activate()

        # Переименуем пусть на .docx
        new_file_abs = os.path.abspath(path)
        new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

        # Сохраняем и закрываем
        word.ActiveDocument.SaveAs(
            new_file_abs, FileFormat=constants.wdFormatXMLDocument
        )
        doc.Close(False)

        # Удаляем .doc файл
        os.remove(path)

    """Парсим данные 2018 года, всех месяцев по очереди"""
    def __parser_2018(self):
        paths = glob(r'D:\work\rosbank_test\documents\2018\*.doc', recursive=True)
        url = 'http://www.gks.ru/bgd/free/B18_00/?List&Id=-1'

        req = ur.urlopen(url)
        xml = bs(req, 'xml')

        cnt = 12

        for item in xml.find_all('ref'):
            req = ur.urlopen(url[:-2] + item.text[1:3])
            xml = bs(req, 'xml')

            for l in xml.find_all('ref'):
                 if l.text[-7:] == '6-0.doc' and cnt > 1:
                     ur.urlretrieve('http://www.gks.ru' + l.text[0:32] + l.text[65:len(l.text)],
                                    r'D:\work\rosbank_test\documents\2018\stat2018_' + str(cnt) + '.doc')

                     self.__save_as_docx(r'D:\work\rosbank_test\documents\2018\stat2018_' + str(cnt) + '.doc')

                     cnt -= 1

                 elif l.text[-7:] == '5-0.doc' and cnt == 1:
                     ur.urlretrieve('http://www.gks.ru' + l.text[0:32] + l.text[65:len(l.text)],
                                    r'D:\work\rosbank_test\documents\2018\stat2018_' + str(cnt) + '.doc')

                     self.__save_as_docx(r'D:\work\rosbank_test\documents\2018\stat2018_' + str(cnt) + '.doc')

                     cnt -= 1

    """Парсим данные 2018 года, всех месяцев по очереди"""
    def __parser_2019(self):
        paths = glob(r'D:\work\rosbank_test\documents\2019\*.doc', recursive=True)
        url = 'http://www.gks.ru/bgd/free/B19_00/?List&Id=-1'

        req = ur.urlopen(url)
        xml = bs(req, 'xml')

        cnt = 0

        for item in xml.find_all('ref'):
            if item.text == "?7":
                req = ur.urlopen(url[:-2] + item.text[1:3])
                xml = bs(req, 'xml')

                for l in xml.find_all('ref'):

                    if l.text[-7:] == '6-0.doc':
                        cnt += 1
                        ur.urlretrieve('http://www.gks.ru' + l.text[0:32] + l.text[65:len(l.text)],
                                       r'D:\work\rosbank_test\documents\2019\stat2019_' + str(cnt) + '.doc')
                        self.__save_as_docx(r'D:\work\rosbank_test\documents\2019\stat2019_' + str(cnt) + '.doc')

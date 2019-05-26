from my_excelexport import ExcelExporter
from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import QMessageBox
from my_graphdrawer import GraphDrawer
from my_parser import Pars
from my_dataminer_from_docx import DataMiningFromDocx
from my_database import *

"""Запускает функции по скачке, конвертации и загрузки данных в базу данных"""
class DataDownloader(Pars, DataMiningFromDocx):
    def _dataDownload(self):
        w = Window()

        self._start_parser()
        self._start_data_mining()

        """Данные для заполнения listWidget"""
        namesEconomicField = cursor.execute("""SELECT nameEconomicActivity FROM tbData
                                              WHERE curYear = '2019';""")
        namesEconomicField = namesEconomicField.fetchall()

        """Если виджет был заполнен чистим данные"""
        w.listWidgetEconimicField.clear()

        """Заполнение listWidget"""
        for i in range(len(namesEconomicField)):
            for j in range(len(namesEconomicField[i])):
                w.listWidgetEconimicField.addItem(namesEconomicField[i][j])

"""Инициализация интерфейса"""
class Window(QtWidgets.QMainWindow, GraphDrawer, DataDownloader, ExcelExporter):
    """Инициализация окна"""
    def __init__(self):

        QtWidgets.QWidget.__init__(self)
        uic.loadUi(r"interface\window.ui", self)
        # Заполняем виджет названиями
        self.__fillListWidgets()
        self.pushButton.clicked.connect(lambda: self._dataDownload())
        self.pushButton_2.clicked.connect(lambda: self.__draw())
        self.pushButton_3.clicked.connect(lambda: self._start())

    """Заполнение лсит виджетов периодами и названиями экономической деятельности"""
    def __fillListWidgets(self):

        """Данные для заполнения listWidget"""
        namesEconomicField = cursor.execute("""SELECT nameEconomicActivity FROM tbData
                                              WHERE curYear = '2019';""")
        namesEconomicField = namesEconomicField.fetchall()

        """Заполнение listWidget"""
        for i in range(len(namesEconomicField)):
            for j in range(len(namesEconomicField[i])):
                self.listWidgetEconimicField.addItem(namesEconomicField[i][j])

        """Заполнение comxoBox"""
        self.comboBox.addItem('Помесячная динамика з/п')
        self.comboBox.addItem('Сравнение с прошлым годом')

    def __draw(self):
        if self.listWidgetEconimicField.currentItem():
            """Получение номера строки и элемента строки"""
            if str(self.comboBox.currentText()) == 'Помесячная динамика з/п':
                rowListWidget = self.listWidgetEconimicField.currentRow()
                itemListWidget = self.listWidgetEconimicField.item(rowListWidget)
                """Отрисовка помесячного графика"""

                self._drawGraphMonth(itemListWidget.text())
            else:
                rowListWidget = self.listWidgetEconimicField.currentRow()
                itemListWidget = self.listWidgetEconimicField.item(rowListWidget)
                """Отрисовка графика о изменении заработной платы 
                   относительно этого же периода прошлого года"""
                self._drawGraphCompare(itemListWidget.text())

        else:
            QMessageBox.warning(self, "Ошибка", "Выберите отрасль из списка!")
import matplotlib.pyplot as plt
from my_database import *
import numpy as np
import re


"""Класс отвечает за отрисовку всех графиков"""
class GraphDrawer():

    """Представляет информацию в виде графика о динамике заработных плат по секторам экономики помесячно"""
    def _drawGraphMonth(self, economField):
        dataFromDb = cursor.execute("""SELECT sallaryRubles, curMonth, curYear
                                          FROM tbData
                                          WHERE nameEconomicActivity LIKE ?;""",
                                            ([economField]))

        dataFromDb = dataFromDb.fetchall()

        sallary_array = []
        monthyear_array = []

        """Формируем массив помесячных зарплат и массив месяц-год"""
        for i in range(len(dataFromDb)):
            sallary_array.append(dataFromDb[i][0])
            monthyear_array.append(dataFromDb[i][1] + ' ' + dataFromDb[i][2])

        x = np.arange(12)
        plt.figure(figsize=(15,15))
        plt.bar(x, height=sallary_array)
        plt.xticks(x, monthyear_array)
        plt.title('Динамика заработных плат в ' + economField)
        plt.xlabel('Период')
        plt.ylabel('Заработная плата в рублях')
        plt.grid(True)

        plt.show()

    """Представляет информацию в виде графика о изменении заработной платы 
    относительно этого же периода прошлого года"""
    def _drawGraphCompare(self, economField):
        """Массивы для хранения данных из бд"""
        sallary_array_this_year = []
        sallary_array_last_year = []
        monthyear_array = []
        sallary = 0

        """Зарплаты по выбранной отрасли"""
        dataFromDb = cursor.execute("""SELECT sallaryRubles, sallaryLastYearThisMonth, curMonth
                                          FROM tbData
                                          WHERE nameEconomicActivity = ?;""",
                                            ([economField]))
        dataFromDb = dataFromDb.fetchall()

        """Средний показатель зарплаты по экономике"""
        averageSallary = cursor.execute(""" SELECT sallaryRubles FROM tbData WHERE nameEconomicActivity = 'Всего';""")
        averageSallary = averageSallary.fetchall()

        """Формируем массив помесячных зарплат"""
        for i in range(len(dataFromDb)):
            sallary_array_this_year.append(dataFromDb[i][0])
            sallary_array_last_year.append((float(re.sub(',', '.', dataFromDb[i][1]))/100) * float(dataFromDb[i][0]))
            monthyear_array.append(dataFromDb[i][2])

        """Находим средний показатель зарплат по экономике"""
        for i in range(len(averageSallary)):
            sallary += int(averageSallary[i][0])

        sallary = sallary / 12
        sallary = int(sallary)

        """Количество периодов"""
        x = np.arange(12)

        """Настройки баров в гистограмме"""
        plt.figure(figsize=(15, 15))
        bar_width = 0.35
        opacity = 0.8



        this = plt.bar(x, sallary_array_this_year, bar_width, alpha=opacity, color = 'b', label='Этот год')
        last = plt.bar(x + bar_width, sallary_array_last_year, bar_width, alpha=opacity, color = 'g',
                       label='Прошлый год')

        plt.axhline(y=sallary, color='r', linestyle='-', label='Средний показатель по экономики')
        plt.xlabel('Месяц')
        plt.ylabel('Зарплата')
        plt.title('Изменение заработной платы относительно этого же периода прошлого года')
        plt.xticks(x + bar_width, monthyear_array)
        plt.legend()
        plt.tight_layout()
        plt.show()




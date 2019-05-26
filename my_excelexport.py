from my_database import *
import xlsxwriter

"""Экспортируем данные из базы данных в эксель файл"""
class ExcelExporter:
    def _start(self):
        self.__export('Январь', '2018', '-')
        self.__export('Февраль', '2018', 'Январь')
        self.__export('Март', '2018', 'Февраль')
        self.__export('Апрель', '2018', 'Март')
        self.__export('Май', '2018', 'Апрель')
        self.__export('Июнь', '2018', 'Май')
        self.__export('Июль', '2018', 'Июнь')
        self.__export('Август', '2018', 'Июль')
        self.__export('Сентябрь', '2018', 'Август')
        self.__export('Октябрь', '2018', 'Сентябрь')
        self.__export('Ноябрь', '2018', 'Октябрь')
        self.__export('Январь', '2019', '-')

    def __export(self, month, year, lastMonth):
        # Счетчик строк
        row = 0

        workbook = xlsxwriter.Workbook(r'excel\\' + year + r'\\_' + month + '_' + year + '.xlsx')
        worksheet = workbook.add_worksheet()

        if lastMonth == '-':
            dataFromDb = cursor.execute("""SELECT nameEconomicActivity, sallaryRubles, sallaryLastYearThisMonth,
                                              sallaryPercentLastMonth, sallaryPercentТationalLevel
                                              FROM tbData
                                              WHERE curMonth = ? and curYear = ?; """, ([month, year]))

            dataFromDb = dataFromDb.fetchall()


            expenses = ['Название отрасли', month + ' ' + year + ' зарплата в рублях',
                        month + ' ' + str(int(year) - 1) + ' зарплата в %', lastMonth + ' ' + year + ' зарплата в %',
                        'Общероссийскийуровень средне-месячной заработной платы в %']

            # Записываем шапку таблицы
            for i in range((len(expenses))):
                worksheet.write(row, i, expenses[i])
            row += 1

            # Записываем данные в таблицу
            for i in range(len(dataFromDb)):
                for j in range(len(dataFromDb[i])):
                    worksheet.write(row, j, dataFromDb[i][j])
                row += 1

        else:
            dataFromDb = cursor.execute("""SELECT nameEconomicActivity, sallaryRubles, sallaryLastYearThisMonth,
                                      sallaryPercentLastMonth, sallaryPercentТationalLevel, inRublesJanuaryThisMonth, 
                                      inPercentJanuaryThisMonthLastYear FROM tbData
                                      WHERE curMonth = ? and curYear = ?; """, ([month, year]))

            dataFromDb = dataFromDb.fetchall()

            expenses = ['Название отрасли', month + ' ' + year + ' зарплата в рублях',
                        month + ' ' + str(int(year) - 1) + ' зарплата в %', lastMonth + ' ' + year + ' зарплата в %',
                        'Общероссийскийуровень средне-месячной заработной платы в %',
                        'Январь-' + month + ' ' + year + ' зарплата в рублях',
                        'Январь-' + month + ' ' + str(int(year) -1) + ' зарплата в %']

            # Записываем шапку таблицы
            for i in range((len(expenses))):
                worksheet.write(row, i, expenses[i])
            row += 1

            # Записываем данные в таблицу
            for i in range(len(dataFromDb)):
                for j in range(len(dataFromDb[i])):
                    worksheet.write(row, j, dataFromDb[i][j])
                row += 1


        workbook.close()

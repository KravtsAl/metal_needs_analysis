import pandas as pd
import openpyxl
from openpyxl.styles import Font


class needs_analysis:

    """Анализ производственного плана, фильтрация и трансформация в итоговую таблицу

            Методы:

            make_list_of_mat - возвращает список из требуемых материалов и их диаметра/толщины.
            months_names - Формирует названия месяцев
            check_month - Метод для трансформации данных в числовой тип или 0( если значений нет)
            make_form - формирует строки итоговой неотфильтрованной таблицы
            make_table - Метод создает "шапку" таблицы, и разделяет на 4 другие таблицы( по складам)
            filter_chek - метод фильрует и сохраняет итоговую таблицу
            finish - Метод настраивает ширину столбцов, шрифт,  а также скрывает лишние месяца

            """
    def __init__(self):

        """Загружает таблицу с нормами расходов материала в формате "xlsx"
                """

        self.dfMat = pd.read_excel('Mat2.xlsx')     # загрузка норм. расхода материалов


        self.mat_for_combobox = ["Все"]     # Cписок, который передается в ComboBox и пополняется названиями материалов
                                                                 # в зависимости от выбранного склада


        self.my_mat = self.make_list_of_mat()        # Cписок из названий материалов + их диаметр/толщина


        self.dfDG = None

        self.first_month = None
        self.second_month = None
        self.third_month = None
        self.fourth_month = None
        self.five_month = None
        self.all_mont = None           # названия месяцев и список из них, в зависимости от плана



        self.stock_831 = None     # итоговые таблицы с шапкой для каждого склада
        self.stock_816 = None
        self.stock_830 = None
        self.stock_832 = None



        self.filter_materialik = None

        self.all_mater = None

        self.super_final_list = None



    def make_list_of_mat(self):

        """Метод формирует список материалов из загружаемой таблицы с нормами

            :return: список из требуемых материалов и их диаметра/толщины.

        """

        a = [i for i in self.dfMat['Unnamed: 3']]
        b = [i for i in self.dfMat['Unnamed: 4']]
        a = a[2:]
        b = b[2:]
        return list(map(lambda x, y: str(x) + " " + str(y), a, b))





    def months_names(self, path='datefromDG.xlsx'):

        """Формирует названия месяцев исходя из таблицы с планом, а также создает список из этих месяцев, который
        передается в ComboBox_2

                :param path: путь к производственному плану
        """

        self.dfDG = pd.read_excel(path)

        self.first_month = self.dfDG.iat[0, 2]
        self.second_month = self.dfDG.iat[0, 8]
        self.third_month = self.dfDG.iat[0, 11]
        self.fourth_month = self.dfDG.iat[0, 14]
        self.five_month = self.dfDG.iat[0, 17]
        self.all_mont = ["Все", self.first_month, self.second_month, self.third_month, self.fourth_month,
                         self.five_month]




    def check_month(self, df):
        """Метод для трансформации данных в числа или 0( если значений нет)
            param path: путь к производственному плану
        """

        try:
            return int(df)
        except:
            return 0




    def make_form(self):
        """Метод перемножает кол-во планируемых деталей на каждый месяц и норм. расхода материала, а также
        формирует строки итоговой неотфильтрованной таблицы

            :return: список со списками  - названия материала, его характеристиками и
                потребностью на каждый месяц

        """



        list_details_index = []
        for index, val in enumerate(self.dfDG['Наименование'].dropna()):
            list_details_index.append([index + 5, val])
        list_details_index = list_details_index[0:21]                       # список из деталей и индексов


        month = []                                  #кол-во запланированных деталей на каждый месяц
        for i in list_details_index:
            first = self.check_month(self.dfDG.loc[[i[0]]]['Unnamed: 3'])
            second = self.check_month(self.dfDG.loc[[i[0]]]['Unnamed: 8'])
            third = self.check_month(self.dfDG.loc[[i[0]]]['Unnamed: 11'])
            fought = self.check_month(self.dfDG.loc[[i[0]]]['Unnamed: 14'])
            fiven = self.check_month(self.dfDG.loc[[i[0]]]['Unnamed: 17'])
            month.append([first, second, third, fought, fiven])


        final_list = []                                            #формируем список, где кол-во каждой детали  на каждый месяц перемножаем
                                                                #на норму потребности материлов
        for index, val in enumerate(list_details_index):
            result_signal = []
            for index2, val2 in enumerate(self.dfMat[val[1]]):
                if type(val2) == float:
                    temp_month = []
                    for j in month[index]:
                        temp_month.append(val2 * j)
                    result_signal.append(
                        [val[1], int(self.dfMat['Unnamed: 28'][index2]), str(self.dfMat['Unnamed: 1'][index2]),
                         str(self.dfMat['Unnamed: 2'][index2]),
                         str(self.dfMat['Unnamed: 3'][index2]) + ' ' + str(
                             self.dfMat['Unnamed: 4'][index2])] + temp_month)
            final_list += result_signal




        final_dict = {}       #складываем данные по материалу

        for i in final_list:
            if i[4] not in final_dict:
                final_dict[i[4]] = [i[1], i[2], i[3], i[5], i[6], i[7], i[8], i[9]]
            else:
                final_dict[i[4]] = [i[1], i[2], i[3], final_dict[i[4]][3] + i[5], final_dict[i[4]][4] + i[6],
                                    final_dict[i[4]][5] + i[7], final_dict[i[4]][6] + i[8], final_dict[i[4]][7] + i[9]]

        super_final_list = []    #формируем строки итоговой таблицы
        for key in final_dict:
            super_final_list.append(
                [final_dict[key][0], final_dict[key][1], final_dict[key][2], key, final_dict[key][3],
                 final_dict[key][4], final_dict[key][5], final_dict[key][6], final_dict[key][7]])

        self.super_final_list = super_final_list




    def make_table(self):

        """Метод создает "шапку" таблицы, а также
        4 разные таблицы для каждого склада

        """

        df = pd.DataFrame(self.super_final_list,
                          columns=['Склад', "Номенклатурный номер", "Вид материала", "Материал", self.first_month,
                                   self.second_month, self.third_month, self.fourth_month, self.five_month])

        self.stock_831 = df[(df["Склад"] == 831)]
        self.stock_816 = df[(df["Склад"] == 816)]
        self.stock_830 = df[(df["Склад"] == 830)]
        self.stock_832 = df[(df["Склад"] == 832)]





    def filter_chek(self, path, list, val ="Все"):

        """Метод для филтра таблицы и сохранения по указанному пути

                    param path: путь, куда сохранять итоговую таблицу
                    param list: список из таблиц с выбранами складами, которые обьединяются в одну
                    param val: Название конкретного материала,по которому нужны данные.  по умолч. - Все.
                """
        df = pd.concat(list, axis=0)   #Объединяем таблицы в зависимости от складов

        if val != "Все":               #Если нужны данные по конкретному материалу
            df = df[(df["Материал"] == val)]
        df.to_excel(path, index=False)






    def finish(self, path="output22.xlsx", month = "Все"):
        """Метод настраивает ширину столбцов, шрифт,  а также скрывает лишние месяца

            param path: путь, где хранится итоговая таблица
            param month: Название конкретного месяца,по которому нужны данные. по умолч. - Все.

        """

        wb = openpyxl.load_workbook(path)

        sheet = wb.active

        if month == self.first_month:
            sheet.column_dimensions["F"].hidden = True
            sheet.column_dimensions["G"].hidden = True
            sheet.column_dimensions["H"].hidden = True
            sheet.column_dimensions["I"].hidden = True
        if month == self.second_month:
            sheet.column_dimensions["E"].hidden = True
            sheet.column_dimensions["G"].hidden = True
            sheet.column_dimensions["H"].hidden = True
            sheet.column_dimensions["I"].hidden = True
        if month == self.third_month:
            sheet.column_dimensions["E"].hidden = True
            sheet.column_dimensions["F"].hidden = True
            sheet.column_dimensions["H"].hidden = True
            sheet.column_dimensions["I"].hidden = True
        if month == self.fourth_month:
            sheet.column_dimensions["E"].hidden = True
            sheet.column_dimensions["F"].hidden = True
            sheet.column_dimensions["G"].hidden = True
            sheet.column_dimensions["I"].hidden = True
        if month == self.five_month:
            sheet.column_dimensions["E"].hidden = True
            sheet.column_dimensions["F"].hidden = True
            sheet.column_dimensions["G"].hidden = True
            sheet.column_dimensions["H"].hidden = True


        sheet.column_dimensions["B"].width = 55
        sheet.column_dimensions["C"].width = 30
        sheet.column_dimensions["D"].width = 55
        sheet.column_dimensions["E"].width = 20
        sheet.column_dimensions["F"].width = 20
        sheet.column_dimensions["G"].width = 20
        sheet.column_dimensions["H"].width = 20
        sheet.column_dimensions["I"].width = 20
        sheet.column_dimensions["J"].width = 20

        for i in range(1, 11):
            sheet.cell(1, i).font = Font(size=14, name="Times New Roman", bold=True)

        wb.save(path)




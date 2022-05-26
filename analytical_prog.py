from PyQt5.QtWidgets import (QApplication, QWidget, QMessageBox, QLabel,QDialog)
from PyQt5.QtCore import Qt
from PyQt5 import QtWidgets,QtCore
from PyQt5.QtWidgets import QLineEdit,QFileDialog, QMainWindow,QTextEdit, QAction, QApplication
import my_des
import metal_plan

class MyWindow(QtWidgets.QWidget):
    """Отвечает за графический интерфейс

                Методы:

                on_button,2,3,4 - добавляет или удаляет список материлов склада в comboBox
                Donwl_plan - Метод для загрузки плана и передачи данных в сomboBox_1 и сomboBox_2
                savefile - Метод для сохранение итоговой таблицы
                color_map - формирует список из готовых таблиц, которые будут соеденины в одну

                """
    def __init__(self):

        """Метод формирует окно, создает сигналы для кнопок
                    """

        QtWidgets.QWidget.__init__(self)
        self.ui = my_des.Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.download_button.clicked.connect(self.Donwl_plan)
        self.ui.save_button.clicked.connect(self.savefile)

        self.main = metal_plan.needs_analysis()
        self.ui.comboBox.addItems(self.main.mat_for_combobox)
        self.ui.checkBox.stateChanged.connect(self.on_button)
        self.ui.checkBox_2.stateChanged.connect(self.on_button_2)
        self.ui.checkBox_3.stateChanged.connect(self.on_button_3)
        self.ui.checkBox_4.stateChanged.connect(self.on_button_4)



    def on_button(self, state):

        """Метод добавляет или удаляет список материлов склада в comboBox в зависимости от статуса флага на кнопке
        :param state : статус флага

        """
        if state == Qt.Checked:
            self.ui.comboBox.addItems(self.main.my_mat[:27])
        else:
            for i in self.main.my_mat[:27]:
                mom = self.ui.comboBox.findText(i)
                self.ui.comboBox.removeItem(mom)

    def on_button_2(self, state):
        if state == Qt.Checked:
            self.ui.comboBox.addItems(self.main.my_mat[27:60])
        else:
            for i in self.main.my_mat[27:60]:
                mom = self.ui.comboBox.findText(i)
                self.ui.comboBox.removeItem(mom)

    def on_button_3(self, state):
        if state == Qt.Checked:
            self.ui.comboBox.addItems(self.main.my_mat[81:])
        else:
            for i in self.main.my_mat[81:]:
                mom = self.ui.comboBox.findText(i)
                self.ui.comboBox.removeItem(mom)

    def on_button_4(self, state):
        if state == Qt.Checked:
            self.ui.comboBox.addItems(self.main.my_mat[60:81])
        else:
            for i in self.main.my_mat[60:81]:
                mom = self.ui.comboBox.findText(i)
                self.ui.comboBox.removeItem(mom)




    def Donwl_plan(self):

        """Метод для загрузки плана и передачи данных в сomboBox_1 и сomboBox_2
             """


        fname = QtWidgets.QFileDialog.getOpenFileName(self, 'Выберите файл', '.',
        "Файлы Exсel (*.xlsx)")[0]

        if fname:
            self.ui.download_button.setText("План загружен")


            self.main.months_names(fname)
            self.main.make_form()
            self.ui.comboBox_2.addItems(self.main.all_mont)



    def savefile(self):
        
        """Метод для сохранение итоговой таблицы
                     """

        fnew = QFileDialog.getSaveFileName(self, 'Выберите файл', '.',
        "Файлы Exсel (*.xlsx)")[0]


        if fnew:
            self.main.make_table()
            self.main.filter_chek(fnew,self.filter_stocks(),self.ui.comboBox.currentText())
            self.main.finish(fnew, self.ui.comboBox_2.currentText())



    def filter_stocks(self):
        """Метод формирует список из готовых таблиц, которые будут соеденины в одну в зависим. от выбранного склада

        :return: список выбраных таблиц


               """

        a = []
        if self.ui.checkBox.isChecked()==True:
            a.append(self.main.stock_816)
        if self.ui.checkBox_2.isChecked()==True:
            a.append(self.main.stock_830)
        if self.ui.checkBox_3.isChecked()==True:
            a.append(self.main.stock_831)
        if self.ui.checkBox_4.isChecked()==True:
            a.append(self.main.stock_832)

        return a


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    window = MyWindow()
    window.show()
    sys.exit(app.exec_())

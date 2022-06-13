import re
import os
from datetime import datetime

from PyQt5 import uic
from PyQt5 import sip
from PyQt5 import QtCore, QtWidgets, QtGui
from PyQt5.QtWidgets import QApplication, QMainWindow, QDialog, QLabel, QLineEdit, QSpinBox
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import QIcon, QPixmap
import sys
from config import db
import xlsxwriter
import images
from excel_writer import ExcelWriter

class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        uic.loadUi('ui/AWE_main_menu.ui', self)
        self.CreateSpuWindow = None
        self.create_1Button.clicked.connect(self.create_spu_btn)
        self.create_2Button.clicked.connect(self.create_pvh_btn)
        self.create_3Button.clicked.connect(self.create_kmd_btn)
        self.clouse_btn.clicked.connect(self.exit_btn)
        db.delit_main_info_spu()
        db.delit_main_spu_table_info()

    def create_spu_btn(self):
        self.close()
        self.QuantitySpuWindow = QuantitySpuWindow()
        self.QuantitySpuWindow.show()

    def create_pvh_btn(self):
        self.close()
        self.QuantitySpuWindow = QuantitySpuWindow()
        self.QuantitySpuWindow.show()

    def create_kmd_btn(self):
        self.close()
        self.QuantitySpuWindow = QuantitySpuWindow()
        self.QuantitySpuWindow.show()

    def exit_btn(self):
        self.close()

# Диалоговое окно СПУ
class QuantitySpuWindow(QDialog):
    def __init__(self):
        super().__init__()
        uic.loadUi('ui/SPU_dialog.ui', self)
        self.nextButton.clicked.connect(self.next_btn_dialog)
        self.clouse_btn.clicked.connect(self.exit_btn)
        self.spinBox.setValue(1)

    def next_btn_dialog(self):
        number_of_progect = self.numb_prog_spinBox.value()
        number_of_prog_text =''
        if len(str(number_of_progect)) < 2:
            number_of_prog_text = f'00{number_of_progect}'
        elif len(str(number_of_progect)) < 3:
            number_of_prog_text = f'0{number_of_progect}'
        else:
            number_of_prog_text = f'{number_of_progect}'

        quantity_spu = self.spinBox.value()
        db.delite_spu()
        db.add_main_info_spu(number_of_prog_text, quantity_spu)
        self.CreateSpuWindow = CreateSpuWindow(quantity_spu)
        self.CreateSpuWindow.show()
        self.close()

    def exit_btn(self):
        self.MainWindow = MainWindow()
        self.MainWindow.show()
        self.close()

# Диалоговое окно СПУ
class CreateSpuWindow(QMainWindow):
    def __init__(self, value):
        super().__init__()
        uic.loadUi('ui/statement_spu_1.ui', self)
        db.delit_main_spu_table_info()
        self.value = []
        self.spu_list = []
        self.data_quant_spu = []
        self.data_quant_pp = []
        self.data_pp = []
        self.data_spu = []
        self.data_squer = []
        self.data_for_db_squer = []
        self.data_check_list = []
        self.data_luvers = []
        self.dict_quant_pp = {}
        self.dict_quant_spu = {}
        self.dict_for_db_spu_qnt = {}
        self.dict_for_db_spu_sq = {}
        self.dict_for_db_spu_luvers = {}
        for i in range(value):
            self.value.append(i)
            self.spu_list.append(f'CПУ №{i+1}')
            self.spu_label = QLabel(self)
            self.spu_label.setText(f"СПУ №{i+1}. Кол-во:")
            self.spu_label.wordWrap()
            self.quantity_spu = QSpinBox(self)
            self.quantity_spu.setValue(1)
            self.quantity_spu.setMinimum(1)
            self.squer_spu_label = QLabel(self)
            self.squer_spu_label.setText(f"S, м2:")
            self.squer_spu_doublespinbox = QDoubleSpinBox(self)
            self.squer_spu_doublespinbox.setValue(0.0)
            self.squer_spu_doublespinbox.setMinimum(0.0)
            self.squer_spu_doublespinbox.setMaximum(1000000.00)
            self.pp_label = QLabel(self)
            self.pp_label.setText(f"Кол-во видов ПП, шт:")
            self.pp_label.wordWrap()
            self.quantity_pp = QSpinBox(self)
            self.quantity_pp.setValue(1)
            self.quantity_pp.setMinimum(1)
            self.quantity_pp.setMaximum(1000)
            self.luverses = QCheckBox(self)
            self.luverses.setText(f'Есть люверсы?')
            self.gridLayout_2.addWidget(self.spu_label, 2 * i, 0)
            self.gridLayout_2.addWidget(self.quantity_spu, 2 * i, 1)
            self.gridLayout_2.addWidget(self.squer_spu_label, 2 * i, 2)
            self.gridLayout_2.addWidget(self.squer_spu_doublespinbox, 2 * i, 3)
            self.gridLayout_2.addWidget(self.luverses, 2 * i, 4)
            self.gridLayout_2.addWidget(self.pp_label, 2 * i, 5)
            self.gridLayout_2.addWidget(self.quantity_pp, 2 * i, 6)
            self.data_quant_spu.append(self.quantity_spu)
            self.data_squer.append(self.squer_spu_doublespinbox)
            self.data_quant_pp.append(self.quantity_pp)
            self.data_check_list.append(self.luverses)
        self.next_tu_spu_2Button.clicked.connect(self.next_wind_spu_btn)
        self.back_btn.clicked.connect(self.back_btn_cl)


    def next_wind_spu_btn(self):
        # Добавляем в лист количество ПП
        for q_pp in self.data_quant_pp:
            self.data_pp.append(q_pp.value())
        # Добавляем в словарь Номер СПУ и количетсво ПП
        self.dict_quant_pp = dict(zip(self.value, self.data_pp))
        db.delite_spu()
        # Добавляем в базу данных значения из словаря
        for spu in self.dict_quant_pp:
            db.add_spu(spu, self.dict_quant_pp[spu])
        # Добавляем в лист количество СПУ
        for q_spu in self.data_quant_spu:
            self.data_spu.append(q_spu.value())
        # Добавляем в лист площадь СПУ
        for sq_spu in self.data_squer:
            self.data_for_db_squer.append(sq_spu.value())
        # Cоединяем листы номер СПУ и кол-во СПУ в словарь
        self.dict_for_db_spu_sq = dict(zip(self.spu_list, self.data_for_db_squer))
        # Cоединяем листы номер СПУ и площадь СПУ в словарь
        self.dict_for_db_spu_qnt = dict(zip(self.spu_list, self.data_spu))
        # Добавляем значение в базу данных
        for s in self.dict_for_db_spu_qnt:
            db.add_evth_ab_spu(s, self.dict_for_db_spu_qnt[s])
        # Обновляем значение в базу данных
        for s_2 in self.dict_for_db_spu_sq:
            db.update_squer_for_spu(self.dict_for_db_spu_sq[s_2], s_2)

        self.dict_quant_spu = dict(zip(self.value, self.data_spu))
        for spu in self.dict_quant_spu:
            db.update_qant_spu(self.dict_quant_spu[spu], spu)

        # Достаем значение из чек бокса
        for checkbox in self.data_check_list:
            if checkbox.isChecked():
                print('ahaha')
                self.data_luvers.append('True')
            else:
                self.data_luvers.append('False')
        # Объеденяем название СПУ и люверсы в один словарь
        self.dict_for_db_spu_luvers = dict(zip(self.spu_list, self.data_luvers))
        for l in self.dict_for_db_spu_luvers:
            if self.dict_for_db_spu_luvers[l] == 'True':
                db.update_luvers_status(l)

        self.ReadyTableSpuWindow = ReadyTableSpuWindow()
        self.ReadyTableSpuWindow.show()
        self.close()

    def back_btn_cl(self):
        self.close()
        self.QuantitySpuWindow = QuantitySpuWindow()
        self.QuantitySpuWindow.show()


class ReadyTableSpuWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('ui/statement_spu_2.ui', self)
        self.value = db.get_spu()
        self.sputableWidget.setColumnWidth(0,60)
        self.sputableWidget.setColumnWidth(1, 115)
        self.sputableWidget.setColumnWidth(2, 120)
        self.sputableWidget.setColumnWidth(3, 195)
        self.sputableWidget.setColumnWidth(4, 150)
        self.load_data_for_table()
        self.add_str_btn.clicked.connect(self.add_btn_press)
        self.delite_str_btn.clicked.connect(self.delite_btn_press)
        self.back_btn.clicked.connect(self.back_btn_tabl_press)
        self.endButton.clicked.connect(self.press_end_btn)

    def load_data_for_table(self):
        db.delite_name_spu()
        list = db.get_spu()
        pp_name = ''
        # Формирует к каждому СПУ свое количетсво ПП, назначая индексы
        for n in list:
            # Поиск типа int в значении n
            b = int(re.search('\d+', str(n)).group(0))
            # db.get_pp получает количество пп для каждого СПУ
            for pp in db.get_pp(b):
                a = int(re.search('\d+', str(pp)).group(0))
                # Формирование записи и индексов для СПУ и ПП и их соединение между собой в базе данных
                for pp_numb in range(a):
                    spu_mark = f'CПУ - {b + 1}'
                    spu_quant_from_db = db.get_quant_spu(b)
                    spu_quant = int(re.search('\d+', str(spu_quant_from_db)).group(0))
                    spu_name = f'Утеплитель'
                    pp_name = f'ПП {b + 1}.{pp_numb + 1}'
                    db.add_name_spu(spu_mark, spu_quant, spu_name, pp_name)

        # Достаем СПУ и ПП и назначаем в таблицу в виджете
        data_about_spu = db.get_all_name_spu()
        self.sputableWidget.setRowCount(len(data_about_spu))
        tablerow = 0
        for row in db.get_all_name_spu():
            self.sputableWidget.setItem(tablerow, 0, QtWidgets.QTableWidgetItem(row[1]))
            self.sputableWidget.setItem(tablerow, 1, QtWidgets.QTableWidgetItem(str(row[2])))
            self.sputableWidget.setItem(tablerow, 2, QtWidgets.QTableWidgetItem(row[3]))
            self.sputableWidget.setItem(tablerow, 3, QtWidgets.QTableWidgetItem(row[4]))
            # Вставляем spinbox в таблицу
            self.spinbox_quantity_pp = QtWidgets.QSpinBox(self)
            self.spinbox_quantity_pp.setValue(1)
            self.spinbox_quantity_pp.setMinimum(1)
            self.spinbox_quantity_pp.setMaximum(1000)
            self.sputableWidget.setCellWidget(tablerow, 4, self.spinbox_quantity_pp)
            self.spinbox_squer_pp = QtWidgets.QDoubleSpinBox(self)
            self.spinbox_squer_pp.setValue(1.0)
            self.spinbox_squer_pp.setMinimum(1.0)
            self.spinbox_squer_pp.setMaximum(1000000.00)
            self.sputableWidget.setCellWidget(tablerow, 5, self.spinbox_squer_pp)
            tablerow += 1

    def keyPressEvent(self, event):
        super().keyPressEvent(event)
        # Копирование
        if event.key() == Qt.Key.Key_C and (event.modifiers() & Qt.KeyboardModifier.ControlModifier):
            copied_cells = sorted(self.sputableWidget.selectedIndexes())
            copy_text = ''
            max_column = copied_cells[-1].column()
            for c in copied_cells:
                copy_text += self.sputableWidget.item(c.row(), c.column()).text()
                if c.column() == max_column:
                    copy_text += '\n'
                else:
                    copy_text += '\t'
            QApplication.clipboard().setText(copy_text)

        # Удаление значения клавишей delite
        if event.key() == QtCore.Qt.Key_Delete:
            row = self.sputableWidget.currentRow()
            col = self.sputableWidget.currentColumn()
            self.sputableWidget.setItem(row, col, QtWidgets.QTableWidgetItem(''))
        else:
            super().keyPressEvent(event)

    def add_btn_press(self):
        try:
            rowPosition = self.sputableWidget.currentRow()
            self.sputableWidget.insertRow(rowPosition)
        except:
            rowPosition = self.sputableWidget.rowCount()
            self.sputableWidget.insertRow(rowPosition)

    # Кнопка удаляет строку в таблице 1
    def delite_btn_press(self):
        if self.sputableWidget.rowCount() > 0:
            row = self.sputableWidget.currentRow()
            self.sputableWidget.removeRow(row)


    def back_btn_tabl_press(self):
        self.close()
        self.QuantitySpuWindow = QuantitySpuWindow()
        self.QuantitySpuWindow.show()

    def press_end_btn(self):
        try:
            db.delite_new_final_record()
            rows_2 = self.sputableWidget.rowCount()
            cols_2 = self.sputableWidget.columnCount()
            for row in range(self.sputableWidget.rowCount()):
                marks = self.sputableWidget.item(row, 0).text()
                quant_mark = self.sputableWidget.item(row, 1).text()
                name_spu = self.sputableWidget.item(row, 2).text()
                name_pp = self.sputableWidget.item(row, 3).text()
                quant_pp = self.sputableWidget.cellWidget(row, 4).value()
                squer_pp = self.sputableWidget.cellWidget(row, 5).value()
                db.add_new_record_of_pp(marks, quant_mark, name_spu, name_pp, quant_pp, squer_pp)

            ExcelWriter(rows_2, cols_2)
        except:
            error = 'Ошибка cоздания отчета. Возможно файл уже открыт. Попробуйте снова.'
            self.MainWindow = ErrorAddReport(error)
            self.MainWindow.show()

class ErrorAddReport(QDialog):
    def __init__(self, data):
        super().__init__()
        uic.loadUi('ui/errors/error_dialog_report.ui', self)
        self.text_error = data
        self.label_dscr_of_error.clear()
        self.label_dscr_of_error.setText(self.text_error)
        self.ok_btn.clicked.connect(self.ok_btn_press)

    def focusOutEvent(self, event):
        self.activateWindow()
        self.raise_()
        self.show()

    def ok_btn_press(self):
        self.close()

def application():
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon('images/manager.png'))
    window = MainWindow()  # cоздаетcя для каждого окна
    window.show()  # Необходимо для того, чтобы окно показалось
    sys.exit(app.exec_())  # Необходим для корректного завершения работы программы


if __name__ == "__main__":
    application()
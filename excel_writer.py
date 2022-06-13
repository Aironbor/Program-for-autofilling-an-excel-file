import os
import re
from datetime import datetime

import xlsxwriter

from config import db


class ExcelWriter:
    def __init__(self, rows_2, cols_2):
        super().__init__()
        self.rows_2 = rows_2
        self.cols_2 = cols_2
        progect_numb = int(re.search('\d+', str(db.get_info_spu())).group(0))
        if len(str(progect_numb)) < 2:
            self.number_of_prog_tru = f'00{progect_numb}'
        elif len(str(progect_numb)) < 3:
            self.number_of_prog_tru = f'0{progect_numb}'
        else:
            self.number_of_prog_tru = f'{progect_numb}'
        self.write_report_about_spu()

    def write_report_about_spu(self):
        workbook = xlsxwriter.Workbook(f'Готовые производственные задания/{self.number_of_prog_tru} - СПУ.xlsx')
        # Форматы format()
        percent_format = workbook.add_format(
            {'border': 1, 'num_format': '0.0%', 'align': 'center', 'valign': 'vcenter'})
        percent_format_fin = workbook.add_format(
            {'border': 1, 'fg_color': '#C6E0B4', 'num_format': '0.0%', 'align': 'center', 'valign': 'vcenter'})
        squer_format = workbook.add_format({'border': 1, 'num_format': '#,#0', 'align': 'center', 'valign': 'vcenter'})
        squer_format_fin = workbook.add_format(
            {'border': 1, 'fg_color': '#C6E0B4', 'num_format': '0.00', 'align': 'center', 'valign': 'vcenter'})
        qauntity_format = workbook.add_format({'border': 1, 'num_format': '#0', 'align': 'center', 'valign': 'vcenter'})
        date_format = workbook.add_format({'num_format': 'mmmm d yyyy', 'align': 'center', 'valign': 'vcenter'})
        special_numb = workbook.add_format(
            {'num_format': '#0', 'bold': True, 'font_color': 'red', 'align': 'center', 'valign': 'vcenter'})
        special_numb_2 = workbook.add_format(
            {'num_format': '0.00', 'bold': True, 'font_color': 'red', 'align': 'center', 'valign': 'vcenter'})
        numb_2 = workbook.add_format(
            {'border': 1,'fg_color': '#C6E0B4', 'num_format': '0.00', 'align': 'center', 'valign': 'vcenter'})
        special_numb_proc = workbook.add_format(
            {'num_format': '0.0%', 'bold': True, 'font_color': 'red', 'align': 'center', 'valign': 'vcenter'})
        # Форматы заголовков
        main_table_names = workbook.add_format({'bold': True})
        right_align = workbook.add_format({'align': 'right'})
        name_format = workbook.add_format(
            {'border': 1, 'bold': True, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        name_in_tab_format = workbook.add_format(
            {'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        name_second_table_1_format = workbook.add_format(
            {'border': 1, 'fg_color': '#FFFF00', 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        name_second_table_2_format = workbook.add_format(
            {'border': 1, 'fg_color': '#C6E0B4', 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        # Форматы для объединнеых ячеек
        name_merge_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#0',
            'fg_color': '#BDD7EE'
        })
        merge_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'num_format': '#,#0'
        })
        # Установить свойства
        workbook.set_properties({
            'title': f'Производственное задание для утеплителя  П-{self.number_of_prog_tru}',
            'subject': 'With document properties',
            'author': 'Ivan Metliaev',
            'manager': '',
            'company': 'Тентовые конструкции',
            'category': 'Утеплитель',
            'keywords': 'СПУ, Ангары, Полотно',
            'created': datetime.today(),
            'comments': 'Created with Python and Ivan Metliaev program'})
        workbook.set_custom_property('Номер проекта', {self.number_of_prog_tru})

        # Создаваемые листы
        worksheet_1 = workbook.add_worksheet('Раскрой полипропилена')
        worksheet_2 = workbook.add_worksheet('Сшивка полипропилена')
        worksheet_3 = workbook.add_worksheet('Наклейка синтепона')
        worksheet_4 = workbook.add_worksheet('Пробивка люверс')
        worksheet_5 = workbook.add_worksheet('Упаковка')
        worksheet_6 = workbook.add_worksheet(f'Изготовление анагара {self.number_of_prog_tru}')
        current_row = 0
        spu_qaunt = 0
        # Лист 1
        worksheet_1.write('A1', 'Учет выпуска продукции', main_table_names)
        worksheet_1.write('A3', f'Проект {self.number_of_prog_tru}', main_table_names)
        worksheet_1.write('L1', f'Всего смен', right_align)
        worksheet_1.write('L2', f'Средняя производ.', right_align)
        row_name = ['№', 'Марка', 'Наименование', 'Наименование полуфабриката', 'Кол.', 'Площадь марки, м2',
                    'Площадь общая, м2', 'Доля марки в общей массе, %']
        second_table_row_name_1 = ['Требуется изготовить', '% готовности ангара']
        second_table_row_name_2 = ['Изготовлено', 'Площадь, м2']
        # Заголовки первой таблицы
        worksheet_1.write_row(3, 0, row_name, name_format)
        # Заголовки второй таблицы 1 часть желтая
        worksheet_1.write_row(3, 9, second_table_row_name_1, name_second_table_1_format)
        # Заголовки второй таблицы 2 часть зеленая
        worksheet_1.write_row(3, 11, second_table_row_name_2, name_second_table_2_format)

        quantity_rows = self.rows_2
        # Общее количество ПП
        worksheet_1.write('E3', f'=SUM(E5:E{4 + quantity_rows})', special_numb)
        # Общая площадь всех ПП
        worksheet_1.write('G3', f'=SUM(G5:G{4 + quantity_rows})', special_numb_2)
        # Процент объема работ
        worksheet_1.write('H3', f'=SUM(H5:H{4 + quantity_rows})', special_numb_proc)
        # Общее количество сколько нужно изготовить
        worksheet_1.write('J3', f'=SUM(J5:J{4 + quantity_rows})', special_numb_2)
        # Общий процент готовности
        worksheet_1.write('K3', f'=SUM(K5:K{4 + quantity_rows})', special_numb_proc)
        # Изготовлено
        worksheet_1.write('L3', f'=SUM(L5:L{4 + quantity_rows})', special_numb)
        # Общая масса изготовленного
        worksheet_1.write('M3', f'=SUM(M5:M{4 + quantity_rows})', special_numb_2)
        # Средняя производительность
        worksheet_1.write_formula('M2', f'=IFERROR(M3/M1,"0")')
        worksheet_1.write_formula('M1',
                                  f'=СЧЁТЗ(P1,S1,V1,Y1,AB1,AE1,AH1,AK1,AN1,AQ1,AT1,AW1,AZ1,BC1,BF1,BI1,BL1,BO1,BR1,BU1,BX1,EK1,CA1,CD1,CG1)')

        # ЛИСТ 2
        worksheet_2.write('A1', 'Учет выпуска продукции', main_table_names)
        worksheet_2.write('A3', f'Проект {self.number_of_prog_tru}', main_table_names)
        worksheet_2.write('L1', f'Всего смен', right_align)
        worksheet_2.write('L2', f'Средняя производ.', right_align)
        row_name = ['№', 'Марка', 'Наименование', 'Наименование полуфабриката', 'Кол.', 'Площадь марки, м2',
                    'Площадь общая, м2', 'Доля марки в общей массе, %']
        # Заголовки первой таблицы
        worksheet_2.write_row(3, 0, row_name, name_format)
        # Заголовки второй таблицы 1 часть желтая
        worksheet_2.write_row(3, 9, second_table_row_name_1, name_second_table_1_format)
        # Заголовки второй таблицы 2 часть зеленая
        worksheet_2.write_row(3, 11, second_table_row_name_2, name_second_table_2_format)
        # Общее количество ПП
        worksheet_2.write('E3', f'=SUM(E5:E{4 + quantity_rows})', special_numb)
        # Общая площадь всех ПП
        worksheet_2.write('G3', f'=SUM(G5:G{4 + quantity_rows})', special_numb_2)
        # Процент объема работ
        worksheet_2.write('H3', f'=SUM(H5:H{4 + quantity_rows})', special_numb_proc)
        # Общее количество сколько нужно изготовить
        worksheet_2.write('J3', f'=SUM(J5:J{4 + quantity_rows})', special_numb_2)
        # Общий процент готовности
        worksheet_2.write('K3', f'=SUM(K5:K{4 + quantity_rows})', special_numb_proc)
        # Изготовлено
        worksheet_2.write('L3', f'=SUM(L5:L{4 + quantity_rows})', special_numb)
        # Общая масса изготовленного
        worksheet_2.write('M3', f'=SUM(M5:M{4 + quantity_rows})', special_numb_2)
        # Средняя производительность
        worksheet_2.write_formula('M2', f'=IFERROR(M3/M1,"0")')
        worksheet_2.write_formula('M1',
                                  f'=СЧЁТЗ(P1,S1,V1,Y1,AB1,AE1,AH1,AK1,AN1,AQ1,AT1,AW1,AZ1,BC1,BF1,BI1,BL1,BO1,BR1,BU1,BX1,EK1,CA1,CD1,CG1)')
        # Заполнение смен
        num_list = [14]
        num_l = 14
        while num_l <= 83:
            num_l += 3
            num_list.append(num_l)
            print(num_l)
        print(num_list)
        for i in num_list:
            worksheet_1.write(0, i, f'Смена:', right_align)
            worksheet_1.write(2, i, f'Дата:', right_align)
            worksheet_1.merge_range(3, i, 3, i + 1, f'1', name_merge_format)

            worksheet_2.write(0, i, f'Смена:', right_align)
            worksheet_2.write(2, i, f'Дата:', right_align)
            worksheet_2.merge_range(3, i, 3, i + 1, f'1', name_merge_format)

        for row in range(self.rows_2):
            for col in range(self.cols_2):
                current_row = 4 + int(row)
                special_row_mum = 5 + int(row)
                # Нумерация строк таблицы
                worksheet_1.write_formula(current_row, 0, f'{row + 1}', qauntity_format)
                worksheet_2.write_formula(current_row, 0, f'{row + 1}', qauntity_format)
                # Общая площадь для каждого ПП
                worksheet_1.write_formula(current_row, 6, f'=E{5 + row}*F{5 + row}', squer_format)
                worksheet_2.write_formula(current_row, 6, f'=E{5 + row}*F{5 + row}', squer_format)
                # Доля массы в процентах
                worksheet_1.write_formula(current_row, 7, f'=G{5 + row}/$G$3', percent_format)
                worksheet_2.write_formula(current_row, 7, f'=G{5 + row}/$G$3', percent_format)
                # Требуется изготовить
                worksheet_1.write_formula(current_row, 9, f'=E{5 + row}-L{5 + row}', qauntity_format)
                worksheet_2.write_formula(current_row, 9, f'=E{5 + row}-L{5 + row}', qauntity_format)
                # %Готовности ангара
                worksheet_1.write_formula(current_row, 10, f'=M{5 + row}/$G$3', percent_format)
                worksheet_2.write_formula(current_row, 10, f'=M{5 + row}/$G$3', percent_format)
                # Изготовленно
                worksheet_1.write_formula(current_row, 11,
                                          f'=SUM(O{special_row_mum},R{special_row_mum},U{special_row_mum},X{special_row_mum},'
                                          f'AA{special_row_mum},AD{special_row_mum},AG{special_row_mum},AJ{special_row_mum},'
                                          f'AM{special_row_mum},AP{special_row_mum},AS{special_row_mum},AV{special_row_mum},'
                                          f'AY{special_row_mum},BB{special_row_mum},BE{special_row_mum},BH{special_row_mum},'
                                          f'BK{special_row_mum},BN{special_row_mum},BQ{special_row_mum},BT{special_row_mum},BW{special_row_mum},'
                                          f'BZ{special_row_mum},CC{special_row_mum},CF{special_row_mum})', numb_2)
                worksheet_2.write_formula(current_row, 11,
                                          f'=SUM(O{special_row_mum},R{special_row_mum},U{special_row_mum},X{special_row_mum},'
                                          f'AA{special_row_mum},AD{special_row_mum},AG{special_row_mum},AJ{special_row_mum},'
                                          f'AM{special_row_mum},AP{special_row_mum},AS{special_row_mum},AV{special_row_mum},'
                                          f'AY{special_row_mum},BB{special_row_mum},BE{special_row_mum},BH{special_row_mum},'
                                          f'BK{special_row_mum},BN{special_row_mum},BQ{special_row_mum},BT{special_row_mum},BW{special_row_mum},'
                                          f'BZ{special_row_mum},CC{special_row_mum},CF{special_row_mum})', numb_2)
                # Масса
                worksheet_1.write_formula(current_row, 12, f'=F{5 + row}*L{5 + row}', squer_format_fin)
                worksheet_2.write_formula(current_row, 12, f'=F{5 + row}*L{5 + row}', squer_format_fin)
                # Cмены
                for i in num_list:
                    worksheet_1.merge_range(current_row, i, current_row, i + 1, f'', merge_format)
                    worksheet_2.merge_range(current_row, i, current_row, i + 1, f'', merge_format)

        current_row = 4
        for record in db.get_final_record():
            # Колонка Марки
            worksheet_1.write(current_row, 1, record[1], name_in_tab_format)
            worksheet_2.write(current_row, 1, record[1], name_in_tab_format)
            # Количество марки
            spu_qaunt = record[2]
            # Колонка Наименование
            worksheet_1.write(current_row, 2, record[3], name_in_tab_format)
            worksheet_2.write(current_row, 2, record[3], name_in_tab_format)
            # Наименование полуфабриката
            worksheet_1.write(current_row, 3, record[4], name_in_tab_format)
            worksheet_2.write(current_row, 3, record[4], name_in_tab_format)
            # Количество
            worksheet_1.write(current_row, 4, record[5] * spu_qaunt, qauntity_format)
            worksheet_2.write(current_row, 4, record[5] * spu_qaunt, qauntity_format)
            # Площадь марки
            worksheet_1.write(current_row, 5, record[6], squer_format)
            worksheet_2.write(current_row, 5, record[6], squer_format)
            current_row += 1
        # Лист 3
        worksheet_3.write('A1', 'Учет выпуска продукции', main_table_names)
        worksheet_3.write('A3', f'Проект {self.number_of_prog_tru}', main_table_names)
        worksheet_3.write('K1', f'Всего смен', right_align)
        worksheet_3.write('K2', f'Средняя производ.', right_align)
        row_name_3 = ['№', 'Марка', 'Наименование', 'Кол.', 'Площадь марки, м2',
                      'Площадь общая, м2', 'Доля марки в общей массе, %']

        worksheet_4.write('A1', 'Учет выпуска продукции', main_table_names)
        worksheet_4.write('A3', f'Проект {self.number_of_prog_tru}', main_table_names)
        worksheet_4.write('K1', f'Всего смен', right_align)
        worksheet_4.write('K2', f'Средняя производ.', right_align)

        worksheet_5.write('A1', 'Учет выпуска продукции', main_table_names)
        worksheet_5.write('A3', f'Проект {self.number_of_prog_tru}', main_table_names)
        worksheet_5.write('K1', f'Всего смен', right_align)
        worksheet_5.write('K2', f'Средняя производ.', right_align)
        # Заголовки первой таблицы
        worksheet_3.write_row(3, 0, row_name_3, name_format)
        worksheet_4.write_row(3, 0, row_name_3, name_format)
        worksheet_5.write_row(3, 0, row_name_3, name_format)
        # Заголовки второй таблицы 1 часть желтая
        worksheet_3.write_row(3, 8, second_table_row_name_1, name_second_table_1_format)
        worksheet_4.write_row(3, 8, second_table_row_name_1, name_second_table_1_format)
        worksheet_5.write_row(3, 8, second_table_row_name_1, name_second_table_1_format)
        # Заголовки второй таблицы 2 часть зеленая
        worksheet_3.write_row(3, 10, second_table_row_name_2, name_second_table_2_format)
        worksheet_4.write_row(3, 10, second_table_row_name_2, name_second_table_2_format)
        worksheet_5.write_row(3, 10, second_table_row_name_2, name_second_table_2_format)
        db_row = int(re.search('\d+', str(db.get_info_about_qaut_spu())).group(0))

        # Общее количество ПП
        worksheet_3.write('D3', f'=SUM(D5:D{4 + db_row})', special_numb)
        worksheet_4.write('D3', f'=SUM(D5:D{4 + db_row})', special_numb)
        worksheet_5.write('D3', f'=SUM(D5:D{4 + db_row})', special_numb)
        # Общая площадь всех ПП
        worksheet_3.write('F3', f'=SUM(F5:F{4 + db_row})', special_numb_2)
        worksheet_4.write('F3', f'=SUM(F5:F{4 + db_row})', special_numb_2)
        worksheet_5.write('F3', f'=SUM(F5:F{4 + db_row})', special_numb_2)
        # Процент объема работ
        worksheet_3.write('G3', f'=SUM(G5:G{4 + db_row})', special_numb_proc)
        worksheet_4.write('G3', f'=SUM(G5:G{4 + db_row})', special_numb_proc)
        worksheet_5.write('G3', f'=SUM(G5:G{4 + db_row})', special_numb_proc)
        # Общее количество сколько нужно изготовить
        worksheet_3.write('I3', f'=SUM(I5:I{4 + db_row})', special_numb_2)
        worksheet_4.write('I3', f'=SUM(I5:I{4 + db_row})', special_numb_2)
        worksheet_5.write('I3', f'=SUM(I5:I{4 + db_row})', special_numb_2)
        # Общий процент готовности
        worksheet_3.write('J3', f'=SUM(J5:J{4 + db_row})', special_numb_proc)
        worksheet_4.write('J3', f'=SUM(J5:J{4 + db_row})', special_numb_proc)
        worksheet_5.write('J3', f'=SUM(J5:J{4 + db_row})', special_numb_proc)
        # Изготовлено
        worksheet_3.write('K3', f'=SUM(K5:K{4 + db_row})', special_numb)
        worksheet_4.write('K3', f'=SUM(K5:K{4 + db_row})', special_numb)
        worksheet_5.write('K3', f'=SUM(K5:K{4 + db_row})', special_numb)
        # Общая масса изготовленного
        worksheet_3.write('L3', f'=SUM(L5:L{4 + db_row})', special_numb_2)
        worksheet_4.write('L3', f'=SUM(L5:L{4 + db_row})', special_numb_2)
        worksheet_5.write('L3', f'=SUM(L5:L{4 + db_row})', special_numb_2)
        # Средняя производительность
        worksheet_3.write_formula('L2', f'=IFERROR(L3/L1,"0")')
        worksheet_3.write_formula('L1',
                                  f'=СЧЁТЗ(O1,R1,U1,X1,AA1,AD1,AG1,AJ1,AM1,AP1,AS1,AV1,AY1,BB1,BE1,BH1,BK1,BN1,BQ1,BT1,BW1,EK1,BZ1,CC1,CF1)')
        worksheet_4.write_formula('L2', f'=IFERROR(L3/L1,"0")')
        worksheet_4.write_formula('L1',
                                  f'=СЧЁТЗ(O1,R1,U1,X1,AA1,AD1,AG1,AJ1,AM1,AP1,AS1,AV1,AY1,BB1,BE1,BH1,BK1,BN1,BQ1,BT1,BW1,EK1,BZ1,CC1,CF1)')
        worksheet_5.write_formula('L2', f'=IFERROR(L3/L1,"0")')
        worksheet_5.write_formula('L1',
                                  f'=СЧЁТЗ(O1,R1,U1,X1,AA1,AD1,AG1,AJ1,AM1,AP1,AS1,AV1,AY1,BB1,BE1,BH1,BK1,BN1,BQ1,BT1,BW1,EK1,BZ1,CC1,CF1)')
        # Заполнение смен
        num_list_3 = [13]
        num_l = 13
        while num_l <= 82:
            num_l += 3
            num_list_3.append(num_l)
            print(num_l)
        for i in num_list_3:
            worksheet_3.write(0, i, f'Смена:', right_align)
            worksheet_3.write(2, i, f'Дата:', right_align)
            worksheet_3.merge_range(3, i, 3, i + 1, f'1', name_merge_format)

            worksheet_4.write(0, i, f'Смена:', right_align)
            worksheet_4.write(2, i, f'Дата:', right_align)
            worksheet_4.merge_range(3, i, 3, i + 1, f'1', name_merge_format)

            worksheet_5.write(0, i, f'Смена:', right_align)
            worksheet_5.write(2, i, f'Дата:', right_align)
            worksheet_5.merge_range(3, i, 3, i + 1, f'1', name_merge_format)

        for new_row in range(db_row):
            for new_col in range(3):
                current_row = 4 + int(new_row)
                special_row = 5 + int(new_row)
                # Нумерация строк таблицы
                worksheet_3.write_formula(current_row, 0, f'{new_row + 1}', qauntity_format)
                worksheet_5.write_formula(current_row, 0, f'{new_row + 1}', qauntity_format)
                # Общая площадь для каждого ПП
                worksheet_3.write_formula(current_row, 5, f'=D{5 + new_row}*E{5 + new_row}', squer_format)
                worksheet_5.write_formula(current_row, 5, f'=D{5 + new_row}*E{5 + new_row}', squer_format)
                # Доля массы в процентах
                worksheet_3.write_formula(current_row, 6, f'=F{5 + new_row}/$F$3', percent_format)
                worksheet_5.write_formula(current_row, 6, f'=F{5 + new_row}/$F$3', percent_format)
                # Требуется изготовить
                worksheet_3.write_formula(current_row, 8, f'=D{5 + new_row}-K{5 + new_row}', qauntity_format)
                worksheet_5.write_formula(current_row, 8, f'=D{5 + new_row}-K{5 + new_row}', qauntity_format)
                # %Готовности ангара
                worksheet_3.write_formula(current_row, 9, f'=L{5 + new_row}/$F$3', percent_format)
                worksheet_5.write_formula(current_row, 9, f'=L{5 + new_row}/$F$3', percent_format)
                # Изготовленно
                worksheet_3.write_formula(current_row, 10, f'=SUM(N{special_row},Q{special_row},T{special_row},'
                                                           f'W{special_row},Z5{special_row},AC{special_row},AF{special_row},'
                                                           f'AI{special_row},AL{special_row},AO{special_row},AR{special_row},'
                                                           f'AU{special_row},AX{special_row},BA{special_row},BD{special_row},'
                                                           f'BG{special_row},BJ{special_row},BM{special_row},BP{special_row},'
                                                           f'BS{special_row},BV{special_row},BY{special_row},CB{special_row},'
                                                           f'CE{special_row})', squer_format_fin)

                worksheet_5.write_formula(current_row, 10, f'=SUM(N{special_row},Q{special_row},T{special_row},'
                                                           f'W{special_row},Z5{special_row},AC{special_row},AF{special_row},'
                                                           f'AI{special_row},AL{special_row},AO{special_row},AR{special_row},'
                                                           f'AU{special_row},AX{special_row},BA{special_row},BD{special_row},'
                                                           f'BG{special_row},BJ{special_row},BM{special_row},BP{special_row},'
                                                           f'BS{special_row},BV{special_row},BY{special_row},CB{special_row},'
                                                           f'CE{special_row})', squer_format_fin)
                # Масса
                worksheet_3.write_formula(current_row, 11, f'=E{5 + new_row}*K{5 + new_row}', squer_format_fin)
                worksheet_5.write_formula(current_row, 11, f'=E{5 + new_row}*K{5 + new_row}', squer_format_fin)
                # Cмены
                for i in num_list_3:
                    worksheet_3.merge_range(current_row, i, current_row, i + 1, f'', merge_format)
                    worksheet_5.merge_range(current_row, i, current_row, i + 1, f'', merge_format)

        # Колонка Марки
        q = len(db.get_all_ab_spu())
        curnt_numb_row = 4
        for info in db.get_all_ab_spu():
            worksheet_3.write(curnt_numb_row, 1, info[1], name_in_tab_format)
            worksheet_3.write(curnt_numb_row, 2, 'Утеплитель', name_in_tab_format)

            worksheet_5.write(curnt_numb_row, 1, info[1], name_in_tab_format)
            worksheet_5.write(curnt_numb_row, 2, 'Утеплитель', name_in_tab_format)
            # Количество марки
            worksheet_3.write(curnt_numb_row, 3, info[2], qauntity_format)
            worksheet_5.write(curnt_numb_row, 3, info[2], qauntity_format)
            # Площадь марки
            worksheet_3.write(curnt_numb_row, 4, info[3], squer_format_fin)
            worksheet_5.write(curnt_numb_row, 4, info[3], squer_format_fin)
            curnt_numb_row += 1

        curnt_numb_row = 4
        num_of_row = 1
        special_row = 5
        num_for_formuls = 0
        for spu_with_luvers in db.get_spu_with_luvers():
            worksheet_4.write(curnt_numb_row, 0, num_of_row, qauntity_format)
            # Колонка Марки
            worksheet_4.write(curnt_numb_row, 1, spu_with_luvers[1], name_in_tab_format)
            worksheet_4.write(curnt_numb_row, 2, 'Утеплитель', name_in_tab_format)
            # Количество марки
            worksheet_4.write(curnt_numb_row, 3, spu_with_luvers[2], qauntity_format)
            # Площадь марки
            worksheet_4.write(curnt_numb_row, 4, spu_with_luvers[3], squer_format_fin)
            # Общая площадь для каждого ПП
            worksheet_4.write_formula(curnt_numb_row, 5, f'=D{5 + num_for_formuls}*E{5 + num_for_formuls}', squer_format)

            # Изготовленно
            worksheet_4.write_formula(curnt_numb_row, 10, f'=SUM(N{special_row},Q{special_row},T{special_row},'
                                                       f'W{special_row},Z5{special_row},AC{special_row},AF{special_row},'
                                                       f'AI{special_row},AL{special_row},AO{special_row},AR{special_row},'
                                                       f'AU{special_row},AX{special_row},BA{special_row},BD{special_row},'
                                                       f'BG{special_row},BJ{special_row},BM{special_row},BP{special_row},'
                                                       f'BS{special_row},BV{special_row},BY{special_row},CB{special_row},'
                                                       f'CE{special_row})', squer_format_fin)
            # Доля массы в процентах
            worksheet_4.write_formula(curnt_numb_row, 6, f'=F{5 + num_for_formuls}/$F$3', percent_format)
            # Требуется изготовить
            worksheet_4.write_formula(curnt_numb_row, 8, f'=D{5 + num_for_formuls}-K{5 + num_for_formuls}', qauntity_format)
            # %Готовности ангара
            worksheet_4.write_formula(curnt_numb_row, 9, f'=L{5 + num_for_formuls}/$F$3', percent_format)
            # Масса
            worksheet_4.write_formula(curnt_numb_row, 11, f'=E{5 + num_for_formuls}*K{5 + num_for_formuls}', squer_format_fin)
            # Cмены
            for i in num_list_3:
                worksheet_4.merge_range(curnt_numb_row, i, curnt_numb_row, i + 1, f'', merge_format)
            curnt_numb_row += 1
            num_of_row += 1
            special_row += 1
            num_for_formuls += 1



        workbook.close()
        os.startfile(f'Готовые производственные задания\{self.number_of_prog_tru} - СПУ.xlsx')
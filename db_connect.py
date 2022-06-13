import sqlite3

class AWE_DB():
    def __init__(self, date_base):
        """Подключаемся к Базе Данных и сохраняем курсор соединения"""
        self.connection = sqlite3.connect(date_base)
        self.cursor = self.connection.cursor()

    def add_main_info_spu(self, numb_prog, numb_spu):
        """Добавляем номер проекта и количество видов СПУ"""
        return self.cursor.execute("""INSERT INTO main_info_spu (number_of_project, number_of_spu) VALUES (?, ?)""",
                                   (numb_prog, numb_spu,)) and self.connection.commit()
    def delit_main_info_spu(self):
        """Удаляем информацию"""
        result = self.cursor.execute("""DELETE FROM main_info_spu""") and self.connection.commit()
        return result

    def get_info_spu (self):
        """Получаем количество видов СПУ"""
        self.cursor.execute("""SELECT number_of_project FROM main_info_spu""")
        return self.cursor.fetchall()

    def get_info_about_qaut_spu(self):
        """Получаем номер СПУ"""
        self.cursor.execute("""SELECT number_of_spu FROM main_info_spu""")
        return self.cursor.fetchall()

    def add_evth_ab_spu(self, spu_name, qautity_spu):
        """Добавляем номер СПУ и количество"""
        return self.cursor.execute("""INSERT INTO main_spu_table (spu_name, qautity_spu) VALUES (?, ?)""",
                                   (spu_name, qautity_spu,)) and self.connection.commit()

    def update_squer_for_spu(self, squer_spu, spu_name):
        """Обновляем площадь в таблице"""
        result = self.cursor.execute("""UPDATE main_spu_table SET squer_spu = ? WHERE spu_name = ?""", (squer_spu, spu_name,)) and self.connection.commit()
        return result

    def update_luvers_status(self, spu_name):
        """Обновляем площадь в таблице"""
        result = self.cursor.execute("""UPDATE main_spu_table SET luvers = TRUE WHERE spu_name = ?""",
                                     (spu_name,)) and self.connection.commit()
        return result

    def get_spu_with_luvers(self):
        """Получаем СПУ с люверсами"""
        self.cursor.execute("""SELECT * FROM main_spu_table WHERE luvers = TRUE""")
        return self.cursor.fetchall()

    def get_all_ab_spu(self):
        self.cursor.execute("""SELECT * FROM main_spu_table""")
        return self.cursor.fetchall()

    def delit_main_spu_table_info(self):
        """Удаляем информацию"""
        result = self.cursor.execute("""DELETE FROM main_spu_table""") and self.connection.commit()
        return result

    def add_spu(self, spu, pp):
        """Добавляем СПУ и количество PP"""
        return self.cursor.execute("""INSERT INTO spu_table (spu, quant_pp) VALUES (?, ?)""", (spu, pp,)) and self.connection.commit()

    def update_qant_spu(self, quant_spu, spu):
        """Обновляем количество СПУ в таблице"""
        result = self.cursor.execute("""UPDATE spu_table SET quant_spu = ? WHERE spu = ?""", (quant_spu, spu,)) and self.connection.commit()
        return result

    def update_pp(self, quant_pp, spu):
        """Обновляем количество ПП в таблице"""
        result = self.cursor.execute("""UPDATE spu_table SET quant_pp = ? WHERE spu = ?""", (quant_pp, spu,)) and self.connection.commit()
        return result


    def delite_spu(self):
        """Удаляем информацию"""
        result = self.cursor.execute("""DELETE FROM spu_table""") and self.connection.commit()
        return result

    def get_spu(self):
        """Получаем информацию об СПУ"""
        self.cursor.execute("""SELECT spu FROM spu_table""")
        return self.cursor.fetchall()

    def get_pp(self, spu):
        """Получаем количество ПП"""
        self.cursor.execute("""SELECT quant_pp FROM spu_table WHERE spu = ?""", (spu,))
        return self.cursor.fetchall()

    def get_quant_spu(self, spu):
        """Получаем количество каждого СПУ"""
        self.cursor.execute("""SELECT quant_spu FROM spu_table WHERE spu = ?""", (spu,))
        return self.cursor.fetchall()

    def add_name_spu(self, spu_mark:str, spu_quant, spu_name, pp_mark):
        """Добавляем СПУ и количество PP"""
        return self.cursor.execute("""INSERT INTO spu_save_data (spu_mark, spu_quant, spu_name, pp_mark) VALUES (?, ?, ?, ?)""",
                                   (spu_mark, spu_quant, spu_name, pp_mark,)) and self.connection.commit()

    def get_all_name_spu (self):
        """Получаем всю сохраненную информацию"""
        self.cursor.execute("""SELECT * FROM spu_save_data""")
        return self.cursor.fetchall()

    def delite_name_spu(self):
        """Удаляем информацию"""
        result = self.cursor.execute("""DELETE FROM spu_save_data""") and self.connection.commit()
        return result

    def add_new_record_of_pp(self, mark, quant_mark, name_spu, name_pp, quant_pp, squer_pp):
        """Добавляем всю информацию об полуфабрикатах в отдельную таблицу"""
        return self.cursor.execute(
            """INSERT INTO final_table (marks, quant_mark, name_spu, name_pp, quant_pp, squer_pp) VALUES (?, ?, ?, ?, ?, ?)""",
            (mark, quant_mark, name_spu, name_pp, quant_pp, squer_pp,)) and self.connection.commit()

    def get_final_record (self):
        """Получаем всю сохраненную информацию"""
        self.cursor.execute("""SELECT * FROM final_table""")
        return self.cursor.fetchall()

    def delite_new_final_record(self):
        """Удаляем информацию"""
        result = self.cursor.execute("""DELETE FROM final_table""") and self.connection.commit()
        return result
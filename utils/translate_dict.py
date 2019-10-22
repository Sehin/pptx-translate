import sqlite3


class TranslateDict:
    def __init__(self, path):
        self.connection = sqlite3.connect(path)
        self.connection.execute("""
                                CREATE TABLE if not exists translate_dict (id INTEGER PRIMARY KEY, text_to_translate text, translate text)
                                """)

    def create_base_table(self, csv_path):
        with open(csv_path, mode='r') as file:
            for line in file:
                line = line.strip()
                text_to_translate = line.split(";")[0]
                translate = line.split(";")[1]
                self.connection.execute(f"""
                                        INSERT INTO translate_dict(text_to_translate, translate) values ('{text_to_translate}', '{translate}')
                                        """)
            # for week in range(1, 59):
            #     self.connection.execute(f"""
            #                             INSERT INTO translate_dict(text_to_translate, translate) values ('Неделя {week}', 'Week {week}')
            #                             """)
            self.connection.commit()

    def get_dict(self):
        res_dict = {}
        res = self.connection.execute("SELECT * FROM translate_dict")
        res = res.fetchall()
        for line in res:
            res_dict.update({line[1]: line[2]})

        return res_dict

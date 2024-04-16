import sqlite3
import pandas as pd
import numpy as np

# Class for communicating with SQLite3 SQL database.
file_path = 'test.xlsx'
file_path2 = 'transactions.xlsx'
file_path_tree = 'finance-tree.xlsx'

db_name = 'AccountBook.db'
table_name = 'fox'


class SqlBox:
    def __init__(self):
        self.db_name = db_name
        self.database = sqlite3.connect(db_name)
        self.cursor = self.database.cursor()
        self.table_name = table_name
        self.headers = [("TRANSACTION_DATE", "DATE"), ("BRANCH", "STR"), ("CASHFLOW", "INT"), ("DESCRIPTION", "STR")]
        self.init_database()
        self.get_data()

    def init_database(self):
        # 1. Delete all tables in DB
        self.cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        for table in self.cursor.fetchall():
            self.cursor.execute(f"DROP TABLE {table[0]}")

        # 2. Create table
        create_query = f'''CREATE TABLE IF NOT EXISTS {table_name} (
            _Date DATE, 
            _Branch TEXT, 
            _Description TEXT,
            _CashFlow INT
        );'''
        self.cursor.execute(create_query)


    def get_data(self):
        accountBook = pd.read_excel(file_path2)
        for i in range(len(accountBook)):
            doc = accountBook.iloc[i]
            try:
                # 1. Get Datas
                new_data = {key: str(doc[key]) for key in ['_Date', '_Branch', '_Description']}
                new_data['_Branch'] = 'HOME/' + new_data['_Branch'].strip()
                str_in, str_out = doc['_IN'], doc['_OUT']

                if str(str_in) != 'nan':
                    new_data['_CashFlow'] = +int(str_in)
                elif str(str_out) != 'nan':
                    new_data['_CashFlow'] = -int(str_out)
                else:
                    continue

                # 2. Make Sql Query and Execute.
                sql_query = f"INSERT INTO {table_name}"
                sql_query += " ({})".format(",".join(new_data.keys()))
                value_box = []
                for value in new_data.values():
                    if type(value) == int:
                        value = f'{value}'
                    elif type(value) == str:
                        value = f'"{value}"'
                    value_box.append(value)

                sql_query += " VALUES({})".format(",".join(value_box))
                self.cursor.execute(sql_query)
                self.database.commit()
            except Exception as e:
                print(e)

        self.cursor.execute(f'SELECT * FROM {table_name}')

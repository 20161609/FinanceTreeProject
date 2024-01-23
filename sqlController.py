import sqlite3


# Class for communicating with SQLite3 SQL database.
class SqlBox:
    def __init__(self, db_name: str):
        self.db_name = db_name
        self.database = sqlite3.connect(db_name)
        self.cursor = self.database.cursor()
        self.headers = [("TRANSACTION_DATE", "DATE"), ("BRANCH", "STR"), ("CASHFLOW", "INT"), ("DESCRIPTION", "STR")]
        self.init_database()

    def init_database(self):
        create_query = '''CREATE TABLE IF NOT EXISTS ACCOUNTBOOK(
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            TRANSACTION_DATE DATE, 
            BRANCH STR, 
            CASHFLOW INT, 
            DESCRIPTION STR,
            CREATED_TIME TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )'''

        self.cursor.execute(create_query)


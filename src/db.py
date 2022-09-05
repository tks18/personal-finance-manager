import sqlite3 as sql


def connect(config):
    connection = sql.connect(config["DB"]["PATH"])
    connection.row_factory = sql.Row
    return connection


TABLES = [
    {
        "name": "INCOMES",
        "columns": [
            "_id INTEGER PRIMARY KEY",
            "date DATE",
            "master_id INTEGER NOT NULL",
            "is_taxable BOOLEAN NOT NULL CHECK (is_taxable IN (0, 1))",
            "amount REAL",
            "FOREIGN KEY (master_id) REFERENCES INCOME_MASTER (_id)",
        ],
    },
    {
        "name": "EXPENSES",
        "columns": [
            "_id INTEGER PRIMARY KEY",
            "date DATE",
            "master_id INTEGER NOT NULL",
            "amount REAL",
            "FOREIGN KEY (master_id) REFERENCES EXPENSE_MASTER(_id)",
        ],
    },
    {
        "name": "INVESTMENTS",
        "columns": [
            "_id INTEGER PRIMARY KEY",
            "date DATE",
            "master_id INTEGER NOT NULL",
            "units REAL",
            "cost REAL",
            "amount REAL",
            "FOREIGN KEY (master_id) REFERENCES INVESTMENT_MASTER(_id)",
        ],
    },
]


def create_tables(connection: sql.Connection):
    global TABLES

    for table in TABLES:
        sql_create_line = f"CREATE TABLE IF NOT EXISTS {table['name']}({','.join(table['columns'])}) WITHOUT ROWID"
        connection.execute(sql_create_line)


def close_connection(connection: sql.Connection):
    connection.close()


def create_connection(config):
    connection = connect(config)
    connection.execute("PRAGMA foreign_keys = 1")
    return connection

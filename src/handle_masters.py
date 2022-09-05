import sqlite3 as sql
import win32com.client as wincom
import pandas as pd

MASTER_TABLES = [
    {
        "name": "INCOME_MASTER",
        "columns": [
            "_id INTEGER PRIMARY KEY",
            "name TEXT",
            "category TEXT",
            "sub_type TEXT",
            "type TEXT DEFAULT 'Incomes' NOT NULL",
        ],
    },
    {
        "name": "EXPENSE_MASTER",
        "columns": [
            "_id INTEGER PRIMARY KEY",
            "name TEXT",
            "category TEXT",
            "sub_type TEXT",
            "type TEXT DEFAULT 'Expenses' NOT NULL",
        ],
    },
    {
        "name": "INVESTMENT_MASTER",
        "columns": [
            "_id INTEGER PRIMARY KEY",
            "code TEXT NOT NULL UNIQUE",
            "name TEXT",
            "category TEXT",
            "sub_type TEXT",
            "type TEXT DEFAULT 'Investments' NOT NULL",
        ],
    },
]

TABLE_CONFIG = [
    {
        "sheet_name": "Incomes",
        "table_name": "Incomes_Master",
        "type": "Income",
        "sql_name": "INCOME_MASTER",
    },
    {
        "sheet_name": "Expenses",
        "table_name": "Expenses_Master",
        "type": "Expense",
        "sql_name": "EXPENSE_MASTER",
    },
    {
        "sheet_name": "Investments",
        "table_name": "Investments_Master",
        "type": "Investment",
        "sql_name": "INVESTMENT_MASTER",
    },
]


def get_tables_from_excel(file_path, tables_config):
    tables = {}
    XL = wincom.gencache.EnsureDispatch("Excel.Application")
    XL.Visible = False
    workbook = XL.Workbooks.Open(file_path)
    for table in tables_config:
        sheet = workbook.Sheets[table["sheet_name"]]
        tableObject = sheet.ListObjects(table["table_name"])
        table_range = sheet.Range(tableObject.Range.Address)
        df = pd.DataFrame(table_range.Value)
        df_header = df.iloc[0]
        df = df[1:]
        df.columns = df_header
        df.reset_index(drop=True, inplace=True)
        df["_id"] = list(df.index.values)
        df.set_index("_id")
        tables[table["table_name"]] = df
    workbook.Close(False)
    XL.Quit()
    return tables


def create_tables(connection: sql.Connection):
    global MASTER_TABLES

    for table in MASTER_TABLES:
        sql_create_line = f"CREATE TABLE IF NOT EXISTS {table['name']}({','.join(table['columns'])}) WITHOUT ROWID"
        connection.execute(sql_create_line)


def load_masters_data(
    connection: sql.Connection,
    df: pd.DataFrame,
    is_investments: bool,
    table_name: str,
    type: str,
):
    data = []
    if is_investments:
        data = [
            (
                row["_id"],
                row["Code"],
                row["Name"],
                row["Category"],
                row["Type"],
                type,
            )
            for index, row in df.iterrows()
        ]
    else:
        data = [
            (row["_id"], row["Name"], row["Category"], row["Type"], type)
            for index, row in df.iterrows()
        ]
    create_tables(connection=connection)
    if is_investments:
        connection.executemany(f"INSERT INTO {table_name} VALUES(?,?,?,?,?,?)", data)
    else:
        connection.executemany(f"INSERT INTO {table_name} VALUES(?,?,?,?,?)", data)
    connection.commit()


def prepare_masters(config, connection):
    sql_tables = connection.execute("SELECT name from sqlite_master WHERE type='table'")
    tables_in_db = [dbtables["name"] for dbtables in sql_tables]
    tables_to_be_in_db = [table["sql_name"] for table in TABLE_CONFIG]
    if not all(item in tables_to_be_in_db for item in tables_in_db):
        tables = get_tables_from_excel(config["MASTERS"]["PATH"], TABLE_CONFIG)
        for table in TABLE_CONFIG:
            if table["sql_name"] not in tables_in_db:
                load_masters_data(
                    connection=connection,
                    df=tables[table["table_name"]],
                    is_investments=table["type"] == "Investment",
                    table_name=table["sql_name"],
                    type=table["type"],
                )

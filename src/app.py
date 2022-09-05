import plugins.config as config
import plugins.db as db
import helpers.masters as masters

config = config.load_or_create_config()
connection = db.create_connection(config=config)
masters.check_or_load_masters_data(config=config, connection=connection)
db.create_tables(connection=connection)
# connection.execute("INSERT INTO EXPENSES VALUES(?,?,?,?)", (1, "2022-02-02", 0, 100.0))
# connection.commit()
# super = pd.read_sql_query(
#     "SELECT a._id, a.date, b.name, b.category, a.amount FROM EXPENSES AS a INNER JOIN EXPENSE_MASTER as b ON a.master_id = b._id",
#     con=connection,
# )
# print(super["date"])

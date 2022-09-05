import config
import db
import handle_masters

config = config.load_or_create_config()
connection = db.create_connection(config=config)
handle_masters.prepare_masters(config=config, connection=connection)
db.create_tables(connection=connection)

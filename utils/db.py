import mysql.connector

db_config = {
    'host': 'localhost',
    'user': 'root',
    'password': 'Lpsmqlg.',
    'database': 'comersa'
}

def get_connection():
    return mysql.connector.connect(**db_config)

import psycopg2
import cryptocode
from configparser import ConfigParser as Config

def connect_to_db():
    key = "grupo.sysven779.."
    config = Config()
    config.read("c:/Sysven/config.ini")

    database = cryptocode.decrypt(config.get("APP","database"),key)
    server = cryptocode.decrypt(config.get("APP","server"),key)
    user = cryptocode.decrypt(config.get("APP","user"),key)
    password = cryptocode.decrypt(config.get("APP","password"),key)
    port = cryptocode.decrypt(config.get("APP","port"),key)
    try:     
        conn = psycopg2.connect(
            dbname=database,
            user=user,
            password=password,
            host=server,
            port=port
        )
        return conn
    except psycopg2.Error as e:
        print(f"Error al conectar a PostgreSQL: {e}")

        return None


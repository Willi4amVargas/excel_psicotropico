import psycopg2
import cryptocode
from configparser import ConfigParser as Config

def connect_to_db():
    key = "grupo.sysven779.."
    config = Config()
    config.read("c:/Sysven/config.ini")

    database = cryptocode.decrypt(config.get("LOCAL","database"),key)
    server = cryptocode.decrypt(config.get("LOCAL","server"),key)
    user = cryptocode.decrypt(config.get("LOCAL","user"),key)
    password = cryptocode.decrypt(config.get("LOCAL","password"),key)
    port = cryptocode.decrypt(config.get("LOCAL","port"),key)

    # database = 'cadm_syn_try'
    # server = 'localhost'
    # user = 'postgres'
    # password = 'admin'
    # port = 5432
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


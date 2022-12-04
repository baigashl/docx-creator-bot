from enum import unique
from peewee import *

db = PostgresqlDatabase(
    'caspi_db',
    host='localhost',
    port='5432',
    user='caspi_user',
    password='qwe123'
)
db.connect()


class BaseModel(Model):
    class Meta:
        database = db


class Order(BaseModel):
    orderCode = CharField(unique=True)

db.create_tables([Order])

db.close()
from sqlalchemy import Column, Integer, String, ForeignKey, Boolean, event
from sqlalchemy.orm import relationship, Session
from sqlalchemy.orm import declarative_base

Base = declarative_base()

# 定義 table: server ORM for Sqlite

#   host: 'e-smith.odoo.com'
#   db_name: 'odoo13-esmith-master-1011507'
#   url: 'https://e-smith.odoo.com/api/v1/boq_import_api/swagger.json?token=1a7119c2-1feb-401e-bf8d-c1fe5634582d&db=odoo13-esmith-master-1011507'

class Server(Base):
    __tablename__ = 'server'

    id = Column(Integer, primary_key=True)
    host = Column(String)
    db_name = Column(String)
    url = Column(String)
    token = Column(String)
    sync_yaml = Column(Boolean, default=False)
    active = Column(Boolean, default=True)

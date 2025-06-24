from sqlalchemy import Column, Integer, String, ForeignKey, Boolean, event
from sqlalchemy.orm import relationship, Session
from sqlalchemy.orm import declarative_base

Base = declarative_base()

# 定義 table: server ORM for Sqlite

class Server(Base):
    __tablename__ = 'server'

    id = Column(Integer, primary_key=True)
    host = Column(String)
    db_name = Column(String)
    url = Column(String)
    token = Column(String)
    sync_yaml = Column(Boolean, default=False)
    active = Column(Boolean, default=True)

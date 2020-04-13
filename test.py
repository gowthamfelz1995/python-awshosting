from sqlalchemy import create_engine,Column,Integer,String,Date,ForeignKey,MetaData,Table
from sqlalchemy.ext.declarative import declarative_base 
import pymysql.cursors
from _datetime import date



Base = declarative_base()
SQLALCHEMY_DATABASE_URI = 'mysql+pymysql://aspigrow:Aspi@2019@ec2-13-235-103-2.ap-south-1.compute.amazonaws.com:3306/docgen'
meta = MetaData()



class UserLog(Base):
    __tablename__ = "UserLog"

    id= Column('user_id',Integer, primary_key= True)
    username = Column('user_name',String(122), unique= True)
    organizationid = Column('organization_id',String(122), unique= True)
    folderid = Column('folder_id',String(122), unique= True)
    generateddate = Column('generated_date',Date,unique= False )
    filename = Column('file_name',String(122),unique= False)

DB_Engine = create_engine(SQLALCHEMY_DATABASE_URI)
user_log = Table(
   'UserLog', meta, 
   Column('user_id', Integer, primary_key = True), 
   Column('user_name', String), 
   Column('organization_id', String), 
    Column('folder_id', String),
     Column('generated_date', Date),
     Column('file_name', String),
)
connection = DB_Engine.connect() 
ins = user_log.insert().values(user_id = 1234, user_name = 'Gowtham',
organization_id='sfgdfgf0980',folder_id='dsgdfgdf123',file_name="parts.docx",generated_date=date.today())
connection.execute(ins)
Base.metadata.create_all(bind=DB_Engine)

try :
    print("connection succeed-->{}".format(connection))
except :
    print("connection failed-->{}".format(connection))
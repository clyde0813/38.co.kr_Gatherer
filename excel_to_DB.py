import pandas as pd
import pymysql
import os
import sqlalchemy

data = pd.read_excel(os.getcwd() + "\\비상장 30990개 기업개요.xlsx")

conn = sqlalchemy.create_engine('mysql://root:1350@192.168.0.48:3306/Damda_Portal?charset=utf8', echo=False)
data.to_sql('NoStock_nostock', conn, index=False, if_exists='append')

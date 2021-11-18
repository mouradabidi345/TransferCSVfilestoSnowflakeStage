import xlrd
import sqlalchemy
import pandas as pd
import xlrd
import sqlalchemy
from snowflake.connector.pandas_tools import write_pandas
import snowflake.connector

#Agentsbase Ignite file
df = dataframe = pd.read_excel(r"Z:\Shared\Associate Support\WFM/Base.xlsx", sheet_name = 0, engine='openpyxl')
df.to_csv (r"Z:\Shared\Public\BI_Test\Base.csv", 
                  index = None,
                  header=True)


# #Aherence Ignyte file
df2 = dataframe = pd.read_excel(r"Z:\Shared\Associate Support\WFM/ADH.xlsx", sheet_name = 2, engine='openpyxl')
df2.to_csv (r"Z:\Shared\Public\BI_Test\ADH.csv", 
                index = None,
                header=True)



#Qa scores Ignyte file:
df3 = dataframe = pd.read_excel(r"C:\Users\mabidi\ASEA, LLC\AS-Quality Assurance - Documents/QA Scores.xlsx", sheet_name = 2, engine='openpyxl')
df3.to_csv (r"Z:\Shared\Public\BI_Test\QAScores.csv", 
                index = None,
                header=True)



#Attendance Ignyte file;
df4 = dataframe = pd.read_excel(r"Z:\Shared\Associate Support\WFM/NewCalendar.xlsx", sheet_name = 1, engine='openpyxl')
df4.to_csv (r"Z:\Shared\Public\BI_Test\Attendance.csv", 
                index = None,
                header=True)


#ADP Hours Ignyte file:
df5 = dataframe = pd.read_excel(r"Z:\Shared\Associate Support\WFM/ADP Hours.xlsx", sheet_name = 1, engine='openpyxl')
df5.to_csv (r"Z:\Shared\Public\BI_Test\ADPHours.csv", 
                index = None,
                header=True)


#KbAttendance Ignyte file:
df6 = dataframe = pd.read_excel(r"Z:\Shared\Associate Support\WFM/New KB.xlsx", sheet_name = 0, engine='openpyxl')
df6.to_csv (r"Z:\Shared\Public\BI_Test\KbAttendance.csv", 
                index = None,
                header=True)


#InsightArticles:
df7 = dataframe = pd.read_excel(r"C:\Users\mabidi\ASEA, LLC\Associate Support Operations - Insight/Insight Newsfeed Content.xlsx", sheet_name = 0, engine='openpyxl')
df7.to_csv (r"Z:\Shared\Public\BI_Test\InsightArticles.csv", 
                index = None,
                header=True)



ctx = snowflake.connector.connect(
          user='SF_RAW_STAGE_SERVICE',
          password='Zg5XZ!mm%PvA',
          account='ba62849.east-us-2.azure',
          warehouse= 'COMPUTE_MACHINE',
          database='DB_ASEA_REPORTS',
          schema='PUBLIC')   


cur = ctx.cursor()

cur.execute("put file://Z:\Shared\Public\BI_Test\Base.csv @INTERNAL_WFM_STAGE OVERWRITE = True")
cur.execute("put file://Z:\Shared\Public\BI_Test\ADH.csv @INTERNAL_WFM_STAGE OVERWRITE = True")
cur.execute("put file://Z:\Shared\Public\BI_Test\QAScores.csv @INTERNAL_WFM_STAGE OVERWRITE = True")
cur.execute("put file://Z:\Shared\Public\BI_Test\Attendance.csv @INTERNAL_WFM_STAGE OVERWRITE = True")
cur.execute("put file://Z:\Shared\Public\BI_Test\ADPHours.csv @INTERNAL_WFM_STAGE OVERWRITE = True")
cur.execute("put file://Z:\Shared\Public\BI_Test\KbAttendance.csv @INTERNAL_WFM_STAGE OVERWRITE = True")
cur.execute("put file://Z:\Shared\Public\BI_Test\InsightArticles.csv @INTERNAL_WFM_STAGE OVERWRITE = True")
cur.close()



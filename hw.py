import xlrd
import mysql.connector
book =xlrd.open_workbook("Tawjihi-WestBank-2015.xls")
sheet =book.sheet_by_name("tawjihi")

mydb =mysql.connector.connect(host="localhost", user="root", passwd="", database="taw")

mycursor =mydb.cursor()

query = """insert into tawjihi (TAW_YEAR, STUDSUMM_STUD_AVG, STUDENT_NAME,SCHOOL_NAME,BRANCH_NAME) values (%s, %s, %s, %s, %s)"""

for r in range(1, sheet.nrows):
    TAW_YEAR = sheet.cell(r,0).value
    STUDSUMM_STUD_AVG = sheet.cell(r, 1).value
    STUDENT_NAME = sheet.cell(r, 2).value
    SCHOOL_NAME = sheet.cell(r, 3).value
    BRANCH_NAME = sheet.cell(r, 4).value
    val =(TAW_YEAR, STUDSUMM_STUD_AVG, STUDENT_NAME,SCHOOL_NAME,BRANCH_NAME)
    mycursor.executemany(query, (val,))


mydb.commit()
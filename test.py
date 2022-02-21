import mysql.connector
import xlrd

l=list()
loc=("C:\\Users\\hp\\Desktop\\task\\CATIAOP.xlsx")
a=xlrd.open_workbook(loc)
sheet=a.sheet_by_index(0)
sheet.cell_value(0,0)
for i in range(5,806):
    con=mysql.connector.connect(host="localhost",user="root",passwd="12345ss54321",database="shubhamdb1")
    cur=con.cursor()
    
    l.append(tuple(sheet.row_values(i)))

    q="insert into shubham111(Partnumber,Quantity,Type)values(%s,%s,%s)"
    cur.executemany(q,l)
    print(i)
    con.commit()
    con.close()
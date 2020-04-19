import sqlite3
conn = sqlite3.connect('myDatabase.db')



conn.execute('''CREATE TABLE IF NOT EXISTS company_table
         (id serial,
         name TEXT,
         age TEXT,
         department TEXT)''')
conn.commit()



tableStructure=conn.execute("SELECT sql FROM sqlite_master WHERE name = 'company_table';").fetchall()
print(tableStructure[0][0])



'------------------------------------'
name=['rahul','ankit','amit','saurabh','sahil']
age=[21,23,22,24,23]
department=['IT','Frontend','Backend','Marketing','UX']


for i in range(0,len(name)):
    conn.execute('Insert into company_table(name,age,department) values(?,?,?)',(name[i],age[i],department[i]))
    conn.commit()




'------------------------------------'
from xlrd import open_workbook
book = open_workbook("exceldata.xlsx")
sheet = book.sheet_by_index(0)

row=0
for i in range(1,sheet.nrows):
    name=sheet.cell(row, 0).value
    age=sheet.cell(row, 1).value
    department=sheet.cell(row, 2).value
    conn.execute('Insert into company_table(name,age,department) values(?,?,?)',(name,age,department))
    conn.commit()
    row=row+1





'------------------------------------'
import xlsxwriter 
      
workbook = xlsxwriter.Workbook('DBexcel.xlsx') 
worksheet = workbook.add_worksheet()


row = 0
column = 0
worksheet.write(row, column, str('Name'))
worksheet.write(row, column+1, str('Age'))
worksheet.write(row, column+2, str('Department'))
row = 1
column = 0

data=conn.execute('select * from company_table').fetchall()
  
for i in range(0,len(data)):
    worksheet.write(row, column, str(data[i][0]))
    worksheet.write(row, column+1, str(data[i][1]))
    worksheet.write(row, column+2, str(data[i][2]))
    row += 1
      
workbook.close() 

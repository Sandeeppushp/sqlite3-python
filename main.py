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
    worksheet.write(row, column, str(data[i][1]))
    worksheet.write(row, column+1, str(data[i][2]))
    worksheet.write(row, column+2, str(data[i][3]))
    row += 1
      
workbook.close() 

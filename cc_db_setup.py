import sqlite3
import xlrd

db = sqlite3.connect('clients.db')
cur = db.cursor()

try:
    cur.execute("DROP TABLE client;")
except:
    print("client table does not exist")

cur.execute("""CREATE TABLE client(
                code    INTEGER NOT NULL,
                name    TEXT NOT NULL,
                primary key (code)
                )""")

workbook = xlrd.open_workbook('cc.xls')
worksheet = workbook.sheet_by_name('client')

clients = {}
for i in range(1, worksheet.nrows):
    row = worksheet.row_values(i)
    clients[row[0]] = row[1]

clients = [(v, k) for k, v in clients.items()]

cur.executemany("INSERT INTO client VALUES (?,?);", clients)

for r in cur.execute("SELECT* FROM client;"):
    print(r)

db.commit()
db.close()
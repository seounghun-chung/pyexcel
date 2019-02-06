import sqlite3
import random
con = sqlite3.connect("kospi.db")
cursor = con.cursor()
#cursor.execute("CREATE TABLE kakao(Date text, Open int, High int, Low int, Closing int, Volumn int)")
#cursor.execute("INSERT INTO kakao VALUES('16.06.03', 97000, 98600, 96900, 98000, 321405)")

#data = [('test2', 1,2,3,4,5),
#        ('test3', 1,2,3,4,5),
#        ('test4', 1,2,3,4,5),
#        ('test5', 1,2,3,4,5),]
data = list()        
for ii in range(0,10000):
    data.append(('tests%d'%(ii), ii,random.randint(0,100000),random.randint(0,100000),random.randint(0,100000),random.randint(0,100000)))
        
cursor.executemany("INSERT INTO kakao VALUES (?,?,?,?,?,?)", data)
con.commit()
con.close()

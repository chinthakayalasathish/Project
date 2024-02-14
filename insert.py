import pymysql

connection = pymysql.connect(host="localhost",user="root",passwd="",database="Student")
cursor = connection.cursor()
form='678'
insert_record = "INSERT INTO Registration(FORM_NUMBER,NAME,COURSE,SEMISTER,CONTACT_NUMBER,EMAIL_ID,ADDRESS,ROLL_NUMBER) VALUES(%s,'65','89','98','980','65','89','98');"%(form)
cursor.execute(insert_record)
connection.commit()

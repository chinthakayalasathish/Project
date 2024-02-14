import pymysql

#database connection
connection = pymysql.connect(host="localhost", user="root", passwd="", database="Student")
cursor = connection.cursor()
# Query for creating table
enquiryTableSql = """CREATE TABLE Enquiry(
Date CHAR(100),
Form_Number CHAR(100),
Name  CHAR(100),
Course  CHAR(100),
Phone_Number CHAR(100),
Email_Id  CHAR(100),
Address  CHAR(100),
Demo_Request CHAR(100),
Status  CHAR(100))"""
cursor.execute(enquiryTableSql)
connection.close()


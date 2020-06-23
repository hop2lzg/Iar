import pyodbc


class MSSQL:
    def __init__(self, server, db, user, pwd):
        self.server = server
        self.db = db
        self.user = user
        self.pwd = pwd

    def __GetConnect(self):
        if not self.db:
            raise (NameError, "database no value")
        self.conn = pyodbc.connect(DRIVER="{SQL Server}", SERVER=self.server, DATABASE=self.db, UID=self.user,
                                   PWD=self.pwd)

        cur = self.conn.cursor()
        if not cur:
            raise (NameError, "connnect error")
        else:
            return cur

    def ExecQuery(self, sql):
        cur = self.__GetConnect()
        cur.execute(sql)
        resList = cur.fetchall()

        self.conn.close()
        return resList

    def ExecNonQuery(self, sql):
        cur = self.__GetConnect()
        rowcount = cur.execute(sql).rowcount
        self.conn.commit()
        self.conn.close()
        return rowcount

    def ExecNonQuerys(self, sqls):
        cur = self.__GetConnect()
        rowcount = 0
        for sql in sqls:
            rowcount += cur.execute(sql).rowcount
        self.conn.commit()
        self.conn.close()
        return rowcount

    # def ExecNonQueryMany(self, values):
    #     cur = self.__GetConnect()
    #     cur.executemany("insert into test(id,name,age) values (?,?,?);", values)
    #     self.conn.commit()
    #     self.conn.close()

# sql_server = "ZG"
# sql_database = "EDI"
# sql_user = "ZG\lzg"
# sql_pwd = ""
sql_server = "gtt-azure.database.windows.net"
sql_database = "DB_Users"
sql_user = "gttadmin"
sql_pwd = "Ceoaaa45"
ms = MSSQL(server=sql_server, db=sql_database, user=sql_user, pwd=sql_pwd)

sql = ("select top 3 UserName from T_Users")
rows = ms.ExecQuery(sql)
print rows

# list_data = []
# for row in rows:
#     print row
#     v = {}
#     v["TourCode"] = row.TourCode
#     v["UpdatedTourCode"] = row.UpdatedTourCode
#     print v
#     list_data.append(v)
#
# print list_data
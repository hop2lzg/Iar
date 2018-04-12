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
        # self.conn = pyodbc.connect(DRIVER="{SQL Server}", SERVER=self.server, DATABASE=self.db, UID=self.user,
        #                            PWD=self.pwd,trusted_connection='yes')

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

    def ExecNonQueryMany(self, results):
        cur = self.__GetConnect()
        print results
        print cur.executemany("insert into test(id,name,age) values (?,?,?);", results)
        self.conn.commit()
        self.conn.close()
        # return rowcount

    def ExecNonQuerys(self, sqls):
        cur = self.__GetConnect()
        rowcount = 0
        print "start"
        for sql in sqls:
            rowcount += cur.execute(sql).rowcount

        self.conn.commit()
        self.conn.close()
        return rowcount

# sql_server = "ZG"
# sql_database = "EDI"
# sql_user = "ZG\lzg"
# sql_pwd = ""
sql_server = "107.180.69.206,40815"
sql_database = "CollectData"
sql_user = "sa"
sql_pwd = "Ceoaaa45"
ms = MSSQL(server=sql_server, db=sql_database, user=sql_user, pwd=sql_pwd)

sql = ('''declare @t date
set @t=dateadd(day,-7,getdate())
select t.Id,qc.Id qcId,t.TicketNumber,substring(t.TicketNumber,4,10) Ticket,t.IssueDate,t.ArcNumber,t.PaymentType,
t.Comm,t.TourCode,qc.AGComm UpdatedComm,qc.AGTourCode UpdatedTourCode,qc.OPUser,qc.OPLastUser,t.FareType,qc.AGStatus,
'AG' updatedByRole,iar.Id IarId from Ticket t
left join TicketQC qc
on t.Id=qc.TicketId
left join IarUpdate iar
on t.Id=iar.TicketId
where (qc.ARCupdated=0 or (qc.ARCupdated=1 and iar.IsUpdated=0 and iar.runTimes=0))
and qc.AGStatus=3
--and qc.OPStatus<>2
--and (t.Comm<>qc.AGComm or t.TourCode<>qc.AGTourCode)
and (iar.Commission is null or iar.Commission<>qc.AGComm or iar.TourCode<>qc.AGTourCode)
and (iar.AuditorStatus is null or iar.AuditorStatus=0)
and t.CreateDate>=@t
and t.id='25078ECC-7A99-44DB-BF2A-484C83D40691'
''')
rows = ms.ExecQuery(sql)
print rows

list_data = []
for row in rows:
    print row
    v = {}
    v["TourCode"] = row.TourCode
    v["UpdatedTourCode"] = row.UpdatedTourCode
    print v
    list_data.append(v)

print list_data
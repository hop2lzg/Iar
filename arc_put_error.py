import arc
import ConfigParser
import sys
import datetime


def execute(post, action, token, from_date, to_date):
    arcNumber = post['ArcNumber']
    documentNumber = post['Ticket']
    date = post['IssueDate']
    error_code = post['ErrorCode']
    is_check_payment = False
    updating_commission = post['Commission']
    # if post['Status'] == 3:
    #     is_check_payment = True

    date_time = datetime.datetime.strptime(date, '%Y-%m-%d')
    ped = (date_time+datetime.timedelta(days=(6-date_time.weekday()))).strftime('%d%b%y').upper()
    # logger.info("###### PED: "+ped+" ARC: "+arcNumber+" TKT: "+documentNumber)
    logger.info("###### PED: %s, ARC: %s, TKT: %s, ERROR: %s." % (ped, arcNumber, documentNumber, error_code))
    search_html = arc_model.search(ped, action, arcNumber, token, from_date, to_date, documentNumber)
    if not search_html:
        return

    seqNum, documentNumber = arc_regex.search(search_html)
    if not seqNum:
        return

    modify_html = arc_model.modifyTran(seqNum, documentNumber)
    if not modify_html:
        return

    is_void_pass = arc_regex.check_status(modify_html)
    if is_void_pass == 2:
        post['Status'] = 2
        return
    elif is_void_pass == 1:
        post['Status'] = 4
        return

    token, maskedFC, arc_commission, waiverCode, certificates = arc_regex.modifyTran(modify_html)
    logger.debug("REGEX COMMISSION: %s" % arc_commission)

    if arc_commission is None:
        logger.debug("ARC COMM IS NONE, TKT.# %s, HTML: %s" % (documentNumber, modify_html))
        return

    if error_code == 'AG-Agree':
        logger.debug("AG-Agree commission: %s" % updating_commission)
        arc_commission = updating_commission
    elif error_code == "QC-PROFIT":
        logger.debug("QC-PROFIT TKT: %s, ARC: %s" % (documentNumber, arcNumber))

    if not token:
        return

    if not arc_commission:
        arc_commission = updating_commission

    if arc_commission is None:
        return

    logger.debug("Updating commission: %s, ERROR CODE: %s" % (arc_commission, error_code))
    financialDetails_html = arc_model.financialDetails(token, is_check_payment, arc_commission, waiverCode, maskedFC,
                                                       seqNum, documentNumber, "", "", certificates, error_code,
                                                       agent_codes, is_et_button=True)
    if not financialDetails_html:
        return

    token = arc_regex.itineraryEndorsements(financialDetails_html)
    if token:
        transactionConfirmation_html = arc_model.transactionConfirmation(token)
        if transactionConfirmation_html:
            if transactionConfirmation_html.find('Document has been modified') >= 0:
                post['Status'] = 1
            else:
                logger.warning('update may be error')

arc_model = arc.ArcModel("arc put error")
arc_regex = arc.Regex()
logger = arc_model.logger


logger.debug('--------------<<<START>>>--------------')
conf = ConfigParser.ConfigParser()
conf.read('../iar_update.conf')
agent_codes = conf.get("certificate", "agentCodes").split(',')

# #------------------------------------------sql-------------------------------------
logger.debug('select sql')
sql_server = conf.get("sql", "server")
sql_database = conf.get("sql", "database")
sql_user = conf.get("sql", "user")
sql_pwd = conf.get("sql", "pwd")

# #------------------------------------------sql 44-------------------------------------
section_sql_44 = "sql44"
sql_server_44 = conf.get(section_sql_44, "server")
sql_database_44 = conf.get(section_sql_44, "database")
sql_user_44 = conf.get(section_sql_44, "user")
sql_pwd_44 = conf.get(section_sql_44, "pwd")

ms = arc.MSSQL(server=sql_server, db=sql_database, user=sql_user, pwd=sql_pwd)
sql = ('''declare @t date
set @t=DATEADD(DAY,-1,GETDATE())
select t.Id,[SID],TicketNumber,substring(TicketNumber,4,10) Ticket,IssueDate,ArcNumber,PaymentType,t.Comm,'QC-ERROR' ErrorCode,iar.Id iarId from Ticket t
left join IarUpdate iar
on t.Id=iar.TicketId
where Status not like '[NV]%'
and IssueDate=@t
and ISNULL(McoNumber,'')=''
and (iar.Id is null or iar.IsUpdated=0)
and (iar.AuditorStatus is null or iar.AuditorStatus=0)
and ((QCStatus=2
and TicketNumber not like '04[457]7%' 
and TicketNumber not like '04[457]86%'
and TicketNumber not like '13[49]7%' 
and TicketNumber not like '13[49]86%'
and TicketNumber not like '1477%' 
and TicketNumber not like '14786%'
and TicketNumber not like '2307%' 
and TicketNumber not like '23086%'
and TicketNumber not like '2697%' 
and TicketNumber not like '26986%'
and TicketNumber not like '46[29]7%' 
and TicketNumber not like '46[29]86%'
and TicketNumber not like '5447%' 
and TicketNumber not like '54486%'
and TicketNumber not like '8377%' 
and TicketNumber not like '83786%'
and TicketNumber not like '9577%' 
and TicketNumber not like '95786%'
and TicketNumber not like '890%')
or 
(TicketNumber like '1577%' or TicketNumber like '15786%'
or (TicketNumber like '7817%' or TicketNumber like '78186%')))
union
select t.Id,t.[SID],t.TicketNumber,substring(t.TicketNumber,4,10) Ticket,t.IssueDate,t.ArcNumber,PaymentType,t.Comm,'DUP' ErrorCode,iar.Id iarId from Ticket t
right join TicketDuplicate td
on t.id=td.id
left join IarUpdate iar
on t.Id=iar.TicketId
where td.insertDateTime>=DATEADD(day,-7,getdate())
and ISNULL(McoNumber,'')=''
and td.isARCUpdated=0
and (iar.IsUpdated is null or iar.IsUpdated=0)
and (iar.AuditorStatus is null or iar.AuditorStatus=0)
union
select t.Id,t.[SID],t.TicketNumber,substring(t.TicketNumber,4,10) Ticket,t.IssueDate,t.ArcNumber,PaymentType,
CASE WHEN qc.OPComm is null THEN t.QCComm
ELSE qc.OPComm END Comm
,'AG-Agree' ErrorCode,iar.Id iarId from Ticket t
right join TicketQC qc
on t.id=qc.TicketId
left join IarUpdate iar
on t.Id=iar.TicketId
where qc.AGDate>=DATEADD(day,-3,getdate())
and qc.AGStatus=1
and (iar.Id is null or iar.IsPutError=0)
and (iar.AuditorStatus is null or iar.AuditorStatus=0)
and ISNULL(McoNumber,'')=''
union
select aa.Id,aa.[SID],aa.TicketNumber,aa.Ticket,aa.IssueDate,aa.ArcNumber,aa.PaymentType,
aa.Comm,aa.ErrorCode,aa.iarId from (
select t.Id,t.[SID],t.TicketNumber,substring(t.TicketNumber,4,10) Ticket,t.IssueDate,t.ArcNumber,PaymentType,
case when iar.Commission is null then t.Comm else iar.Commission end Comm,
t.Selling - t.Total + case when iar.Commission is null then t.Comm else iar.Commission end Profit,
t.Base
,'QC-PROFIT' ErrorCode,iar.Id iarId from [CollectData].[dbo].Ticket t
left join [CollectData].[dbo].TicketQC qc
on t.Id=qc.TicketId
left join IarUpdate iar
on t.Id=iar.TicketId
--where t.CreateDate>='20180305' and t.CreateDate<'20180306'
where t.CreateDate>=@t and t.CreateDate<DATEADD(DAY,1,@t)
and t.[Status] not like '[NV]%'
and (t.TicketType is null or t.TicketType<>'EX')
and t.Comm=0
and t.FareType in ('SR','BULK')
and qc.OPStatus=14
and iar.isPutError=0
and iar.AuditorStatus=0
and t.Selling>0
and ISNULL(McoNumber,'')=''
) aa
where aa.Profit<0
and -aa.Profit > Base*0.05
union
select t.Id,[SID],TicketNumber,substring(TicketNumber,4,10) Ticket,IssueDate,ArcNumber,PaymentType,t.Comm,'QC-ERROR' ErrorCode,iar.Id iarId from Ticket t
left join IarUpdate iar
on t.Id=iar.TicketId
where t.sourceFrom='AUTOTICKET'
and t.IssueDate=@t
and PaymentType='C'
and (t.TicketNumber like '78[14]%' or t.TicketNumber like '731%')
and t.Charge<t.Total
and t.GDS='1A'
and ISNULL(t.McoNumber,'')<>''
and (iar.Id is null or iar.IsPutError=0)
and (iar.AuditorStatus is null or iar.AuditorStatus=0)
and ISNULL(McoNumber,'')=''
union
select t.Id,[SID],TicketNumber,substring(TicketNumber,4,10) Ticket,IssueDate,ArcNumber,PaymentType,t.Comm,'QC-RE' ErrorCode,iar.Id iarId from Ticket t
left join IarUpdate iar
on t.Id=iar.TicketId
where t.IssueDate=@t
and ((t.ArcNumber='45668571' and t.AgentSign in ('MM','MG','ZG','YU','S7'))
	or (t.ArcNumber='45666574' and t.AgentSign in ('PF','NQ','LL','XL')))
and TicketNumber not like '180%'
and TicketNumber not like '784%'
and TicketNumber not like '230%'
and TicketNumber not like '098%'
and TicketNumber not like '890%'
and (TicketNumber like '[0-9][0-9][0-9]7%' or TicketNumber like '[0-9][0-9][0-9]86%')
and ISNULL(t.McoNumber,'')=''
and (iar.Id is null or iar.IsPutError=0)
and (iar.AuditorStatus is null or iar.AuditorStatus=0)
union
select t.Id,[SID],TicketNumber,substring(TicketNumber,4,10) Ticket,IssueDate,ArcNumber,PaymentType,t.Comm,'QC-ERROR' ErrorCode,iar.Id iarId from Ticket t
left join IarUpdate iar
on t.Id=iar.TicketId
where t.IssueDate=@t
and PaymentType='C'
and Status not like '[NV]%'
and 
(TicketNumber like '1[02]55%'
or 
((TicketNumber like '1087%' or TicketNumber like '10886%'
or TicketNumber like '0757%' or TicketNumber like '07586%')
and (FareType='BULK' or FareType='SR')
)
or 
(TicketNumber like '1085%' or TicketNumber like '0755%')
)
and (iar.Id is null or iar.IsPutError=0)
and (iar.AuditorStatus is null or iar.AuditorStatus=0)
order by ArcNumber,Ticket
''')

list_data = arc_model.load()
list_id = []

if not list_data:
    rows = ms.ExecQuery(sql)
    if len(rows) == 0:
        sys.exit(0)
    list_data = []
    for row in rows:
        v = {}
        v['Id'] = row.Id
        if v['Id'] not in list_id:
            list_id.append(v['Id'])
        else:
            continue
        v['TicketNumber'] = row.TicketNumber
        v['Ticket'] = row.Ticket
        v['IssueDate'] = str(row.IssueDate)
        v['ArcNumber'] = row.ArcNumber
        v['Status'] = 0
        v['ErrorCode'] = row.ErrorCode
        v['IarId'] = row.iarId
        v['Commission'] = None
        if row.Comm is not None:
            v['Commission'] = str(row.Comm)

        list_data.append(v)

ms_44 = arc.MSSQL(server=sql_server_44, db=sql_database_44, user=sql_user_44, pwd=sql_pwd_44)
sql_44 = "select distinct AccountCode from dbo.T_Users where Highrisk=1"

high_risk_account_code_rows = ms_44.ExecQuery(sql_44)
if len(high_risk_account_code_rows) > 0:
    high_risk_account_codes = []
    for high_risk_account_code_row in high_risk_account_code_rows:
        high_risk_account_code = high_risk_account_code_row.AccountCode
        if not high_risk_account_code:
            continue

        if high_risk_account_code not in high_risk_account_codes:
            high_risk_account_codes.append(high_risk_account_code)

    if high_risk_account_codes:
        acct_where = ""
        if len(high_risk_account_codes) == 1:
            acct_where = "and t.AccountCode='" + high_risk_account_codes[0] + "' "
        else:
            acct_where = "and t.AccountCode in ('" + (",".join(high_risk_account_codes).replace(",", "','")) + "') "

        sql_high_risk = '''
        declare @date date
set @date=DATEADD(DAY,-1,GETDATE())
select t.Id,[SID],TicketNumber,substring(TicketNumber,4,10) Ticket,IssueDate,ArcNumber,PaymentType,t.Comm,'QC-HRISK' ErrorCode,iar.Id iarId from Ticket t
left join IarUpdate iar
on t.Id=iar.TicketId
where t.IssueDate=@date
and t.Status not like '[NV]%'
and ISNULL(t.McoNumber,'')=''
''' + acct_where + '''
and (iar.Id is null or iar.IsPutError=0)
and (iar.AuditorStatus is null or iar.AuditorStatus=0)
'''
        high_risk_rows = ms.ExecQuery(sql_high_risk)
        if len(high_risk_rows) > 0:
            for high_risk_row in high_risk_rows:
                v = {}
                v['Id'] = high_risk_row.Id
                if v['Id'] not in list_id:
                    list_id.append(v['Id'])
                else:
                    continue
                v['TicketNumber'] = high_risk_row.TicketNumber
                v['Ticket'] = high_risk_row.Ticket
                v['IssueDate'] = str(high_risk_row.IssueDate)
                v['ArcNumber'] = high_risk_row.ArcNumber
                v['Status'] = 0
                v['ErrorCode'] = high_risk_row.ErrorCode
                v['IarId'] = high_risk_row.iarId
                v['Commission'] = None
                if high_risk_row.Comm is not None:
                    v['Commission'] = str(high_risk_row.Comm)

                list_data.append(v)


def run(user_name, datas):
    # ----------------------login
    logger.debug(user_name)
    password = conf.get("login", user_name)

    if not arc_model.execute_login(user_name, password):
        return

    # -------------------go to IAR
    iar_html = arc_model.iar()
    if not iar_html:
        logger.error('iar error')
        return

    ped, action, arcNumber = arc_regex.iar(iar_html, False)
    if not action:
        logger.error('regex iar error')
        arc_model.logout()
        return

    listTransactions_html = arc_model.listTransactions(ped, action, arcNumber)
    if not listTransactions_html:
        logger.error('listTransactions error')
        arc_model.iar_logout(ped, action, arcNumber)
        arc_model.logout()
        return

    token, from_date, to_date = arc_regex.listTransactions(listTransactions_html)
    if not token:
        logger.error('regex listTransactions error')
        arc_model.iar_logout(ped, action, arcNumber)
        arc_model.logout()
        return

    try:
        for data in datas:
            if data['Status'] != 0:
                continue
            execute(data, action, token, from_date, to_date)
    except Exception as ex:
        logger.critical(ex)
    arc_model.iar_logout(ped, action, arcNumber)
    arc_model.logout()


def update(datas):
    ids = []
    for data in datas:
        if data['ErrorCode'] == 'DUP' and data['Status'] != 0 and data['Id'] not in ids:
            ids.append("'" + data['Id'] + "'")

    if ids:
        update_sql = "update TicketDuplicate set isARCUpdated=1 where id in (%s)" % ','.join(ids)
        logger.debug(update_sql)
        if ms.ExecNonQuery(update_sql) > 0:
            logger.info('update sql success')
        else:
            logger.error('update sql error')


def insert(datas):
    if not datas:
        return

    ids = []
    sqls = []
    for data in datas:
        if data['Id'] in ids:
            continue
        ids.append(data['Id'])

        if not data['IarId']:
            sqls.append('''insert into IarUpdate(Id,TicketId,channel,IsPutError,errorCode) values (newid(),'%s',4,1,'%s');''' % (
                data['Id'], data['ErrorCode']))
        else:
            sqls.append('''update IarUpdate set channel=4,IsPutError=1,errorCode='%s' where Id='%s';''' % (
                data['ErrorCode'], data['IarId']))

    if not sqls:
        logger.warn("Insert or update no data")
        return

    logger.info("".join(sqls))
    rowcount = ms.ExecNonQuerys(sqls)
    if rowcount != len(sqls):
        logger.warn("updating:%s, updated:%s" % (len(sqls), rowcount))

    if rowcount > 0:
        logger.info('insert and update success')
    else:
        logger.error('insert and update error')

try:
    section = "arc"
    for option in conf.options(section):
        logger.debug(option)
        arc_numbers = conf.get(section, option).split(',')
        list_data_account = filter(lambda x: x['ArcNumber'] in arc_numbers, list_data)
        if not list_data_account:
            continue
        account_id = "muling-"
        if option == "all":
            account_id = conf.get("accounts", "all")
        else:
            account_id = account_id + option
        run(account_id, list_data_account)
except Exception as e:
    logger.critical(e)
finally:
    arc_model.store(list_data)

try:
    update(list_data)
except Exception as e:
    logger.critical(e)

try:
    insert(list_data)
except Exception as e:
    logger.critical(e)

# ##---------------------------send email
mail_smtp_server = conf.get("email", "smtp_server")
mail_from_addr = conf.get("email", "from")
mail_to_addr = conf.get("email", "to_put_error").split(';')
mail_subject = conf.get("email", "subject")+' put error'

try:
    body = ''
    for i in list_data:
        status = 'No'
        if i['Status'] == 1:
            status = 'Yes'
        elif i['Status'] == 2:
            status = 'Void'
        elif i['Status'] == 3:
            status = 'Check'
        elif i['Status'] == 4:
            status = 'Pass'

        body = body+'''<tr>
            <td>%s</td>
            <td>%s</td>
            <td>%s</td>
            <td>%s</td>
        </tr>''' % (i['ArcNumber'], i['TicketNumber'], i['IssueDate'], status)

    body = '''<table border=1>
    <thead>
        <tr>
            <th>ARC</th>
            <th>TicketNumber</th>
            <th>Date</th>
            <th>Updated</th>
        </tr>
    </thead>
    <tbody>%s
    </tbody>
</table>
''' % body
    mail = arc.Email(smtp_server=mail_smtp_server)
    mail.send(mail_from_addr, mail_to_addr, mail_subject, body)
except Exception as e:
    logger.critical(e)


logger.debug('--------------<<<END>>>--------------')

# if __name__ == "__main__":
# 	print "main"


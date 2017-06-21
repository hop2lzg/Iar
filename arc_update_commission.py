import arc
import ConfigParser
import sys
import datetime


def execute(post, action, token, from_date, to_date):
    arcNumber = post['ArcNumber']
    documentNumber = post['Ticket']
    date = post['IssueDate']
    commission = post['QCComm']
    tour_code = ""
    if post['TourCode']:
        tour_code = post['TourCode']
    qc_tour_code = ""
    if post['QCTourCode']:
        qc_tour_code = post['QCTourCode']

    is_check = False

    if post['Status'] == 3:
        is_check = True

    date_time = datetime.datetime.strptime(date, '%Y-%m-%d')
    ped = (date_time + datetime.timedelta(days=(6 - date_time.weekday()))).strftime('%d%b%y').upper()

    logger.info("UPDATING PED: " + ped + " arc: " + arcNumber + " tkt: " + documentNumber)

    search_html = arc_model.search(ped, action, arcNumber, token, from_date, to_date, documentNumber)
    if not search_html:
        return
    seqNum, documentNumber = arc_regex.search(search_html)
    if not seqNum:
        return
    modify_html = arc_model.modifyTran(seqNum, documentNumber)
    if not modify_html:
        return
    voided_index = modify_html.find('Document is being displayed as view only')
    if voided_index >= 0:
        post['Status'] = 2
        return

    token, maskedFC, arc_commission, waiverCode, certificates = arc_regex.modifyTran(modify_html)
    if not token:
        return
    post['ArcComm'] = arc_commission

    financialDetails_html = arc_model.financialDetails(token, is_check, commission, waiverCode, maskedFC, seqNum,
                                                       documentNumber, tour_code, qc_tour_code, certificates)
    if not financialDetails_html:
        return

    token, arc_tour_code, backOfficeRemarks, ticketDesignators = arc_regex.financialDetails(financialDetails_html)
    if not token:
        return
    post['ArcTourCode'] = arc_tour_code

    if ticketDesignators:
        list_ticketDesignator = []
        for ticketDesignator in ticketDesignators:
            list_ticketDesignator.append(ticketDesignator[1])
        post['TicketDesignator'] = '/'.join(list_ticketDesignator)
    if tour_code != qc_tour_code:
        itineraryEndorsements_html = arc_model.itineraryEndorsements(token, qc_tour_code, backOfficeRemarks,
                                                                     ticketDesignators)
        if not itineraryEndorsements_html:
            return
        token = arc_regex.itineraryEndorsements(itineraryEndorsements_html)
    if token:
        transactionConfirmation_html = arc_model.transactionConfirmation(token)
        if transactionConfirmation_html:
            if transactionConfirmation_html.find('Document has been modified') >= 0:
                post['Status'] = 1
            else:
                logger.warning('update may be error')

                # updated_ticket_number,updated_commission=RegexTransactionConfirmation(transactionConfirmation_html)

                # # logger.info("documentNumber:"+updated_ticket_number+" commission:"+updated_commission)
                # # print documentNumber,updated_ticket_number,type(documentNumber),type(updated_ticket_number)
                # # print commission,updated_commission,type(commission),type(updated_commission)
                # if(documentNumber==updated_ticket_number and str(commission)==updated_commission):
                # 	# Write(documentNumber)
                # 	post['Status']=1
                # else:
                # 	print 'no write'


def check(post, action, token, from_date, to_date):
    arcNumber = post['ArcNumber']
    documentNumber = post['Ticket']
    date = post['IssueDate']
    commission = post['QCComm']
    tour_code = ""
    if post['TourCode']:
        tour_code = post['TourCode']
    qc_tour_code = ""
    if post['QCTourCode']:
        qc_tour_code = post['QCTourCode']

    is_check = False

    if post['Status'] == 3:
        is_check = True

    date_time = datetime.datetime.strptime(date, '%Y-%m-%d')
    ped = (date_time + datetime.timedelta(days=(6 - date_time.weekday()))).strftime('%d%b%y').upper()

    logger.info("CHECK PED: " + ped + " arc: " + arcNumber + " tkt: " + documentNumber)

    search_html = arc_model.search(ped, action, arcNumber, token, from_date, to_date, documentNumber)
    if not search_html:
        return
    seqNum, documentNumber = arc_regex.search(search_html)
    if not seqNum:
        return
    modify_html = arc_model.modifyTran(seqNum, documentNumber)
    if not modify_html:
        return
    voided_index = modify_html.find('Document is being displayed as view only')
    if voided_index >= 0:
        post['Status'] = 2
        return

    token, maskedFC, arc_commission, waiverCode, certificates = arc_regex.modifyTran(modify_html)
    if not token:
        return
    post['ArcCommUpdated'] = arc_commission

    financialDetails_html = arc_model.financialDetails(token, is_check, commission, waiverCode, maskedFC, seqNum,
                                                       documentNumber, tour_code, qc_tour_code, certificates, True)
    if not financialDetails_html:
        return

    token, arc_tour_code, backOfficeRemarks, ticketDesignators = arc_regex.financialDetails(financialDetails_html)
    if not token:
        return
    post['ArcTourCodeUpdated'] = arc_tour_code


def insert(datas):
    if not datas:
        return
    insert_sql = ''
    ids = []
    for data in datas:
        if data['Id'] in ids:
            continue
        ids.append(data['Id'])
        data['QCComm'] = "" if data['QCComm'] is None else data['QCComm']
        data['ArcCommUpdated'] = "" if data['ArcCommUpdated'] is None else data['ArcCommUpdated']
        data['QCTourCode'] = "" if data['QCTourCode'] is None else data['QCTourCode']
        data['ArcTourCodeUpdated'] = "" if data['ArcTourCodeUpdated'] is None else data['ArcTourCodeUpdated']
        is_updated = 0
        if data['QCComm'] == data['ArcCommUpdated'] and data['QCTourCode'] == data['ArcTourCodeUpdated']:
            is_updated = 1
        comm = None
        try:
            comm = float(data['ArcCommUpdated'])
        except ValueError:
            comm = "null"
        if not data['ArcId']:
            insert_sql = insert_sql + '''insert into IarUpdate(Id,Commission,TourCode,TicketDesignator,IsUpdated,TicketId) values (newid(),%s,'%s','%s',%d,'%s');''' % (
                comm, data['ArcTourCodeUpdated'], data['TicketDesignator'], is_updated, data['Id'])
        else:
            insert_sql = insert_sql + '''update IarUpdate set Commission=%s,TourCode='%s',TicketDesignator='%s',IsUpdated=%d where Id='%s';''' % (
                comm, data['ArcTourCodeUpdated'], data['TicketDesignator'], is_updated, data['ArcId'])

    logger.info(insert_sql)
    rowcount = ms.ExecNonQuery(insert_sql)
    if rowcount > 0:
        logger.info('insert success')
    else:
        logger.error('insert error')


def run(user_name, datas):
    # ----------------------login
    password = conf.get("login", user_name)
    login_html = arc_model.login(user_name, password)
    if login_html.find('You are already logged into My ARC') < 0 and login_html.find('Account Settings :') < 0:
        logger.error('login error')
        return
    # -----------------go to IAR
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
            if data['Status'] != 0 and data['Status'] != 3:
                continue
            execute(data, action, token, from_date, to_date)

            if data['Status'] == 1 or data['Status'] == 3:
                check(data, action, token, from_date, to_date)
    except Exception as ex:
        logger.critical(ex)
    # finally:
    #     arc_model.store(datas)

    arc_model.iar_logout(ped, action, arcNumber)
    arc_model.logout()


arc_model = arc.ArcModel("arc update commission")
arc_regex = arc.Regex()
logger = arc_model.logger

logger.debug('--------------<<<START>>>--------------')

logger.debug('select sql')

conf = ConfigParser.ConfigParser()
conf.read('../iar_update.conf')
sql_server = conf.get("sql", "server")
sql_database = conf.get("sql", "database")
sql_user = conf.get("sql", "user")
sql_pwd = conf.get("sql", "pwd")

ms = arc.MSSQL(server=sql_server, db=sql_database, user=sql_user, pwd=sql_pwd)
sql = ('''declare @t date
set @t=DATEADD(DAY,-1,GETDATE())
select t.Id,[SID],TicketNumber,substring(TicketNumber,4,10) Ticket,IssueDate,ArcNumber,t.Comm,QCComm,t.TourCode,QCTourCode,PaymentType,iu.Id IarId from Ticket t
left join IarUpdate iu
on t.Id=iu.TicketId
where Status not like '[NV]%'
and FareType not in ('BULK','SR')
and IssueDate=@t
and t.Comm>=0
and t.Comm<QCComm-5
and (ISNULL(t.TourCode,'')='' or ((TicketNumber like '00[16]7%' or TicketNumber like '0147%' or TicketNumber like '1767%' or TicketNumber like '6077%' or TicketNumber like '0727%' or TicketNumber like '0657%' or TicketNumber like '1577%' or TicketNumber like '2357%' or TicketNumber like '5557%' or TicketNumber like '1607%') and t.TourCode<>''))
order by ArcNumber,Ticket
''')

list_data = arc_model.load()

if not list_data:
    rows = ms.ExecQuery(sql)
    # print len(rows)
    if len(rows) == 0:
        sys.exit(0)
    list_data = []
    for row in rows:
        v = {}
        v['Id'] = row.Id
        v['TicketNumber'] = row.TicketNumber
        v['Ticket'] = row.Ticket
        v['IssueDate'] = str(row.IssueDate)
        v['ArcNumber'] = row.ArcNumber
        v['Comm'] = str(row.Comm)
        v['QCComm'] = str(row.QCComm)
        v['TourCode'] = row.TourCode
        v['QCTourCode'] = row.QCTourCode
        v['ArcComm'] = ''
        v['ArcTourCode'] = ''
        v['TicketDesignator'] = ''
        v['ArcCommUpdated'] = ''
        v['ArcTourCodeUpdated'] = ''
        v['Status'] = 0
        v['ArcId'] = row.IarId
        if row.PaymentType != 'C':
            v['Status'] = 3
        list_data.append(v)


try:
    section = "arc"
    for option in conf.options(section):
        logger.debug(option)
        # print option
        arc_numbers = conf.get(section, option).split(',')
        # print arc_numbers
        list_data_account = filter(lambda x: x['ArcNumber'] in arc_numbers, list_data)
        if not list_data_account:
            continue

        account_id = "muling-"
        if option == "all":
            account_id = "mulingpeng"
        else:
            account_id = account_id + option

        run(account_id, list_data_account)
except Exception as e:
    logger.critical(e)
finally:
    arc_model.store(list_data)

#
# name = "mulingpeng"
# password = conf.get("login", name)
# login_html = arc_model.login(name, password)
# if login_html.find('You are already logged into My ARC') < 0 and login_html.find('Account Settings :') < 0:
#     logger.error('login error')
#     sys.exit(0)

# # -------------------go to IAR
# iar_html = arc_model.iar()
# if not iar_html:
#     logger.error('iar error')
#     sys.exit(0)
# ped, action, arcNumber = arc_regex.iar(iar_html, False)
# if not action:
#     logger.error('regex iar error')
#     arc_model.logout()
#     sys.exit(0)
#
# listTransactions_html = arc_model.listTransactions(ped, action, arcNumber)
# if not listTransactions_html:
#     logger.error('listTransactions error')
#     arc_model.iar_logout(ped, action, arcNumber)
#     arc_model.logout()
#     sys.exit(0)
#
# token, from_date, to_date = arc_regex.listTransactions(listTransactions_html)
# if not token:
#     logger.error('regex listTransactions error')
#     arc_model.iar_logout(ped, action, arcNumber)
#     arc_model.logout()
#     sys.exit(0)

# list_data=[]
# v={}
# v['Id']='id'
# v['TicketNumber']='0017996642045'
# v['Ticket']='7996642045'
# v['IssueDate']='2017-04-23'
# v['ArcNumber']='05507073'
# v['Comm']='0.00'
# v['QCComm']='8.76'
# v['TourCode']=None
# v['QCTourCode']='S602'
# v['ArcComm']=''
# v['ArcTourCode']=None
# v['TicketDesignator']=''
# v['ArcCommUpdated']=''
# v['ArcTourCodeUpdated']=''
# v['Status']=0
# list_data.append(v)

# try:
#     for data in list_data:
#         if data['Status'] != 0 and data['Status'] != 3:
#             continue
#         # # print 'source'
#         # # print data
#         execute(data, action, token, from_date, to_date)
#         # # print 'execute'
#         # # print data
#         if data['Status'] == 1 or data['Status'] == 3:
#             check(data, action, token, from_date, to_date)
#             # print 'check'
#             # print data
#             # time.sleep(3)
# except Exception as e:
#     print e
#     logger.critical(e)
# finally:
#     arc_model.store(list_data)

try:
    insert(list_data)
except Exception as e:
    logger.critical(e)

###-----------------export excel
file_name = "iar_update_commission"
try:
    arc_model.exportExcel(list_data, file_name)
except Exception as e:
    logger.critical(e)

mail_smtp_server = conf.get("email", "smtp_server")
mail_from_addr = conf.get("email", "from")
mail_to_addr = conf.get("email", "to_update").split(';')
mail_subject = conf.get("email", "subject") + " for accounting"
try:
    body = ''
    for i in list_data:
        status, updated = arc_model.convertStatus(i)
        body = body + '''<tr>
            <td>%s</td>
            <td>%s</td>
            <td>%s</td>
            <td>%s</td>
            <td>%s</td>
            <td>%s</td>
            <td>%s</td>
            <td>%s</td>
            <td>%s</td>
            <td>%s</td>
        </tr>''' % (i['ArcNumber'], i['TicketNumber'],
                    i['IssueDate'], i['Comm'], i['QCComm'], i['ArcCommUpdated'], i['TourCode'], i['QCTourCode'],
                    i['ArcTourCodeUpdated'], updated)

    # print body
    body = '''<table border=1>
    <thead>
        <tr>
            <th>ARC</th>
            <th>TicketNumber</th>
            <th>Date</th>
            <th>TKCM</th>
            <th>QCCM</th>
            <th>ARCCM</th>
            <th>TKTC</th>
            <th>QCTC</th>
            <th>ARCTC</th>
            <th>Status</th>
        </tr>
    </thead>
    <tbody>%s
    </tbody></table>''' % body

    mail = arc.Email(smtp_server=mail_smtp_server)
    mail.send(mail_from_addr, mail_to_addr, mail_subject, body, ['excel/' + file_name + '.xlsx'])

    logger.info('email sent')
except Exception as e:
    logger.critical(e)

logger.debug('--------------<<<END>>>--------------')
import arc
import ConfigParser
import sys
import datetime


# def execute(post, action, token, from_date, to_date):
#     arcNumber = post['ArcNumber']
#     documentNumber = post['Ticket']
#     date = post['IssueDate']
#     commission = post['QCComm']
#     tour_code = ""
#     if post['TourCode']:
#         tour_code = post['TourCode']
#     qc_tour_code = ""
#     if post['QCTourCode']:
#         qc_tour_code = post['QCTourCode']
#
#     is_check_payment = False
#
#     if post['Status'] == 3:
#         is_check_payment = True
#
#     date_time = datetime.datetime.strptime(date, '%Y-%m-%d')
#     ped = (date_time + datetime.timedelta(days=(6 - date_time.weekday()))).strftime('%d%b%y').upper()
#
#     logger.info("UPDATING PED: " + ped + " arc: " + arcNumber + " tkt: " + documentNumber)
#
#     search_html = arc_model.search(ped, action, arcNumber, token, from_date, to_date, documentNumber)
#     if not search_html:
#         return
#     seqNum, documentNumber = arc_regex.search(search_html)
#     if not seqNum:
#         return
#     modify_html = arc_model.modifyTran(seqNum, documentNumber)
#     if not modify_html:
#         return
#
#     is_void_pass = arc_regex.check_status(modify_html)
#     if is_void_pass == 2:
#         post['Status'] = 2
#         return
#     elif is_void_pass == 1:
#         post['Status'] = 4
#         return
#     # voided_index = modify_html.find('Document is being displayed as view only')
#     # if voided_index >= 0:
#     #     post['Status'] = 2
#     #     return
#
#     token, maskedFC, arc_commission, waiverCode, certificates = arc_regex.modifyTran(modify_html)
#     if not token:
#         return
#
#     post['ArcComm'] = arc_commission
#     financialDetails_html = arc_model.financialDetails(token, is_check_payment, commission, waiverCode, maskedFC,
#                                                        seqNum, documentNumber, tour_code, qc_tour_code, certificates,
#                                                        "MJ", agent_codes, is_check_update=False)
#     # financialDetails_html = arc_model.financialDetails(token, is_check, commission, waiverCode, maskedFC, seqNum,
#     #                                                    documentNumber, tour_code, qc_tour_code, certificates)
#
#     if not financialDetails_html:
#         return
#
#     token, arc_tour_code, backOfficeRemarks, ticketDesignators = arc_regex.financialDetails(financialDetails_html)
#     if not token:
#         return
#     post['ArcTourCode'] = arc_tour_code
#
#     if ticketDesignators:
#         list_ticketDesignator = []
#         for ticketDesignator in ticketDesignators:
#             list_ticketDesignator.append(ticketDesignator[1])
#         post['TicketDesignator'] = '/'.join(list_ticketDesignator)
#     # if tour_code != qc_tour_code:
#     #     itineraryEndorsements_html = arc_model.itineraryEndorsements(token, qc_tour_code, backOfficeRemarks,
#     #                                                                  ticketDesignators)
#     #     if not itineraryEndorsements_html:
#     #         return
#     #     token = arc_regex.itineraryEndorsements(itineraryEndorsements_html)
#     # if token:
#     #     transactionConfirmation_html = arc_model.transactionConfirmation(token)
#     #     if transactionConfirmation_html:
#     #         if transactionConfirmation_html.find('Document has been modified') >= 0:
#     #             post['Status'] = 1
#     #         else:
#     #             logger.warning('update may be error')


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

    is_check_payment = False

    if post['Status'] == 3:
        is_check_payment = True

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

    is_void_pass = arc_regex.check_status(modify_html)

    if is_void_pass == 2:
        post['Status'] = 2
        return
    elif is_void_pass == 1:
        post['Status'] = 4

    # voided_index = modify_html.find('Document is being displayed as view only')
    # if voided_index >= 0:
    #     post['Status'] = 4
    #     if modify_html.find('Unable to modify a voided document') >= 0:
    #         post['Status'] = 2
    #         return

    token, maskedFC, arc_commission, waiverCode, certificates = arc_regex.modifyTran(modify_html)
    if not token:
        return
    post['ArcCommUpdated'] = arc_commission
    financialDetails_html = arc_model.financialDetails(token, is_check_payment, commission, waiverCode, maskedFC,
                                                       seqNum, documentNumber, tour_code, qc_tour_code, certificates,
                                                       "MJ", agent_codes, is_check_update=True)
    # financialDetails_html = arc_model.financialDetails(token, is_check, commission, waiverCode, maskedFC, seqNum,
    #                                                    documentNumber, tour_code, qc_tour_code, certificates, True)
    if not financialDetails_html:
        return

    token, arc_tour_code, backOfficeRemarks, ticketDesignators = arc_regex.financialDetails(financialDetails_html)
    if not token:
        return

    post['ArcTourCodeUpdated'] = arc_tour_code
    if ticketDesignators:
        list_ticketDesignator = []
        for ticketDesignator in ticketDesignators:
            list_ticketDesignator.append(ticketDesignator[1])
        post['TicketDesignator'] = '/'.join(list_ticketDesignator)


def update(datas):
    ids = []
    for data in datas:
        if data['Status'] != 0:
            if data['QcId'] not in ids:
                ids.append("'" + data['QcId'] + "'")
    if ids:
        update_sql = "update TicketQC set ARCupdated=1 where Id in (%s)" % ','.join(ids)
        logger.debug(update_sql)
        if ms.ExecNonQuery(update_sql) > 0:
            logger.info('update sql success')
        else:
            logger.error('update sql error')


def insert(datas):
    if not datas:
        return
    # insert_sql = ''

    ids = []
    sqls = []
    for data in datas:
        if data['Id'] in ids:
            continue
        ids.append(data['Id'])
        data['QCComm'] = "" if data['QCComm'] is None else data['QCComm']
        data['ArcCommUpdated'] = "" if data['ArcCommUpdated'] is None else data['ArcCommUpdated']
        data['QCTourCode'] = "" if data['QCTourCode'] is None else data['QCTourCode'].upper()
        data['ArcTourCodeUpdated'] = "" if data['ArcTourCodeUpdated'] is None else data['ArcTourCodeUpdated']
        is_updated = 0
        if data['QCComm'] == data['ArcCommUpdated'] and data['QCTourCode'] == data['ArcTourCodeUpdated']:
            is_updated = 1

        logger.debug("Ticket ID:%s, IAR ID:%s, QC comm:%s, IAR updated comm:%s,QC tour code:%s, IAR tour code:%s." %
                     (data['Id'], data['ArcId'], data['QCComm'], data['ArcCommUpdated'], data['QCTourCode'], data['ArcTourCodeUpdated']))
        comm = None
        try:
            comm = float(data['ArcCommUpdated'])
        except ValueError:
            comm = "null"
        if not data['ArcId']:
            sqls.append('''insert into IarUpdate(Id,Commission,TourCode,TicketDesignator,IsUpdated,TicketId,channel) values (newid(),%s,'%s','%s',%d,'%s',3);''' % (
                comm, data['ArcTourCodeUpdated'], data['TicketDesignator'], is_updated, data['Id']))
            # insert_sql = insert_sql + '''insert into IarUpdate(Id,Commission,TourCode,TicketDesignator,IsUpdated,TicketId,channel) values (newid(),%s,'%s','%s',%d,'%s',3);''' % (
            #     comm, data['ArcTourCodeUpdated'], data['TicketDesignator'], is_updated, data['Id'])
        else:
            sqls.append('''update IarUpdate set Commission=%s,TourCode='%s',TicketDesignator='%s',IsUpdated=%d,channel=5, updateDateTime=GETDATE() where Id='%s';''' % (
                comm, data['ArcTourCodeUpdated'], data['TicketDesignator'], is_updated, data['ArcId']))
            # insert_sql = insert_sql + '''update IarUpdate set Commission=%s,TourCode='%s',TicketDesignator='%s',IsUpdated=%d,channel=5, updateDateTime=GETDATE() where Id='%s';''' % (
            #     comm, data['ArcTourCodeUpdated'], data['TicketDesignator'], is_updated, data['ArcId'])

    if not sqls:
        logger.warn("Insert or update no data")
        return

    logger.info("".join(sqls))
    rowcount = ms.ExecNonQuerys(sqls)
    if rowcount != len(sqls):
        logger.warn("update:%s, updated:%s" % (len(sqls), rowcount))

    if rowcount > 0:
        logger.info('insert success')
    else:
        logger.error('insert error')


arc_model = arc.ArcModel("arc update commission by user sync")
arc_regex = arc.Regex()
logger = arc_model.logger

logger.debug('--------------<<<START>>>--------------')
logger.debug('select sql')

conf = ConfigParser.ConfigParser()
conf.read('../iar_update.conf')
agent_codes = conf.get("certificate", "agentCodes").split(',')
sql_server = conf.get("sql", "server")
sql_database = conf.get("sql", "database")
sql_user = conf.get("sql", "user")
sql_pwd = conf.get("sql", "pwd")
ms = arc.MSSQL(server=sql_server, db=sql_database, user=sql_user, pwd=sql_pwd)

mail_smtp_server = conf.get("email", "smtp_server")
mail_from_addr = conf.get("email", "from")
mail_to_addr = conf.get("email", "to_update_user_sync").split(';')
mail_subject = conf.get("email", "subject") + " by user sync"


mail = arc.Email(smtp_server=mail_smtp_server)

sql = ('''
declare @start date
declare @end date
set @start=dateadd(day,-5,getdate())
set @end=GETDATE()
select t.Id,qc.Id qcId,t.TicketNumber,substring(t.TicketNumber,4,10) Ticket,t.IssueDate,t.ArcNumber,t.PaymentType,
t.Comm,t.TourCode,qc.OPComm UpdateComm,qc.OPTourCode UpdateTourCode,iar.Id IarId from Ticket t
left join TicketQC qc
on t.Id=qc.TicketId
left join IarUpdate iar
on t.Id=iar.TicketId
where  --(qc.OPUser in ('GW','KM','PF') or OPLastUser in ('GW','KM','PF'))
qc.OPStatus=2
and qc.AGStatus<>3
and t.IssueDate>=@start and t.IssueDate<@end
and (iar.Commission is null or qc.OPComm<>iar.Commission or qc.OPTourCode<>iar.TourCode)
and (iar.AuditorStatus is null or iar.AuditorStatus=0)
union
select t.Id,qc.Id qcId,t.TicketNumber,substring(t.TicketNumber,4,10) Ticket,t.IssueDate,t.ArcNumber,t.PaymentType,
t.Comm,t.TourCode,qc.AGComm UpdateComm,qc.AGTourCode UpdateTourCode,iar.Id IarId from Ticket t
left join TicketQC qc
on t.Id=qc.TicketId
left join IarUpdate iar
on t.Id=iar.TicketId
where qc.AGStatus=3
and (iar.Commission is null or iar.Commission<>qc.AGComm or iar.TourCode<>qc.AGTourCode)
and t.IssueDate>=@start
and t.IssueDate<@end
and (iar.AuditorStatus is null or iar.AuditorStatus=0)
order by t.ArcNumber
''')

rows = ms.ExecQuery(sql)
# print len(rows)

if len(rows) == 0:
    mail.send(mail_from_addr, mail_to_addr, mail_subject, "No data,please confirm.")
    sys.exit(0)

# list_data_sql=arc_model.load()
list_data = []
for i in rows:
    v = {}
    v['Id'] = i.Id
    v['QcId'] = i.qcId
    v['TicketNumber'] = i.TicketNumber
    v['Ticket'] = i.Ticket
    v['IssueDate'] = str(i.IssueDate)
    v['ArcNumber'] = i.ArcNumber
    v['Comm'] = str(i.Comm)
    v['QCComm'] = str(i.UpdateComm)
    v['TourCode'] = i.TourCode
    v['QCTourCode'] = i.UpdateTourCode
    v['ArcComm'] = ''
    v['ArcTourCode'] = ''
    v['TicketDesignator'] = ''
    v['ArcCommUpdated'] = ''
    v['ArcTourCodeUpdated'] = ''
    v['ArcId'] = i.IarId
    v['Status'] = 0
    if i.PaymentType != 'C':
        v['Status'] = 3
    list_data.append(v)


def run(user_name, datas):
    # ----------------------login
    logger.debug("Run:" + user_name)
    password = conf.get("geoff", user_name)
    login_html = arc_model.login(user_name, password)
    if login_html.find('You are already logged into My ARC') < 0 and login_html.find('Account Settings :') < 0:
        logger.error('login error: '+user_name)
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
            # print data
            # if data['Status'] != 0 and data['Status'] != 3:
            #     continue
            # execute(data, action, token, from_date, to_date)

            check(data, action, token, from_date, to_date)
            # if data['Status'] == 1:
            #     check(data, action, token, from_date, to_date)
    except Exception as ex:
        # print e
        logger.critical(ex)
    # finally:
    # 	arc_model.store(list_data_sql)

    arc_model.iar_logout(ped, action, arcNumber)
    arc_model.logout()


try:
    section = "arc"
    for option in conf.options(section):
        logger.debug(option)
        arc_numbers = conf.get(section, option).split(',')
        list_data_account = filter(lambda x: x['ArcNumber'] in arc_numbers, list_data)
        if not list_data_account:
            continue
        account_id = "gttqc02"
        if option == "all":
            account_id = "gttqc02"
        else:
            break
            # account_id = account_id + option
        run(account_id, list_data_account)
except Exception as e:
    logger.critical(e)
finally:
    arc_model.store(list_data)

# try:
#     update(list_data)
# except Exception as e:
#     logger.critical(e)

try:
    insert(list_data)
except Exception as e:
    logger.critical(e)


#-----------------export excel
file_name = "iar_update_commission_by_user_sync"
try:
    arc_model.exportExcel(list_data, file_name)
except Exception as e:
    logger.critical(e)


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
		</tr>''' % (i['ArcNumber'], i['TicketNumber'],i['IssueDate'], i['Comm'], i['QCComm'], i['ArcCommUpdated']
                                          , i['TourCode'], i['QCTourCode'], i['ArcTourCodeUpdated'], updated)

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

    # mail = arc.Email(smtp_server=mail_smtp_server)
    mail.send(mail_from_addr, mail_to_addr, mail_subject, body, ['excel/' + file_name + '.xlsx'])
    logger.info('email sent')
except Exception as e:
    logger.critical(e)

logger.debug('--------------<<<END>>>--------------')

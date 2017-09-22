import arc
import ConfigParser
import sys
import datetime


def convert_to_float(s):
    f = 0
    try:
        f = float(s)
    except Exception as ex:
        logger.warn(ex)

    return f


def run(user_name, datas):
    # ----------------------login
    logger.debug("Run:" + user_name)
    password = conf.get("login", user_name)
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
            if data['Status'] != 0 and data['Status'] != 3:
                continue
            execute(data, action, token, from_date, to_date)
            if data['Status'] == 1:
                check(data, action, token, from_date, to_date)
    except Exception as ex:
        # print e
        logger.critical(ex)
    # finally:
    # 	arc_model.store(list_data_sql)

    arc_model.iar_logout(ped, action, arcNumber)
    arc_model.logout()


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

    is_check_payment = False
    if post['Status'] == 3:
        is_check_payment = True

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
    is_et_button = False
    agent_code = "MJ"
    logger.info("ARC commission: %s  updating commission: %s" % (arc_commission, commission))
    arc_commission_float = convert_to_float(arc_commission)
    commission_float = convert_to_float(commission)
    if post['OPUser'] != "GW" and post['updatedByRole'] != "AG" and arc_commission_float > commission_float:
        agent_code = "M2"
        is_et_button = True
        commission = arc_commission

    financialDetails_html = arc_model.financialDetails(token, is_check_payment, commission, waiverCode, maskedFC, seqNum,
                                                       documentNumber, tour_code, qc_tour_code, certificates, agent_code,
                                                       agent_codes, is_et_button, is_check_update=False)
    if not financialDetails_html:
        return

    if not is_et_button:
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
    else:
        token = arc_regex.itineraryEndorsements(financialDetails_html)

    if token:
        transactionConfirmation_html = arc_model.transactionConfirmation(token)
        if transactionConfirmation_html:
            if transactionConfirmation_html.find('Document has been modified') >= 0:
                post['Status'] = 1
            else:
                logger.warning('update may be error')


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

    voided_index = modify_html.find('Document is being displayed as view only')
    if voided_index >= 0:
        post['Status'] = 2
        return

    token, maskedFC, arc_commission, waiverCode, certificates = arc_regex.modifyTran(modify_html)
    if not token:
        return

    post['ArcCommUpdated'] = arc_commission
    agent_code = "MJ"
    logger.info("ARC commission: %s  updating commission: %s" % (arc_commission, commission))
    arc_commission_float = convert_to_float(arc_commission)
    commission_float = convert_to_float(commission)
    if arc_commission_float > commission_float:
        agent_code = "M2"
        post['ArcCommUpdated'] = commission

    financialDetails_html = arc_model.financialDetails(token, is_check_payment, arc_commission, waiverCode, maskedFC,
                                                       seqNum, documentNumber, tour_code, qc_tour_code, certificates,
                                                       agent_code, agent_codes, is_check_update=True)
    if not financialDetails_html:
        return

    token, arc_tour_code, backOfficeRemarks, ticketDesignators = arc_regex.financialDetails(financialDetails_html)
    if not token:
        return

    post['ArcTourCodeUpdated'] = arc_tour_code


def update(datas):
    ids = []
    for data in datas:
        if data['Status'] != 0:
            if data['QcId'] not in ids:
                ids.append("'" + data['QcId'] + "'")
    if ids:
        update_sql = "update TicketQC set ARCupdated=1 where Id in (%s)" % ','.join(ids)
        if ms.ExecNonQuery(update_sql) > 0:
            logger.info('update sql success')
        else:
            logger.error('update sql error')


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
        data['QCTourCode'] = "" if data['QCTourCode'] is None else data['QCTourCode'].upper()
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


# def sample(post, action, token, from_date, to_date):
#     arcNumber = post['ArcNumber']
#     documentNumber = post['Ticket']
#     date = post['IssueDate']
#     commission = post['QCComm']
#     tour_code = ""
#     if post['TourCode']:
#         tour_code = post['TourCode']
#
#     qc_tour_code = ""
#     if post['QCTourCode']:
#         qc_tour_code = post['QCTourCode']
#
#     is_check_payment = False
#     if post['Status'] == 3:
#         is_check_payment = True
#
#     date_time = datetime.datetime.strptime(date, '%Y-%m-%d')
#     ped = (date_time + datetime.timedelta(days=(6 - date_time.weekday()))).strftime('%d%b%y').upper()
#     logger.info("UPDATING PED: " + ped + " arc: " + arcNumber + " tkt: " + documentNumber)
#     search_html = arc_model.search(ped, action, arcNumber, token, from_date, to_date, documentNumber)
#     if not search_html:
#         return
#
#     seqNum, documentNumber = arc_regex.search(search_html)
#     if not seqNum:
#         return
#
#     modify_html = arc_model.modifyTran(seqNum, documentNumber)
#     if not modify_html:
#         return
#
#     voided_index = modify_html.find('Document is being displayed as view only')
#     if voided_index >= 0:
#         post['Status'] = 2
#         return
#
#     token, maskedFC, arc_commission, waiverCode, certificates = arc_regex.modifyTran(modify_html)
#     if not token:
#         return
#
#     post['ArcComm'] = arc_commission
#     is_et_button = False
#     agent_code = "MJ"
#     if float(arc_commission) > float(commission):
#         agent_code = "M2"
#         is_et_button = True
#         commission = arc_commission
#
#     financialDetails_html = arc_model.financialDetails(token, is_check_payment, commission, waiverCode, maskedFC,
#                                                        seqNum,
#                                                        documentNumber, tour_code, qc_tour_code, certificates,
#                                                        agent_code,
#                                                        agent_codes, is_et_button, is_check_update=False)
#     if not financialDetails_html:
#         return


arc_model = arc.ArcModel("arc update commission by status")
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

sql = ('''
declare @t date
set @t=dateadd(day,-7,getdate())
select aa.*,iu.Id IarId from (
--select t.Id,qc.Id qcId,t.TicketNumber,substring(t.TicketNumber,4,10) Ticket,t.IssueDate,t.ArcNumber,t.PaymentType,
--t.Comm,t.TourCode,t.QCComm UpdatedComm,t.QCTourCode UpdatedTourCode from Ticket t
--left join TicketQC qc
--on t.Id=qc.TicketId
--where qc.ARCupdated=0
--and t.QCStatus=2 and qc.AGStatus=1 and qc.OPStatus<>2
--and t.CreateDate>=@t
--union
select t.Id,qc.Id qcId,t.TicketNumber,substring(t.TicketNumber,4,10) Ticket,t.IssueDate,t.ArcNumber,t.PaymentType,
t.Comm,t.TourCode,qc.AGComm UpdatedComm,qc.AGTourCode UpdatedTourCode,qc.OPUser,qc.OPLastUser,
'AG' updatedByRole from Ticket t
left join TicketQC qc
on t.Id=qc.TicketId
where qc.ARCupdated=0
--and t.QCStatus=2 
and qc.AGStatus=3
and ((t.Comm<>qc.AGComm) or (t.TourCode<>qc.AGTourCode))
--and qc.OPStatus<>2
and t.CreateDate>=@t
union
select t.Id,qc.Id qcId,t.TicketNumber,substring(t.TicketNumber,4,10) Ticket,t.IssueDate,t.ArcNumber,t.PaymentType,
t.Comm,t.TourCode,qc.OPComm UpdatedComm,qc.OPTourCode UpdatedTourCode,qc.OPUser,qc.OPLastUser,
'OP' updatedByRole from Ticket t
left join TicketQC qc
on t.Id=qc.TicketId
where qc.ARCupdated=0
and qc.OPStatus=2
and ((t.Comm<>qc.OPComm) or (t.TourCode<>qc.OPTourCode))
and t.CreateDate>=@t
) aa
left join IarUpdate iu
on aa.Id=iu.TicketId
order by IssueDate
''')

rows = ms.ExecQuery(sql)
if len(rows) == 0:
    sys.exit(0)

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
    v['QCComm'] = str(i.UpdatedComm)
    v['TourCode'] = i.TourCode
    v['QCTourCode'] = i.UpdatedTourCode
    v['ArcComm'] = ''
    v['ArcTourCode'] = ''
    v['TicketDesignator'] = ''
    v['ArcCommUpdated'] = ''
    v['ArcTourCodeUpdated'] = ''
    v['ArcId'] = i.IarId
    v['Status'] = 0
    if i.PaymentType != 'C':
        v['Status'] = 3

    op_user = ""
    if i.OPUser:
        op_user = str(i.OPUser).strip().upper()

    if op_user != "GW" and i.OPLastUser:
        op_user = str(i.OPLastUser).strip().upper()

    v['OPUser'] = op_user
    # v['AGStatus'] = i.AGStatus
    v['updatedByRole'] = i.updatedByRole
    list_data.append(v)


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
            account_id = "mulingpeng"
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


#-----------------export excel
file_name = "iar_update_commission_by_status"
try:
    arc_model.exportExcel(list_data, file_name)
except Exception as e:
    logger.critical(e)

mail_smtp_server = conf.get("email", "smtp_server")
mail_from_addr = conf.get("email", "from")
mail_to_addr = conf.get("email", "to_update_status").split(';')
mail_subject = conf.get("email", "subject") + " by status"
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
    # mail.send(mail_from_addr, mail_to_addr, mail_subject, body)
    logger.info('email sent')
except Exception as e:
    logger.critical(e)

logger.debug('--------------<<<END>>>--------------')

import arc
import ConfigParser
import sys
import datetime


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

    # search_html = arc_model.search(ped, action, arcNumber, token, from_date, to_date, documentNumber)
    search_html = arc_model.create_list(token=token, ped=ped, action=action, arcNumber=arcNumber,
                                        viewFromDate=from_date, viewToDate=to_date, documentNumber=documentNumber)
    if not search_html:
        logger.debug("search html None")
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

    token, maskedFC, arc_commission, waiverCode, certificates = arc_regex.modifyTran(modify_html)
    logger.debug("IAR COMM: %s." % arc_commission)
    if not token:
        return

    if arc_commission == "":
        arc_commission = "0"

    post['ArcCommUpdated'] = arc_commission

    financialDetails_html = arc_model.financialDetails(token, is_check_payment, commission, waiverCode, maskedFC,
                                                       seqNum, documentNumber, tour_code, qc_tour_code, certificates,
                                                       "MJ", agent_codes, is_check_update=True)

    if not financialDetails_html:
        return

    token, arc_tour_code, backOfficeRemarks, ticketDesignators = arc_regex.financialDetails(financialDetails_html)
    logger.debug("IAR TOUR CODE: %s" % arc_tour_code)
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
            sqls.append('''insert into IarUpdate(Id,Commission,TourCode,TicketDesignator,IsUpdated,TicketId,channel,syncTimes) values (newid(),%s,'%s','%s',%d,'%s',3,1);''' % (
                comm, data['ArcTourCodeUpdated'], data['TicketDesignator'], is_updated, data['Id']))
        else:
            sqls.append('''update IarUpdate set Commission=%s,TourCode='%s',TicketDesignator='%s',IsUpdated=%d,channel=5, updateDateTime=GETDATE(),syncTimes=syncTimes+1 where Id='%s';''' % (
                comm, data['ArcTourCodeUpdated'], data['TicketDesignator'], is_updated, data['ArcId']))

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


def run(section, user_name, datas):
    logger.debug("RUN: %s" % datas)
    # ----------------------login
    logger.debug(user_name)
    # password = conf.get("geoff", user_name)
    password = conf.get(section, user_name)
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
            logger.debug(data)
            check(data, action, token, from_date, to_date)
    except Exception as ex:
        logger.critical(ex)

    arc_model.iar_logout(ped, action, arcNumber)
    arc_model.logout()

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

mail_is_local = conf.get("email", "is_local").lower() == "true"
mail_smtp_server = conf.get("email", "smtp_server")
mail_smtp_port = conf.get("email", "smtp_port")
mail_is_enable_ssl = conf.get("email", "is_enable_ssl").lower() == "true"
mail_user = conf.get("email", "user")
mail_password = conf.get("email", "password")
mail_from_addr = conf.get("email", "from")
mail_to_addr = conf.get("email", "to_update_user_sync").split(';')
mail_subject = conf.get("email", "subject") + " by user sync"


mail = arc.Email(is_local=mail_is_local, smtp_server=mail_smtp_server, smtp_port=mail_smtp_port, is_enable_ssl=mail_is_enable_ssl,
                     user=mail_user, password=mail_password)

sql = ('''select t.Id,qc.Id qcId,t.TicketNumber,substring(t.TicketNumber,4,10) Ticket,t.IssueDate,t.ArcNumber,t.PaymentType,
t.Comm,t.TourCode,qc.AGComm UpdateComm,qc.AGTourCode UpdateTourCode,iar.Id IarId from Ticket t
left join TicketQC qc
on t.Id=qc.TicketId
left join IarUpdate iar
on t.Id=iar.TicketId
where qc.AGStatus=3
and ISNULL(qc.AGDate,'1900-1-1') >= ISNULL(qc.OPDate,'1900-1-1')
and (iar.Commission is null or iar.Commission<>qc.AGComm or ISNULL(iar.TourCode,'')<>ISNULL(qc.AGTourCode,''))
and t.IssueDate>=CAST(DATEADD(DAY,-3,GETDATE()) AS DATE)
and t.IssueDate<CAST(GETDATE() AS DATE)
and (iar.AuditorStatus is null or iar.AuditorStatus=0)
and iar.syncTimes<3
union
select t.Id,qc.Id qcId,t.TicketNumber,substring(t.TicketNumber,4,10) Ticket,t.IssueDate,t.ArcNumber,t.PaymentType,
t.Comm,t.TourCode,qc.OPComm UpdateComm,qc.OPTourCode UpdateTourCode,iar.Id IarId from Ticket t
left join TicketQC qc
on t.Id=qc.TicketId
left join IarUpdate iar
on t.Id=iar.TicketId
where qc.OPStatus in (1, 2, 15)
and qc.OPComm is not null
and (qc.AGStatus<>3 or (qc.AGStatus=3 and ISNULL(qc.OPDate,'1900-1-1') >= ISNULL(qc.AGDate,'1900-1-1')))
and (iar.Commission is null or iar.Commission<>qc.OPComm or ISNULL(iar.TourCode,'')<>ISNULL(qc.OPTourCode,''))
and (iar.AuditorStatus is null or iar.AuditorStatus=0)
and iar.syncTimes<3
and t.IssueDate>=CAST(DATEADD(DAY,-3,GETDATE()) AS DATE)
and t.IssueDate<CAST(GETDATE() AS DATE)
order by t.ArcNumber
''')

rows = ms.ExecQuery(sql)

if len(rows) == 0:
    mail.send(mail_from_addr, mail_to_addr, mail_subject, "No data,please confirm.")
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

# logger.debug(list_data)

try:
    sectionsToOptions = []
    sections = ["geoff", "login"]
    for section in sections:
        for option in conf.options(section):
            if section == "login" and conf.get("idsToArcs", option) == "all":
                continue
            sectionsToOptions.append({"section": section, "option": option})

    for item in sectionsToOptions:
        arc_name = conf.get("idsToArcs", item["option"])
        arc_numbers = conf.get("arc", arc_name).split(',')
        logger.debug("arc numbers conf: %s" % arc_numbers)
        list_data_account = filter(lambda x: x['ArcNumber'] in arc_numbers, list_data)
        if not list_data_account:
            continue

        run(item["section"], item["option"], list_data_account)
    # section = "arc"
    # for option in conf.options(section):
    #     logger.debug(option)
    #     arc_numbers = conf.get(section, option).split(',')
    #     logger.debug("arc numbers conf: %s" % arc_numbers)
    #     list_data_account = filter(lambda x: x['ArcNumber'] in arc_numbers, list_data)
    #     # logger.debug("lambda: %s" % list_data_account)
    #     if not list_data_account:
    #         continue
    #
    #     account_id = ""
    #     login_section = ""
    #     if option == "all":
    #         account_id = "gttqc02"
    #         login_section = "geoff"
    #     elif option == "aca":
    #         account_id = "muling-aca"
    #         login_section = "login"
    #     elif option == "yww":
    #         account_id = "muling-yww"
    #         login_section = "login"
    #     elif option == "tvo":
    #         account_id = "muling-tvo"
    #         login_section = "login"
    #
    #     run(login_section, account_id, list_data_account)
except Exception as e:
    logger.critical(e)
finally:
    arc_model.store(list_data)

try:
    insert(list_data)
except Exception as e:
    logger.critical(e)

# -----------------export excel
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

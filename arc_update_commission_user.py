import arc
import ConfigParser
import sys
import datetime


def convert_to_float(s):
    is_exception = False
    f = 0
    try:
        if s == "":
            s = "0"

        f = float(s)
    except Exception as ex:
        logger.warn(ex)
        is_exception = True

    return is_exception, f


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
    # if post['Status'] == 3:
    #     is_check_payment = True

    is_et_button = False
    if post['FareType'] and (post['FareType'] == "BULK" or post['FareType'] == "SR") and not post['QCTourCode']:
        is_et_button = True

    if tour_code == qc_tour_code:
        is_et_button = True

    date_time = datetime.datetime.strptime(date, '%Y-%m-%d')
    ped = (date_time + datetime.timedelta(days=(6 - date_time.weekday()))).strftime('%d%b%y').upper()
    logger.info("UPDATING PED: " + ped + " arc: " + arcNumber + " tkt: " + documentNumber)
    # search_html = arc_model.search(ped, action, arcNumber, token, from_date, to_date, documentNumber)
    search_html = arc_model.create_list(token=token, ped=ped, action=action, arcNumber=arcNumber,
                                        viewFromDate=from_date, viewToDate=to_date, documentNumber=documentNumber)
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
    if not token:
        return

    post['ArcComm'] = arc_commission
    agent_code = "QC-" + post["updatedByRole"]
    if agent_code == "QC-OPA":
        logger.info("OP Agree ARC : %s, TKT: %s." % (arcNumber, documentNumber))
        agent_code = "QC-OP"

    logger.info("ARC commission: %s  updating commission: %s" % (arc_commission, commission))
    arc_commission_is_exception, arc_commission_float = convert_to_float(arc_commission)
    commission_is_exception, commission_float = convert_to_float(commission)
    if arc_commission_is_exception or commission_is_exception or (post['updatedByRole'] != "AG" and arc_commission_float > commission_float):
        if arc_commission_is_exception or commission_is_exception:
            agent_code = "QC-FAIL"
        elif post['AGStatus'] == 0:
            agent_code = "AG-PENDING"
        elif post['AGStatus'] == 1:
            agent_code = "AG-AGREE"

        is_et_button = True
        if arc_commission:
            logger.debug(modify_html)
            commission = arc_commission
        else:
            commission = post["Comm"]

        post['isPutError'] = True
    # elif post['updatedByRole'] != "AG" and arc_commission_float == 0:
    #     agent_code = "OP"
    post['errorCode'] = agent_code
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

        if not is_et_button:
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
    # if post['Status'] == 3:
    #     is_check_payment = True

    date_time = datetime.datetime.strptime(date, '%Y-%m-%d')
    ped = (date_time + datetime.timedelta(days=(6 - date_time.weekday()))).strftime('%d%b%y').upper()
    logger.info("CHECK PED: " + ped + " arc: " + arcNumber + " tkt: " + documentNumber)
    # search_html = arc_model.search(ped, action, arcNumber, token, from_date, to_date, documentNumber)
    search_html = arc_model.create_list(token=token, ped=ped, action=action, arcNumber=arcNumber,
                                        viewFromDate=from_date, viewToDate=to_date, documentNumber=documentNumber)
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

    token, maskedFC, arc_commission, waiverCode, certificates = arc_regex.modifyTran(modify_html)
    if not token:
        return

    post['ArcCommUpdated'] = arc_commission
    agent_code = "QC-" + post["updatedByRole"]
    logger.info("ARC commission: %s  updating commission: %s" % (arc_commission, commission))
    # arc_commission_float = convert_to_float(arc_commission)
    # commission_float = convert_to_float(commission)
    if post['isPutError']:
        logger.debug("put error")
        if ("QC-FAIL" in certificates) or ("AG-PENDING" in certificates) or ("AG-AGREE" in certificates):
            post['hasPutError'] = True
        post['ArcTourCodeUpdated'] = post['ArcTourCode']
        return

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

        comm = None
        try:
            comm = float(data['ArcCommUpdated'])
        except ValueError:
            comm = "null"

        if not data['ArcId']:
            sqls.append('''insert into IarUpdate(Id,Commission,TourCode,TicketDesignator,IsUpdated,TicketId,channel,errorCode) values (newid(),%s,'%s','%s',%d,'%s',3,'%s');''' % (
                comm, data['ArcTourCodeUpdated'], data['TicketDesignator'], is_updated, data['Id'], data['errorCode']))
        else:
            run_time = ""
            if data['Status'] != 0:
                run_time = ",runTimes=ISNULL(runTimes,0)+1"

            sqls.append('''update IarUpdate set Commission=%s,TourCode='%s',TicketDesignator='%s',IsUpdated=%d,channel=3,updateDateTime=GETDATE()%s,errorCode='%s' where Id='%s';''' % (
                comm, data['ArcTourCodeUpdated'], data['TicketDesignator'], is_updated, run_time, data['errorCode'], data['ArcId']))

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


arc_model = arc.ArcModel("arc update commission by user")
arc_regex = arc.Regex()
logger = arc_model.logger

logger.debug('--------------<<<START>>>--------------')
logger.debug('select sql')

conf = ConfigParser.ConfigParser()
conf.read('../iar_update.conf')
agent_codes = conf.get("certificate", "agentCodes").split(',')
gw_level_agent_codes = filter(None, conf.get("gw_level", "agentCodes").split(','))
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
mail_to_addr = conf.get("email", "to_update_user").split(';')
mail_subject = conf.get("email", "subject") + " by user"

mail = arc.Email(is_local=mail_is_local, smtp_server=mail_smtp_server, smtp_port=mail_smtp_port, is_enable_ssl=mail_is_enable_ssl,
                     user=mail_user, password=mail_password)

op_users_sql = ""
if gw_level_agent_codes:
    op_users_where = ",".join([("'" + x + "'") for x in gw_level_agent_codes])
    op_users_sql = " and (qc.OPUser in (" + op_users_where + ") or OPLastUser in (" + op_users_where + ")) "

sql = ('''
declare @start date
declare @end date
set @start=dateadd(day,-3,getdate())
set @end=getdate()
select t.Id,qc.Id qcId,t.TicketNumber,substring(t.TicketNumber,4,10) Ticket,t.IssueDate,t.ArcNumber,t.PaymentType,t.FareType,
t.Comm,t.TourCode,qc.AGComm UpdateComm,qc.AGTourCode UpdateTourCode,qc.AGStatus,'AG' updatedByRole,iar.Id IarId from Ticket t
left join TicketQC qc
on t.Id=qc.TicketId
left join IarUpdate iar
on t.Id=iar.TicketId
where (qc.ARCupdated=0 or (qc.ARCupdated=1 and iar.IsUpdated=0 and iar.runTimes=0))
and (iar.AuditorStatus is null or iar.AuditorStatus=0)
and qc.AGStatus=3
--and qc.OPStatus<>2
--and (t.Comm<>qc.AGComm or t.TourCode<>qc.AGTourCode)
and (iar.Commission is null or iar.Commission<>qc.AGComm or iar.TourCode<>qc.AGTourCode)
and t.IssueDate>=@start
and t.IssueDate<@end
and ISNULL(qc.AGDate,'1900-1-1') >= ISNULL(qc.OPDate,'1900-1-1')
union
select t.Id,qc.Id qcId,t.TicketNumber,substring(t.TicketNumber,4,10) Ticket,t.IssueDate,t.ArcNumber,t.PaymentType,t.FareType,
t.Comm,t.TourCode,qc.OPComm UpdateComm,qc.OPTourCode UpdateTourCode,qc.AGStatus,'OP' updatedByRole,iar.Id IarId from Ticket t
left join TicketQC qc
on t.Id=qc.TicketId
left join IarUpdate iar
on t.Id=iar.TicketId
where (qc.ARCupdated=0 or (qc.ARCupdated=1 and iar.IsUpdated=0 and iar.runTimes=0))
and (iar.AuditorStatus is null or iar.AuditorStatus=0)
and qc.OPStatus=2
and (qc.AGStatus<>3 or (qc.AGStatus=3 and ISNULL(qc.OPDate,'1900-1-1') >= ISNULL(qc.AGDate,'1900-1-1')))
and (t.Comm<>qc.OPComm or t.TourCode<>qc.OPTourCode)
and (iar.Commission is null or iar.Commission<>qc.OPComm or iar.TourCode<>qc.OPTourCode)''' + op_users_sql + '''
and t.IssueDate>=@start
and t.IssueDate<@end
union
select t.Id,qc.Id qcId,t.TicketNumber,substring(t.TicketNumber,4,10) Ticket,t.IssueDate,t.ArcNumber,t.PaymentType,t.FareType,
t.Comm,t.TourCode,qc.OPComm UpdateComm,qc.OPTourCode UpdateTourCode,qc.AGStatus,'OPA' updatedByRole,iar.Id IarId from Ticket t
left join TicketQC qc
on t.Id=qc.TicketId
left join IarUpdate iar
on t.Id=iar.TicketId
where (qc.ARCupdated=0 or (qc.ARCupdated=1 and iar.IsUpdated=0 and iar.runTimes=0))
and qc.OPStatus in (1, 15)
and qc.OPComm is not null
and (qc.AGStatus<>3 or (qc.AGStatus=3 and ISNULL(qc.OPDate,'1900-1-1') >= ISNULL(qc.AGDate,'1900-1-1')))
and (qc.OPComm<> t.Comm or ISNULL(qc.OPTourCode,'')<>ISNULL(t.TourCode,''))
and (iar.AuditorStatus is null or iar.AuditorStatus=0)
and (iar.Commission is null or iar.Commission<>qc.OPComm or ISNULL(iar.TourCode,'')<>ISNULL(qc.OPTourCode,''))''' + op_users_sql + '''
and t.IssueDate>=@start
and t.IssueDate<@end
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
    logger.debug("TKT#: %s" % v['TicketNumber'])
    if v['TicketNumber'] and len(v['TicketNumber']) > 3 and v["TicketNumber"][0:3] == "890":
        logger.info("THIS IS 890, TKT#: %s " % v["TicketNumber"])
        continue
    v['Ticket'] = i.Ticket
    v['IssueDate'] = str(i.IssueDate)
    v['ArcNumber'] = i.ArcNumber
    v['Comm'] = str(i.Comm)
    v['QCComm'] = ""
    if i.UpdateComm is not None:
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

    v['updatedByRole'] = i.updatedByRole
    v['AGStatus'] = i.AGStatus
    v['isPutError'] = False
    v['hasPutError'] = False
    v['FareType'] = i.FareType
    v['errorCode'] = ""
    list_data.append(v)


def run(section, user_name, datas):
    # ----------------------login
    logger.debug(user_name)
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
            # print data
            if data['Status'] != 0 and data['Status'] != 3:
                continue
            execute(data, action, token, from_date, to_date)
            if data['Status'] == 1 or data['Status'] == 3 or data['Status'] == 4:
                check(data, action, token, from_date, to_date)
    except Exception as ex:
        # print e
        logger.critical(ex)
    # finally:
    # 	arc_model.store(list_data_sql)

    arc_model.iar_logout(ped, action, arcNumber)
    arc_model.logout()


try:
    section = "geoff"
    for option in conf.options(section):
        account_id = option
        arc_name = conf.get("idsToArcs", account_id)
        arc_numbers = conf.get("arc", arc_name).split(',')
        list_data_account = filter(lambda x: x['ArcNumber'] in arc_numbers, list_data)
        if not list_data_account:
            continue

        run(section, account_id, list_data_account)

    # section = "arc"
    # for option in conf.options(section):
    #     logger.debug(option)
    #     arc_numbers = conf.get(section, option).split(',')
    #     list_data_account = filter(lambda x: x['ArcNumber'] in arc_numbers, list_data)
    #     if not list_data_account:
    #         continue
    #     account_id = "gttqc02"
    #     if option == "all":
    #         account_id = "gttqc02"
    #     else:
    #         break
    #         # account_id = account_id + option
    #     run(account_id, list_data_account)
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


# -----------------export excel
file_name = "iar_update_commission_by_user"
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

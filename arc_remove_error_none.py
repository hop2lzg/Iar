import arc
import ConfigParser
import datetime
import time


arc_model = arc.ArcModel("arc remove error")
arc_regex = arc.Regex()
logger = arc_model.logger

logger.debug('--------------<<<START>>>--------------')
conf = ConfigParser.ConfigParser()
conf.read('../iar_update.conf')
agent_codes = conf.get("certificate", "agentCodes").split(',')
error_codes = conf.get("error", "errorCodes")

sql_server = conf.get("sql", "server")
sql_database = conf.get("sql", "database")
sql_user = conf.get("sql", "user")
sql_pwd = conf.get("sql", "pwd")
ms = arc.MSSQL(server=sql_server, db=sql_database, user=sql_user, pwd=sql_pwd)


def read_ticket(ticket_number):
    commission = ""
    sql = ('''
    select t.TicketNumber,t.IssueDate,t.Comm,t.QCComm,t.QCStatus,qc.OPStatus,qc.OPComm,qc.AGStatus,qc.AGComm,iar.AuditorStatus,iar.Commission from Ticket t
left join TicketQC qc
on t.Id=qc.TicketId
left join IarUpdate iar
on t.id=iar.TicketId
where TicketNumber like ''' + "'" + ticket_number + '''%'
order by t.IssueDate
    ''')
    rows = ms.ExecQuery(sql)
    if len(rows) == 0:
        return commission

    for row in rows:
        commission = str(row["Comm"])
        auditor_status = row["AuditorStatus"]
        agent_status = row["AGStatus"]
        op_status = row["OPStatus"]
        if auditor_status is not None and auditor_status != 0 and row["Commission"] is not None:
            commission = str(row["Commission"])
        elif agent_status is not None and agent_status != 0:
            if agent_status == 1 and row["OPComm"] is not None:
                commission = str(row["OPComm"])
            elif agent_status == 3 and row["AGComm"] is not None:
                commission = str(row["AGComm"])
        elif op_status is not None and op_status == 2 and row["OPComm"] is not None:
            commission = str(row["OPComm"])

    if not commission:
        commission = "0"

    return commission


def run(user_name, arc_numbers, ped, action):
    # ----------------------login
    logger.debug(user_name)
    password = conf.get("login", user_name)

    if not arc_model.execute_login(user_name, password):
        return

    # -------------------go to IAR
    iar_html = arc_model.iar()
    if not iar_html:
        logger.error('GO TO IAR ERROR')
        return

    last_arc_number = ""
    for arc_number in arc_numbers:
        logger.debug(arc_number)
        last_arc_number = arc_number
        remove(date_time, date_week, ped, action, arc_number)

    arc_model.iar_logout(ped, action, last_arc_number)
    arc_model.logout()


def execute(data):
    seqNum = data['seqNum']
    documentNumber = data['documentNumber']
    logger.debug('ducumentNumber: %s' % documentNumber)
    modify_html = arc_model.modifyTran(seqNum, documentNumber)
    if not modify_html:
        return

    is_void_pass = arc_regex.check_status(modify_html)

    if is_void_pass == 2:
        data['Status'] = 2
        return
    elif is_void_pass == 1:
        data['Status'] = 4
        return

    token, maskedFC, commission, waiverCode, certificates = arc_regex.modifyTran(modify_html)
    logger.debug("Regex commission: %s" % commission)
    commission = read_ticket(data["ticketNumber"])
    time.sleep(2)
    if commission is None:
        logger.debug("ARC COMM IS NONE, TKT.# %s, HTML: %s" % (documentNumber, modify_html))
        return

    if not token or not maskedFC:
        return

    logger.debug("Updating ticket: %s, commission: %s." % (data["ticketNumber"], commission))
    financialDetails_html = arc_model.financialDetails(token, False, commission, waiverCode, maskedFC,
                                                       seqNum, documentNumber, "", "", certificates, "QC-ERROR", agent_codes,
                                                       is_et_button=True)
    if not financialDetails_html:
        return

    token = arc_regex.itineraryEndorsements(financialDetails_html)
    if token:
        transactionConfirmation_html = arc_model.transactionConfirmation(token)
        if transactionConfirmation_html:
            if transactionConfirmation_html.find('Document has been modified') >= 0:
                data['Status'] = 1
            else:
                logger.warning('Document can not modify')


def remove(today, weekday, ped, action, arc_number):
    listTransactions_html = arc_model.listTransactions(ped, action, arc_number)
    if not listTransactions_html:
        logger.error('go to listTransactions_html error')
        return
    token, from_date, to_date = arc_regex.listTransactions(listTransactions_html)
    if not token:
        logger.error('regex listTransactions token error')
        return
    search_html = arc_model.searchError(ped, action, arc_number, token, from_date, to_date)
    if not search_html:
        logger.error('go to seach error')
        return

    list_entry_date = []
    if weekday >= 2:
        list_entry_date.append((today + datetime.timedelta(days=-2)).strftime('%d%b%y').upper())

    entry_date = '\d{2}[A-Z]{3}\d{2}'
    if list_entry_date:
        entry_date = '|'.join(list_entry_date)

    list_regex_search = arc_regex.search_error(search_html, entry_date, "NONE")

    if not list_regex_search:
        logger.warning('Regex search error')
        return

    logger.info(list_regex_search)
    for search in list_regex_search:
        v = {}
        v['ticketNumber'] = search[2] + search[0]
        v['seqNum'] = search[1]
        v['documentNumber'] = search[0]
        v['date'] = search[4]
        v['arcNumber'] = arc_number
        v['Status'] = 0
        # print v
        execute(v)

        list_data.append(v)

list_data = []
try:
    date_time = datetime.datetime.now()
    date_week = date_time.weekday()
    date_ped = date_time + datetime.timedelta(days=(6 - date_time.weekday()))
    if date_week < 2:
        date_ped = date_ped + datetime.timedelta(days=-7)
    # from_date=(date_ped+datetime.timedelta(days = -6)).strftime('%d%b%y').upper()
    ped = date_ped.strftime('%d%b%y').upper()
    action = "7"
    section = "arc"
    for option in conf.options(section):
        logger.debug(option)
        arc_numbers = conf.get(section, option).split(',')
        account_id = "muling-"
        if option == "all":
            account_id = conf.get("accounts", "all")
        else:
            account_id = account_id + option
        run(account_id, arc_numbers, ped, action)
except Exception as e:
    logger.critical(e)

mail_smtp_server = conf.get("email", "smtp_server")
mail_from_addr = conf.get("email", "from")
mail_to_addr = conf.get("email", "to_remove_error").split(';')
mail_subject = conf.get("email", "subject") + " remove error"

try:
    body = ''
    for i in list_data:
        status, updated = arc_model.convertStatus(i, is_remove_error=True)

        body = body + '''<tr>
        <td>%s</td>
        <td>%s</td>
        <td>%s</td>
        <td>%s</td></tr>''' % (i['arcNumber'], i['ticketNumber'], i['date'], status)
    body = '''<table border=1>
    <thead>
        <tr>
            <th>ARC</th>
            <th>TicketNumber</th>
            <th>Date</th>
            <th>Status</th>
        </tr>
    </thead>
    <tbody>%s</tbody>
    </table>''' % body
    mail = arc.Email(smtp_server=mail_smtp_server)
    mail.send(mail_from_addr, mail_to_addr, mail_subject, body)
except Exception as e:
    logger.critical(e)

logger.debug('--------------<<<END>>>--------------')

import arc
import ConfigParser
import datetime


arc_model = arc.ArcModel("arc remove error")
arc_regex = arc.Regex()
logger = arc_model.logger

logger.debug('--------------<<<START>>>--------------')
conf = ConfigParser.ConfigParser()
conf.read('../iar_update.conf')
agent_codes = conf.get("certificate", "agentCodes").split(',')


def run(user_name, arc_numbers, ped, action):
    # ----------------------login
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

    # voided_index = modify_html.find('Document is being displayed as view only')
    # if voided_index >= 0:
    #     data['status'] = 2
    #     return

    token, maskedFC, commission, waiverCode, certificates = arc_regex.modifyTran(modify_html)
    if commission is None:
        logger.debug("ARC COMM IS NONE, TKT.# %s, HTML: %s" % (documentNumber, modify_html))
        return

    logger.debug("regex commission: %s" % commission)
    if not token or not maskedFC:
        return

    financialDetails_html = arc_model.financialDetails(token, False, commission, waiverCode, maskedFC,
                                                       seqNum, documentNumber, "", "", certificates, "", agent_codes,
                                                       is_et_button=True)
    if not financialDetails_html:
        return

    token = arc_regex.itineraryEndorsements(financialDetails_html)
    if token:
        transactionConfirmation_html = arc_model.transactionConfirmation(token)
        if transactionConfirmation_html:
            if transactionConfirmation_html.find('Document has been modified') >= 0:
                data['status'] = 1
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
    # list_entry_date.append((today+datetime.timedelta(days = -1)).strftime('%d%b%y').upper())
    # if weekday==3 or weekday==4:
    # 	list_entry_date.append((today+datetime.timedelta(days = -3)).strftime('%d%b%y').upper())
    # elif weekday==5:
    # 	list_entry_date.append((today+datetime.timedelta(days = -3)).strftime('%d%b%y').upper())
    # 	list_entry_date.append((today+datetime.timedelta(days = -2)).strftime('%d%b%y').upper())

    entry_date = '\d{2}[A-Z]{3}\d{2}'
    if list_entry_date:
        entry_date = '|'.join(list_entry_date)

    list_regex_search = arc_regex.searchError(search_html, entry_date)

    if not list_regex_search:
        logger.warning('regex seach error')
        return

    logger.info(list_regex_search)
    for search in list_regex_search:
        v = {}
        v['ticketNumber'] = search[2] + search[0]
        v['seqNum'] = search[1]
        v['documentNumber'] = search[0]
        v['date'] = search[4]
        v['arcNumber'] = arc_number
        v['status'] = 0
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
            account_id = "mulingpeng"
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
        status, updated = arc_model.convertStatus(i)

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

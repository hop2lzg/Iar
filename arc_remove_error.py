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
error_codes = conf.get("error", "errorCodes")
list_data = []


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
    logger.debug("seqNum: %s, documentNumber: %s.", seqNum, documentNumber)
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
    if commission is None:
        logger.debug("ARC COMM IS NONE, TKT.# %s, HTML: %s" % (documentNumber, modify_html))
        return

    logger.debug("Regex commission: %s" % commission)
    if not token:
        logger.warn("MODIFY TRAN REGEX ERROR.")
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
                data['Status'] = 1
            else:
                logger.warning('Document can not modify')


def remove(today, weekday, ped, action, arc_number):
    list_entry_date = []
    if weekday == 4:
        list_entry_date.append((today + datetime.timedelta(days=-4)).strftime('%d%b%y').upper())
        list_entry_date.append((today + datetime.timedelta(days=-3)).strftime('%d%b%y').upper())
        list_entry_date.append((today + datetime.timedelta(days=-2)).strftime('%d%b%y').upper())

    entry_date = '\d{2}[A-Z]{3}\d{2}'
    if list_entry_date:
        entry_date = '|'.join(list_entry_date)

    listTransactions_html = arc_model.listTransactions(ped, action, arc_number)
    if not listTransactions_html:
        logger.error('go to listTransactions_html error')
        return
    token, from_date, to_date = arc_regex.listTransactions(listTransactions_html)
    if not token:
        logger.error('regex listTransactions token error')
        return

    searchs = []
    for page in range(0, 10):
        is_next_page = False
        if page > 0:
            is_next_page = True

        # create_list_html = arc_model.create_list(token, ped, action, arcNumber=arc_number, selectedStatusId="E",
        #                                          selectedTransactionType="", selectedFormOfPayment="",
        #                                          dateTypeRadioButtons="ped", viewFromDate=from_date,
        #                                          viewToDate=to_date, selectedNumberOfResults="500",
        #                                          isNext=is_next_page, page=page)

        create_list_html = arc_model.create_list(token, ped, action, arcNumber=arc_number, viewFromDate=from_date,
                                                 viewToDate=to_date, documentNumber="", selectedStatusId="E",
                                                 selectedTransactionType="", selectedFormOfPayment="",
                                                 dateTypeRadioButtons="ped", selectedNumberOfResults="500",
                                                 is_next=is_next_page, page=page)

        if not create_list_html:
            logger.error('GO TO CREATE LIST ERROR')
            break

        token = arc_regex.get_token(create_list_html)
        if not token:
            logger.warning("SEARCH ERROR HTML REGEX TOKEN ERROR, AT PAGE: %d", page)
            break
        list_regex_search = arc_regex.search_error(create_list_html, entry_date, error_codes)
        logger.info("Page: %d, Regex errors: %s" % (page, list_regex_search))
        if list_regex_search:
            for search in list_regex_search:
                v = {}
                v['ticketNumber'] = search[2] + search[0]
                v['seqNum'] = search[1]
                v['documentNumber'] = search[0]
                v['date'] = search[4]
                v['arcNumber'] = arc_number
                v['Status'] = 0
                searchs.append(v)
        else:
            logger.warning("Regex search maybe error, at page: %d" % page)

        if create_list_html.find('title="Next Page" alt="Next Page">') >= 0:
            logger.debug("HAS NEXT BUTTON, GO TO NEXT PAGE.")
        else:
            logger.debug("NOT NEXT BUTTON.")
            break

    if searchs:
        for t in searchs:
            execute(t)
            list_data.append(t)

try:
    date_time = datetime.datetime.now()
    date_week = date_time.weekday()
    if date_week == 0:
        error_codes = error_codes + "|QC-RE|AT-ERROR"
    date_ped = date_time + datetime.timedelta(days=(6 - date_time.weekday()))
    if date_week < 2:
        date_ped = date_ped + datetime.timedelta(days=-7)

    ped = date_ped.strftime('%d%b%y').upper()
    action = "7"
    for option in conf.options("login"):
        account_id = option
        logger.debug("ID: %s" % account_id)
        arc_name = conf.get("idsToArcs", account_id)
        logger.debug("ARC NAME: %s" % arc_name)
        arc_numbers = conf.get("arc", arc_name).split(',')
        run(account_id, arc_numbers, ped, action)
except Exception as e:
    logger.critical(e)

mail_is_local = conf.get("email", "is_local").lower() == "true"
mail_smtp_server = conf.get("email", "smtp_server")
mail_smtp_port = conf.get("email", "smtp_port")
mail_is_enable_ssl = conf.get("email", "is_enable_ssl").lower() == "true"
mail_user = conf.get("email", "user")
mail_password = conf.get("email", "password")
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
    mail = arc.Email(is_local=mail_is_local, smtp_server=mail_smtp_server, smtp_port=mail_smtp_port,
                     is_enable_ssl=mail_is_enable_ssl,
                     user=mail_user, password=mail_password)
    mail.send(mail_from_addr, mail_to_addr, mail_subject, body, is_html=True)
except Exception as e:
    logger.critical(e)

logger.debug('--------------<<<END>>>--------------')

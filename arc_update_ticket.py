import arc
import ConfigParser
import datetime


arc_model = arc.ArcModel("ARC UPDATE TICKET")
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


def execute(action, token, from_date, to_date, arc_number, document_number, date):
    logger.info("EXECUTE ARC: %s, DOCUMENT: %s, DATE: %s", arc_number, document_number, date)
    # date_time = datetime.datetime.strptime(date, '%Y-%m-%d')
    # ped = (date_time + datetime.timedelta(days=(6 - date_time.weekday()))).strftime('%d%b%y').upper()
    ped = "17JUN18"
    search_html = arc_model.search(ped, action, arc_number, token, from_date, to_date, document_number)

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
        return
    elif is_void_pass == 1:
        return


def run(section, user_name, data):
    # ----------------------login
    logger.debug(user_name)
    password = conf.get(section, user_name)
    if not arc_model.execute_login(user_name, password):
        return

    # -----------------go to IAR
    iar_html = arc_model.iar()
    if not iar_html:
        logger.error('IAR ERROR')
        return

    ped, action, arc_number = arc_regex.iar(iar_html, False)
    if not action:
        logger.error('REGEX IAR ERROR')
        arc_model.logout()
        return

    list_transactions_html = arc_model.listTransactions(ped, action, arc_number)
    if not list_transactions_html:
        logger.error('LIST TRANSACTIONS ERROR')
        arc_model.iar_logout(ped, action, arc_number)
        arc_model.logout()
        return

    token, from_date, to_date = arc_regex.listTransactions(list_transactions_html)
    if not token:
        logger.error('REGEX LIST TRANSACTIONS ERROR')
        arc_model.iar_logout(ped, action, arc_number)
        arc_model.logout()
        return

    ticket_number = "0277118863336"
    airline_number = ticket_number[0:3]
    document_number = ticket_number[3:13]
    print "Ticket number: %s, AIR: %s, TKT: %s." % (ticket_number, airline_number, document_number)
    search_html = arc_model.search(ped, action, arc_number, token, from_date, to_date, document_number)
    if not search_html:
        return

    seq_num, document_number = arc_regex.search(search_html)
    if not seq_num:
        return

    modify_html = arc_model.modifyTran(seq_num, document_number)
    if not modify_html:
        return

    is_void_pass = arc_regex.check_status(modify_html)
    if is_void_pass == 2:
        return
    elif is_void_pass == 1:
        return

    token, masked_fc, arc_commission, waiver_code, certificates = arc_regex.modifyTran(modify_html)
    if not token:
        logger.warn("MODIFY TRAN REGEX ERROR.")
        return

    if arc_commission is None:
        logger.debug("ARC COMM IS NONE, TKT.# %s, HTML: %s" % (document_number, modify_html))
        return

    add_old_document_html = arc_model.add_old_document(token, airline_number, "0611111111", seq_num, document_number)

    token = arc_regex.get_token(add_old_document_html)
    if not token:
        return

    exchange_input_html = arc_model.exchange_input(token, document_number)
    token = arc_regex.get_token(exchange_input_html)
    if not token:
        return

    exchange_summary_html = arc_model.exchange_summary(token, document_number)
    token = arc_regex.get_token(exchange_summary_html)
    if not token:
        return

    remove_old_document_html = arc_model.remove_old_document(token, arc_commission, masked_fc, document_number)

    # financialDetails_html = arc_model.financialDetails(token, False, arc_commission, waiver_code, masked_fc,
    #                                                    seq_num, document_number, "", "", certificates,
    #                                                    "QC-PAY", agent_codes, False, is_check_update=False)

    arc_model.iar_logout(ped, action, arc_number)
    arc_model.logout()


try:
    data = None
    run("geoff", "gttqc02", data)
except Exception as e:
    print e

logger.debug('--------------<<<END>>>--------------')
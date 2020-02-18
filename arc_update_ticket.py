import arc
import ConfigParser
import sys


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


def execute(ped, action, arc_number, airline_number, document_number, token, from_date, to_date):
    result = {'void': 0, 'update': 0}
    logger.info("EXECUTE ARC: %s, AIR: %s, TKT: %s, DATE: %s", arc_number, airline_number, document_number, ped)
    # search_html = arc_model.search(ped, action, arc_number, token, from_date, to_date, document_number)
    search_html = arc_model.create_list(token=token, ped=ped, action=action, arcNumber=arc_number,
                                        viewFromDate=from_date, viewToDate=to_date, documentNumber=document_number)
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
        result["void"] = 2
        return
    elif is_void_pass == 1:
        result["void"] = 1
        return

    token, masked_fc, arc_commission, waiver_code, certificates = arc_regex.modifyTran(modify_html)
    if not token:
        logger.warn("MODIFY TRAN REGEX ERROR.")
        return

    if arc_commission is None:
        logger.debug("ARC COMM IS NONE, TKT.# %s, HTML: %s" % (document_number, modify_html))
        return

    certificateItems = arc_model.set_certificate_item(certificates, "QC-ERROR", agent_codes)
    add_old_document_html = arc_model.add_old_document(token, airline_number, "0611111111", seq_num, document_number,
                                                       arc_commission, waiver_code, certificateItems, masked_fc)

    token = arc_regex.get_token(add_old_document_html)
    if not token:
        logger.debug("ADD OLD DOCUMENT REGEX ERROR.")
        return

    exchange_input_html = arc_model.exchange_input(token, document_number)
    token = arc_regex.get_token(exchange_input_html)
    if not token:
        logger.debug("EXCHANGE INPUT REGEX ERROR.")
        return

    exchange_summary_html = arc_model.exchange_summary(token, document_number)
    token = arc_regex.get_token(exchange_summary_html)
    if not token:
        logger.debug("EXCHANGE SUMMARY REGEX ERROR.")
        return

    remove_old_document_html = arc_model.remove_old_document(token, arc_commission, document_number, waiver_code,
                                                             certificateItems, masked_fc)

    token = arc_regex.get_token(remove_old_document_html)

    financial_details_html = arc_model.financial_details(token, arc_commission, waiver_code, certificateItems,
                                                         masked_fc,
                                                         is_et_button=True)

    token = arc_regex.get_token(financial_details_html)

    transaction_confirmation_html = arc_model.transactionConfirmation(token)
    if transaction_confirmation_html:
        if transaction_confirmation_html.find('Document has been modified') >= 0:
            result["update"] = 1
            logger.info("UPDATE SUCCESS.")
        else:
            result["update"] = 2
            logger.warning('UPDATE MAY BE ERROR.')

    return result


def run(section, user_name, data, is_this_week = True):
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

    ped, action, arc_number = arc_regex.iar(iar_html, is_this_week)
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

    for ticket in data:
        result = execute(ped, action, ticket["arcNumber"], ticket["airlineNumber"], ticket["ticketNumber"], token,
                         from_date, to_date)
        if result:
            ticket["void"] = result["void"]
            ticket["update"] = result["update"]
    arc_model.iar_logout(ped, action, arc_number)
    arc_model.logout()


def update_sql(data):
    if not data:
        return

    sqls = []
    for ticket in data:
        result = 0
        if ticket["void"] != 0:
            result = ticket["void"] + 2
        elif ticket["update"] != 0:
            result = ticket["update"]
        sqls.append("update IAR.dbo.ticket set resultStatus=%d,runCount=runCount+1,updateDateTime=GETDATE() where id='%s';" %(
            result, ticket['id']
        ))

    if not sqls:
        return

    logger.info("".join(sqls))
    rowcount = ms.ExecNonQuerys(sqls)
    if rowcount != len(sqls):
        logger.warn("update:%s, updated:%s" % (len(sqls), rowcount))

    if rowcount > 0:
        logger.info('update success')
    else:
        logger.error('update error')

ms = arc.MSSQL(server=sql_server, db=sql_database, user=sql_user, pwd=sql_pwd)
sql = ('''
select id,arcNumber,airlineNumber,ticketNumber from IAR.dbo.ticket
where runCount<2
order by arcNumber,airlineNumber
''')

data = []
rows = ms.ExecQuery(sql)
if len(rows) == 0:
    sys.exit(0)

for row in rows:
    v = {}
    v['id'] = row.id
    v['arcNumber'] = row.arcNumber
    v['airlineNumber'] = row.airlineNumber
    v['ticketNumber'] = row.ticketNumber
    v['void'] = 0
    v['update'] = 0
    data.append(v)


try:
    # data.append({'id': "0123456", 'arcNumber': "45668571", 'airlineNumber': "057", 'ticketNumber': "7119026178", 'void': 0, 'update': 0})
    # data.append({'arcNumber': "10522374", 'ticketNumber': "0167156259141"})
    if data:
        is_this_week = True
        ped_text = arc_model.read_file("file", "ped.txt")
        if ped_text and ped_text.upper() == "FALSE":
            is_this_week = False

        section = "geoff"
        for option in conf.options(section):
            account_id = option
            run(section, account_id, data, is_this_week)

    update_sql(data)

except Exception as ex:
    logger.fatal(ex)

logger.debug('--------------<<<END>>>--------------')
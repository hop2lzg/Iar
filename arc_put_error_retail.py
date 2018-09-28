import arc
import ConfigParser
import datetime

arc_model = arc.ArcModel("ARC PUT ERROR RETAIL")
arc_regex = arc.Regex()
logger = arc_model.logger

logger.debug('--------------<<<START>>>--------------')

conf = ConfigParser.ConfigParser()
conf.read('../iar_update.conf')
agent_codes = conf.get("certificate", "agentCodes").split(',')
airline_exceptions = conf.get("retail", "airline_exceptions").split(',')
retail_arc_numbers = conf.get("retail", "arcs").split(',')
coi_arc_numbers = conf.get("retail", "coi").split(',')


def execute(seqNum, documentNumber, error_code):
    logger.debug("seqNum: %s, documentNumber: %s.", seqNum, documentNumber)
    result = {'void': 0, 'update': 0}
    modify_html = arc_model.modifyTran(seqNum, documentNumber)
    if not modify_html:
        return

    is_void_pass = arc_regex.check_status(modify_html)
    if is_void_pass == 2:
        result['void'] = 2
        return
    elif is_void_pass == 1:
        result['void'] = 1
        return

    token, maskedFC, arc_commission, waiverCode, certificates = arc_regex.modifyTran(modify_html)
    logger.debug("REGEX COMMISSION: %s" % arc_commission)
    if not token:
        logger.error('MODIFY TRAN ERROR')
        return

    if arc_commission is None:
        logger.debug("ARC COMM IS NONE, TKT.# %s, HTML: %s" % (documentNumber, modify_html))
        return

    financialDetails_html = arc_model.financialDetails(token, False, arc_commission, waiverCode, maskedFC,
                                                       seqNum, documentNumber, "", "", certificates, error_code,
                                                       agent_codes, is_et_button=True)
    if not financialDetails_html:
        return

    token = arc_regex.itineraryEndorsements(financialDetails_html)
    if token:
        transactionConfirmation_html = arc_model.transactionConfirmation(token)
        if transactionConfirmation_html:
            if transactionConfirmation_html.find('Document has been modified') >= 0:
                result['update'] = 1
            else:
                result['update'] = 2
                logger.warning('update may be error')

    return result


def get_tickets(token, ped, action, arcNumber, viewFromDate, viewToDate, documentNumber, selectedStatusId,
                selectedDocumentType, selectedTransactionType, selectedFormOfPayment, dateTypeRadioButtons,
                selectedNumberOfResults):
    tickets = []
    for i in range(0, 10):
        logger.debug("PAGE: %d.", i)
        is_next_page = False
        if i > 0:
            is_next_page = True

        create_list_html = arc_model.create_list(token, ped, action, arcNumber=arcNumber,
                                                 viewFromDate=viewFromDate, viewToDate=viewToDate, documentNumber=documentNumber,
                                                 selectedStatusId=selectedStatusId, selectedDocumentType=selectedDocumentType,
                                                 selectedTransactionType=selectedTransactionType, selectedFormOfPayment=selectedFormOfPayment,
                                                 dateTypeRadioButtons=dateTypeRadioButtons,
                                                 selectedNumberOfResults=selectedNumberOfResults, is_next=is_next_page, page=i)

        token = arc_regex.get_token(create_list_html)
        if not token:
            logger.error('GO TO CREATE LIST ERROR')
            continue

        modify_trans = arc_regex.modify_trans(create_list_html)
        if modify_trans:
            if arcNumber in retail_arc_numbers:
                for modify_tran in modify_trans:
                    if modify_tran[0] in airline_exceptions or modify_tran[3] == "EX":
                        continue

                    ticket = {"airline": modify_tran[0], "seqNum": modify_tran[1], "documentNumber": modify_tran[2],
                              "transactionType": modify_tran[3], "result": 0}
                    tickets.append(ticket)
            elif arcNumber in coi_arc_numbers:
                if selectedFormOfPayment == "CA":
                    for modify_tran in modify_trans:
                        ticket = {"airline": modify_tran[0], "seqNum": modify_tran[1], "documentNumber": modify_tran[2],
                                  "transactionType": modify_tran[3], "result": 0}
                        tickets.append(ticket)
                elif selectedFormOfPayment == "CC":
                    for modify_tran in modify_trans:
                        if modify_tran[3] != "ET":
                            continue
                        ticket = {"airline": modify_tran[0], "seqNum": modify_tran[1], "documentNumber": modify_tran[2],
                                  "transactionType": modify_tran[3], "result": 0}
                        tickets.append(ticket)

        if create_list_html and create_list_html.find('title="Next Page" alt="Next Page">') >= 0:
            logger.debug("NEXT")
        else:
            break

    return tickets


def run(section, user_name, is_this_week=True):
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

    # today = (datetime.datetime.now() + datetime.timedelta(days=-1)).strftime('%d%b%y').upper()
    today = datetime.datetime.now().strftime('%d%b%y').upper()

    for retail_arc_number in retail_arc_numbers:
        tickets = get_tickets(token, ped=ped, action=action, arcNumber=retail_arc_number, viewFromDate=today, viewToDate=today,
                              documentNumber="", selectedStatusId="", selectedDocumentType="", selectedTransactionType="ET",
                              selectedFormOfPayment="", dateTypeRadioButtons="entryDate", selectedNumberOfResults="500")
        # print tickets
        if tickets:
            for t in tickets:
                result = execute(t['seqNum'], t['documentNumber'], "QC-RE")
                if result:
                    if result['void'] != 0:
                        t['result'] = result['void'] + 2
                    elif result['update'] != 0:
                        t['result'] = result['update']
            arc_model.store(tickets, retail_arc_number)
        else:
            logger.warn("%s NOT FOUND MODIFY TRANS!" % retail_arc_number)

    for coi_arc_number in coi_arc_numbers:
        coi_cash_tickets = get_tickets(token, ped=ped, action=action, arcNumber=coi_arc_number, viewFromDate=today,
                              viewToDate=today,
                              documentNumber="", selectedStatusId="", selectedDocumentType="",
                              selectedTransactionType="ET",
                              selectedFormOfPayment="CA", dateTypeRadioButtons="entryDate", selectedNumberOfResults="500")

        # print coi_cash_tickets
        coi_credit_tickets = get_tickets(token, ped=ped, action=action, arcNumber=coi_arc_number, viewFromDate=today,
                              viewToDate=today,
                              documentNumber="", selectedStatusId="", selectedDocumentType="",
                              selectedTransactionType="ET",
                              selectedFormOfPayment="CC", dateTypeRadioButtons="entryDate", selectedNumberOfResults="500")
        # print coi_credit_tickets
        coi_tickets = coi_cash_tickets + coi_credit_tickets
        # print coi_tickets
        if coi_tickets:
            for t in coi_tickets:
                result = execute(t['seqNum'], t['documentNumber'], "AT-ERROR")
                if result:
                    if result['void'] != 0:
                        t['result'] = result['void'] + 2
                    elif result['update'] != 0:
                        t['result'] = result['update']
            arc_model.store(coi_tickets, coi_arc_number)
        else:
            logger.warn("%s NOT FOUND MODIFY TRANS!" % coi_arc_number)

    arc_model.iar_logout(ped, action, arc_number)
    arc_model.logout()

try:
    run("geoff", "gttqc02", is_this_week=True)
except Exception as ex:
    logger.fatal(ex)
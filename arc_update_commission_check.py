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

    is_check = False

    if post['Status'] == 3:
        is_check = True

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

    financialDetails_html = arc_model.financialDetails(token, is_check, commission, waiverCode, maskedFC, seqNum,
                                                       documentNumber, tour_code, qc_tour_code, certificates, True)
    if not financialDetails_html:
        return

    token, arc_tour_code, backOfficeRemarks, ticketDesignators = arc_regex.financialDetails(financialDetails_html)
    if not token:
        return
    post['ArcTourCodeUpdated'] = arc_tour_code

def run(user_name, datas):
    # ----------------------login
    password = conf.get("login", user_name)
    login_html = arc_model.login(user_name, password)
    if login_html.find('You are already logged into My ARC') < 0 and login_html.find('Account Settings :') < 0:
        logger.error('login error')
        return
    # -----------------go to IAR
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
            check(data, action, token, from_date, to_date)
    except Exception as ex:
        logger.critical(ex)

    arc_model.iar_logout(ped, action, arcNumber)
    arc_model.logout()


arc_model = arc.ArcModel("arc update commission")
arc_regex = arc.Regex()
logger = arc_model.logger

logger.debug('--------------<<<START>>>--------------')
conf = ConfigParser.ConfigParser()
conf.read('../iar_update.conf')

list_data = []
v = {}
v['Ticket'] = ""
v['IssueDate'] = ""
v['ArcNumber'] = ""
v['Comm'] = ""
v['QCComm'] = ""
v['TourCode'] = ""
v['QCTourCode'] = ""
v['ArcComm'] = ''
v['ArcTourCode'] = ''
v['TicketDesignator'] = ''
v['ArcCommUpdated'] = ''
v['ArcTourCodeUpdated'] = ''
v['Status'] = 0


try:
    run("mulingpeng", list_data)
except Exception as e:
    logger.critical(e)


logger.debug('--------------<<<END>>>--------------')

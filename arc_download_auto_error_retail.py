import arc
import ConfigParser
import datetime


conf = ConfigParser.ConfigParser()
conf.read('../iar_update.conf')
arc_model = arc.ArcModel("arc download")
arc_regex = arc.Regex()
logger = arc_model.logger
logger.debug("<<<<<<<<<<<<<<<<<<<<<START>>>>>>>>>>>>>>>>>>>>>")
csv_lines = []


def get_total(ped, action, arc_number, token, view_from_date, view_to_date, document_number,
                     date_type_radio_buttons):
    search_html = arc_model.search(ped, action, arc_number, token, view_from_date, view_to_date, document_number,
                     date_type_radio_buttons)

    if not search_html:
        logger.warn("GO TO SEARCH ERROR.")
        return

    seq_num, document_number = arc_regex.search(search_html)
    if not seq_num:
        logger.warn("SEARCH REGEX ERROR.")
        return

    modify_html = arc_model.modifyTran(seq_num, document_number)
    if not modify_html:
        logger.warn("GO TO MODIFY TRAN ERROR.")
        return

    total = arc_regex.get_total(modify_html)
    if total:
        return total[0]


def run(section, user_name, today, is_this_week=True):
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

    arc_number = "23534803"
    selected_status_id = "E"
    selected_transaction_type = "SA"
    selected_form_of_payment = "CA"
    date_type_radio_buttons = "entryDate"
    view_from_date = today
    view_to_date = today
    create_list_html = arc_model.create_list(token, ped, action, arcNumber=arc_number, selectedStatusId="E",
                                             selectedTransactionType="SA", selectedFormOfPayment="CA",
                                             dateTypeRadioButtons="entryDate", viewFromDate=view_from_date,
                                             viewToDate=view_to_date, selectedNumberOfResults="20")

    token = arc_regex.get_token(create_list_html)
    if not token:
        logger.error('CREATE LIST REGEX ERROR.')

    csv_text = arc_model.get_csv(user_name, is_this_week, ped, action, arc_number, token, view_from_date, view_to_date,
                                 dateTypeRadioButtons=date_type_radio_buttons, selectedStatusId=selected_status_id,
                                 selectedTransactionType=selected_transaction_type, selectedFormOfPayment=selected_form_of_payment)

    if csv_text:
        lines = csv_text.split('\r\n')
        for line in lines:
            if line:
                if lines.index(line) == 0:
                    csv_lines.append(line + ",TOTAL")
                else:
                    cells = line.split(',')
                    if len(cells) == 0:
                        continue

                    document_number = cells[2]
                    incorrect_data = cells[5]
                    if not incorrect_data or incorrect_data != "QC-RE":
                        continue
                    total = get_total(ped, action, arc_number, token, view_from_date, view_to_date, document_number,
                                          date_type_radio_buttons)

                    # print "AIR: %s, TKT: %s, ERR: %s, TTL: %s." % (cells[1], document_number, incorrect_data, total)
                    total_cell = "" if total is None else total
                    csv_lines.append(line + "," + total_cell)

    arc_model.iar_logout(ped, action, arc_number)
    arc_model.logout()

try:
    is_this_week = True
    now = datetime.datetime.now()
    week = now.weekday()
    if week == 0:
        now = now + datetime.timedelta(days=(-1))
        is_this_week = False
    run("geoff", "gttqc02", now.strftime('%d%b%y').upper(), is_this_week=is_this_week)

except Exception as ex:
    logger.fatal(ex)

if csv_lines:
    mail_smtp_server = conf.get("email", "smtp_server")
    mail_from_addr = conf.get("email", "from")
    mail_to_addr = conf.get("email", "to_arc_download_auto_error_retail").split(';')
    mail_subject = 'IAR DOWNLOAD AUTO ERROR RETAIL'
    try:
        body = "\r\n".join(csv_lines)
        folder = "csv"
        file_name = "auto_error_retail.csv"
        arc_model.csv_write(folder, file_name, body)
        mail = arc.Email(smtp_server=mail_smtp_server)
        mail.send(mail_from_addr, mail_to_addr, mail_subject, "Please check the attachment.", [folder + "/" + file_name])
        logger.debug("sent email")
    except Exception as ex:
        logger.critical(ex)
import time, datetime
import ConfigParser
import arc
# import thread
import sys
import threading


class MyThread(threading.Thread):
    def __init__(self, name, is_this_week):
        threading.Thread.__init__(self)
        self.name = name
        self.is_this_week = is_this_week

    def run(self):
        # print "runnuing: " + self.name + '\n'
        thread_lock.acquire()
        try:
            execute(self.name, self.is_this_week)
        except Exception as ex:
            logger.fatal(ex)
        finally:
            thread_lock.release()


def append_csv_line(csv_text, arc_number, csv_lines, iars):
    if csv_text:
        lines = csv_text.split('\r\n')
        for line in lines:
            if line:
                if lines.index(line) == 0:
                    lines.append("ARC," + line)
                else:
                    cells = line.split(',')
                    if len(cells) == 0:
                        continue

                    if cells[3] == "CJ":
                        continue

                    carrier = cells[1]
                    document_number = cells[2]

                    if not carrier and not document_number:
                        continue

                    csv_lines.append(arc_number + "," + line)
                    iars.append({"arcNumber": arc_number, "status": cells[0], "carrier": carrier, "documentNumber": document_number,
                                 "ticketType": cells[3], "FOP": cells[4], "total": cells[5], "comm": cells[6],
                                 "i": cells[7], "net": cells[8], "entryDate": cells[9], "windowDate": cells[10],
                                 "voidDate": cells[11], "esac": cells[12]})


def execute(name, is_this_week):
    # print "start name:%s" % name
    password = conf.get("login", name)
    if not arc_model.execute_login(name, password):
        return

    date_time = datetime.datetime.now().strftime('%d%b%y').upper()

    iar_html = arc_model.iar()
    if not iar_html:
        logger.error('open iar error: '+name)
        return

    ped, action, arcNumber = arc_regex.iar(iar_html, is_this_week)
    if not action:
        logger.error('regex iar error: '+name)
        arc_model.logout()
        return

    try:
        arc_name = conf.get("idsToArcs", name)
        arc_numbers = conf.get("arc", arc_name).split(',')
        if arcNumber in arc_numbers:
            arc_number_index = arc_numbers.index(arcNumber)
            arc_numbers[arc_number_index] = arc_numbers[0]
            arc_numbers[0] = arcNumber
        else:
            arc_numbers.insert(0, arcNumber)

        for arc_number in arc_numbers:
            print "downloading arc:%s name:%s" % (arc_number, name)
            list_transactions_html = arc_model.listTransactions(ped, action, arc_number)
            token, from_date, to_date = arc_regex.listTransactions(list_transactions_html)
            if not token:
                continue

            date_type_radio_buttons = "entryDate"
            create_list_html = arc_model.create_list(token=token, ped=ped, action=action, arcNumber=arc_number,
                                                     viewFromDate=date_time, viewToDate=date_time, documentNumber="",
                                                     selectedStatusId="", selectedDocumentType="",
                                                     selectedTransactionType="RF", selectedFormOfPayment="",
                                                     dateTypeRadioButtons=date_type_radio_buttons, selectedNumberOfResults="20")

            token = arc_regex.create_list(create_list_html)

            refund_csv_text = arc_model.get_csv(arc_name, is_this_week, ped, action, arc_number, token, viewFromDate=date_time, viewToDate=date_time,
                              dateTypeRadioButtons=date_type_radio_buttons, selectedTransactionType="RF", tick="RF")

            append_csv_line(refund_csv_text, arc_number, refund_csv_lines, refund_iars)

            time.sleep(5)

            create_list_html = arc_model.create_list(token=token, ped=ped, action=action, arcNumber=arc_number,
                                                     viewFromDate=date_time, viewToDate=date_time, documentNumber="",
                                                     selectedStatusId="", selectedDocumentType="",
                                                     selectedTransactionType="WV", selectedFormOfPayment="",
                                                     dateTypeRadioButtons=date_type_radio_buttons, selectedNumberOfResults="20")

            token = arc_regex.create_list(create_list_html)

            waiver_code_csv_text = arc_model.get_csv(arc_name, is_this_week, ped, action, arc_number, token, viewFromDate=date_time, viewToDate=date_time,
                              dateTypeRadioButtons="entryDate", selectedTransactionType="WV", tick="WV")

            append_csv_line(waiver_code_csv_text, arc_number, waiver_code_csv_lines, waiver_code_iars)
            # if csv_text:
            #     lines = csv_text.split('\r\n')
            #     for line in lines:
            #         if line:
            #             if lines.index(line) == 0:
            #                 csv_lines.append("ARC," + line)
            #             else:
            #                 cells = line.split(',')
            #                 if len(cells) == 0:
            #                     continue
            #
            #                 carrier = cells[1]
            #                 document_number = cells[2]
            #
            #                 if not carrier and not document_number:
            #                     continue
            #
            #                 csv_lines.append(arc_number + "," + line)
            #                 iars.append({"status": cells[0], "carrier": carrier, "documentNumber": document_number,
            #                             "ticketType": cells[3], "FOP": cells[4], "total": cells[5], "comm": cells[6],
            #                             "i": cells[7], "net": cells[8], "entryDate": cells[9], "windowDate": cells[10],
            #                             "voidDate": cells[11], "esac": cells[12]})
            time.sleep(3)
            # break
    except Exception as ex:
        logger.critical(ex)
    finally:
        arc_model.iar_logout(ped, action, arcNumber)
        arc_model.logout()
    print "over name:%s" % name


def thread_set(is_this_week):
    threads = []
    # thread_lock = threading.Lock()
    for option in conf.options("login"):
        # if option != "muling-aca":
        #     continue

        thread = MyThread(option, is_this_week)
        thread.start()
        threads.append(thread)
        # break
    for t in threads:
        t.join()


conf = ConfigParser.ConfigParser()
conf.read('../iar_update.conf')
arc_model = arc.ArcModel("arc download")
arc_regex = arc.Regex()
logger = arc_model.logger
logger.debug("<<<<<<<<<<<<<<<<<<<<<START>>>>>>>>>>>>>>>>>>>>>")
sql_server = conf.get("sqlRefund", "server")
sql_database = conf.get("sqlRefund", "database")
sql_user = conf.get("sqlRefund", "user")
sql_pwd = conf.get("sqlRefund", "pwd")
ms = arc.MSSQL(server=sql_server, db=sql_database, user=sql_user, pwd=sql_pwd)


thread_lock = threading.Lock()
# csv_lines = []
# iars = []
refund_csv_lines = []
refund_iars = []
waiver_code_csv_lines = []
waiver_code_iars = []

# date_time = datetime.datetime.now()
# date_week = date_time.weekday()
thread_set(True)

def get_arc_number(branchs, branchCode, pcc):
    if branchs:
        for branch in branchs:
            if branch.WorldspanSID == pcc or branch.SabreSID == pcc or branch.AmadeusSID == pcc or branch.ApolloSID == pcc or branch.GalileoSID == pcc:
                return branch.IataNumber

        for branch in branchs:
            if branch.BranchCode == branchCode:
                return branch.IataNumber

    return ""


def export_refund(rows, folder, file_name, branchs):
    if not arc.os.path.exists(folder):
        arc.os.makedirs(folder)
    # book = Workbook()
    # sheet1 = book.active
    # sheet1['A1'] = 'This is A1'
    # sheet1.title = 'sheet1'
    # sheet2 = book.create_sheet(title='sheet2')
    wb = arc.Workbook()
    sheet = wb.active
    sheet['A1'] = 'Branch'
    sheet['B1'] = 'ARC #'
    sheet['C1'] = 'TKT.#'
    sheet['D1'] = 'TKT10'
    sheet['E1'] = 'GDS Status'
    sheet['F1'] = 'IAR Status'
    sheet['G1'] = 'Payment Type'
    sheet['H1'] = 'Waiver Code'
    sheet['I1'] = 'Date'

    row_index = 2
    for r in rows:
        sheet.cell(row=row_index, column=1).value = r.Branch
        sheet.cell(row=row_index, column=2).value = get_arc_number(branchs, r.Branch, r.PCC)
        sheet.cell(row=row_index, column=3).value = r.TicketNumber[0:13]
        sheet.cell(row=row_index, column=4).value = r.TicketNumber[3:13]
        sheet.cell(row=row_index, column=5).value = r.GdsStatus
        sheet.cell(row=row_index, column=6).value = r.IARStatus
        payment_type = ""
        if r.PaymentType == "C":
            payment_type = "CC"
        elif r.PaymentType == "K":
            payment_type = "CK"

        sheet.cell(row=row_index, column=7).value = payment_type
        sheet.cell(row=row_index, column=8).value = r.WaiverCode
        sheet.cell(row=row_index, column=9).value = r.CreateDate.strftime('%m/%d/%Y')
        row_index += 1

    wb.save(filename=folder + "/" + file_name)


def export_csv(csv_lines, folder, file_name):
    body = "\r\n".join(csv_lines)
    arc_model.csv_write(folder, file_name, body)


# def export_vs(rows, iars, folder, file_name, branchs):
#     if not arc.os.path.exists(folder):
#         arc.os.makedirs(folder)
#
#     wb = arc.Workbook()
#     sheet = wb.active
#     sheet['A1'] = 'Branch'
#     sheet['B1'] = 'ARC #'
#     sheet['C1'] = 'TKT.#'
#     sheet['D1'] = 'TKT10'
#     sheet['E1'] = 'GDS Status'
#     sheet['F1'] = 'IAR Status'
#     sheet['G1'] = 'Payment Type'
#     sheet['H1'] = 'Waiver Code'
#     sheet['I1'] = 'Date'
#     sheet['J1'] = 'Result'
#     sheet['K1'] = 'Remark'
#     sheet['L1'] = 'STATUS'
#     sheet['M1'] = 'CARRIER'
#     sheet['N1'] = 'DOCUMENT #'
#     sheet['O1'] = 'TT'
#     sheet['P1'] = 'FOP'
#     sheet['Q1'] = 'TOTAL'
#     sheet['R1'] = 'COMM'
#     sheet['S1'] = 'NET'
#     sheet['T1'] = 'ENTRY DATE'
#
#     row_index = 2
#     for r in rows:
#         sheet.cell(row=row_index, column=1).value = r.Branch
#         sheet.cell(row=row_index, column=2).value = get_arc_number(branchs, r.Branch, r.PCC)
#         sheet.cell(row=row_index, column=3).value = r.TicketNumber[0:13]
#         sheet.cell(row=row_index, column=4).value = r.TicketNumber[3:13]
#         sheet.cell(row=row_index, column=5).value = r.GdsStatus
#         sheet.cell(row=row_index, column=6).value = r.IARStatus
#         payment_type = ""
#         if r.PaymentType == "C":
#             payment_type = "CC"
#         elif r.PaymentType == "K":
#             payment_type = "CK"
#
#         sheet.cell(row=row_index, column=7).value = payment_type
#         sheet.cell(row=row_index, column=8).value = r.WaiverCode
#         sheet.cell(row=row_index, column=9).value = r.CreateDate.strftime('%m/%d/%Y')
#         iar = get_iar(r.TicketNumber)
#         if not iar:
#             sheet.cell(row=row_index, column=10).value = "NOT FOUND"
#         else:
#             sheet.cell(row=row_index, column=11).value = "FOUND"
#             sheet.cell(row=row_index, column=12).value = iar["status"]
#             sheet.cell(row=row_index, column=13).value = iar["carrier"]
#             sheet.cell(row=row_index, column=14).value = iar["documentNumber"]
#             sheet.cell(row=row_index, column=15).value = iar["ticketType"]
#             sheet.cell(row=row_index, column=16).value = iar["FOP"]
#             sheet.cell(row=row_index, column=17).value = iar["total"]
#             sheet.cell(row=row_index, column=18).value = iar["comm"]
#             sheet.cell(row=row_index, column=19).value = iar["net"]
#             sheet.cell(row=row_index, column=20).value = iar["entryDate"]
#
#                             # iars.append({"status": cells[0], "carrier": carrier, "documentNumber": document_number,
#                             #             "ticketType": cells[3], "FOP": cells[4], "total": cells[5], "comm": cells[6],
#                             #             "i": cells[7], "net": cells[8], "entryDate": cells[9], "windowDate": cells[10],
#                             #             "voidDate": cells[11], "esac": cells[12]})
#
#
#         row_index += 1
#
#     wb.save(filename=folder + "/" + file_name)

def get_refund(carrier, documentNumber, rows_list):
    for rows in rows_list:
        for row in rows:
            if row.TicketNumber[0:13] == carrier + documentNumber:
                return row

    return None

def export_vs(iars, rows_list, folder, file_name):
    if not arc.os.path.exists(folder):
        arc.os.makedirs(folder)

    wb = arc.Workbook()
    sheet = wb.active
    sheet['A1'] = 'ARC #'
    sheet['B1'] = 'STATUS'
    sheet['C1'] = 'CARRIER'
    sheet['D1'] = 'DOCUMENT #'
    sheet['E1'] = 'TT'
    sheet['F1'] = 'FOP'
    sheet['G1'] = 'TOTAL'
    sheet['H1'] = 'COMM'
    sheet['I1'] = 'NET'
    sheet['J1'] = 'ENTRY DATE'
    sheet['K1'] = 'Result'
    sheet['L1'] = 'Remark'
    sheet['M1'] = 'Branch'
    sheet['N1'] = 'TKT.#'
    # sheet['D1'] = 'TKT10'
    sheet['O1'] = 'GDS Status'
    sheet['P1'] = 'IAR Status'
    sheet['Q1'] = 'Payment Type'
    sheet['R1'] = 'Waiver Code'
    sheet['S1'] = 'Date'


    row_index = 2
    for iar in iars:
        refund = get_refund(iar["carrier"], iar["documentNumber"], rows_list)
        if refund and not refund.WaiverCode:
            continue

        sheet.cell(row=row_index, column=1).value = iar["arcNumber"]
        sheet.cell(row=row_index, column=2).value = iar["status"]
        sheet.cell(row=row_index, column=3).value = iar["carrier"]
        sheet.cell(row=row_index, column=4).value = iar["documentNumber"]
        sheet.cell(row=row_index, column=5).value = iar["ticketType"]
        sheet.cell(row=row_index, column=6).value = iar["FOP"]
        sheet.cell(row=row_index, column=7).value = iar["total"]
        sheet.cell(row=row_index, column=8).value = iar["comm"]
        sheet.cell(row=row_index, column=9).value = iar["net"]
        sheet.cell(row=row_index, column=10).value = iar["entryDate"]
        if not refund:
            sheet.cell(row=row_index, column=11).value = "NOT FOUND"
            sheet.cell(row=row_index, column=12).value = "Refund only on IAR"
        else:
            sheet.cell(row=row_index, column=11).value = "FOUND"
            if refund.WaiverCode:
                sheet.cell(row=row_index, column=12).value = "Waiver Code missing from IAR"
            sheet.cell(row=row_index, column=13).value = refund.Branch
            sheet.cell(row=row_index, column=14).value = refund.TicketNumber[0:13]
            sheet.cell(row=row_index, column=15).value = refund.GdsStatus
            sheet.cell(row=row_index, column=16).value = refund.IARStatus
            payment_type = ""
            if refund.PaymentType == "C":
                payment_type = "CC"
            elif refund.PaymentType == "K":
                payment_type = "CK"

            sheet.cell(row=row_index, column=17).value = payment_type
            sheet.cell(row=row_index, column=18).value = refund.WaiverCode
            sheet.cell(row=row_index, column=19).value = refund.CreateDate.strftime('%m/%d/%Y')

        row_index += 1

    wb.save(filename=folder + "/" + file_name)


def get_non_waiver_code_refund_iars(refund_iars, waiver_code_iars):
    list = []
    for refund in refund_iars:
        is_found = False
        for waiver_code in waiver_code_iars:
            if refund["carrier"] == waiver_code["carrier"] and refund["documentNumber"] == waiver_code["documentNumber"]:
                is_found = True
                break

        if not is_found:
            list.append(refund)

    return list

# def get_iar(ticket_number):
#     for iar in iars:
#         if iar["carrier"] + iar["documentNumber"] == ticket_number[0:13]:
#             return iar
#
#     return None

# export excel (data)
folder = "file"
refund_file_name = datetime.datetime.now().strftime('refund-%Y%m%d.xlsx')
csv_refund_file_name = datetime.datetime.now().strftime('iar-RF-%Y%m%d.csv')
csv_waiver_code_file_name = datetime.datetime.now().strftime('iar-WV-%Y%m%d.csv')
csv_non_waiver_code_refund_file_name = datetime.datetime.now().strftime('iar-Non-WV-%Y%m%d.csv')
vs_file_name = datetime.datetime.now().strftime('vs-%Y%m%d.xlsx')
# is_write_refund = False
# is_write_csv = False
is_write_vs = False
# try:
#     export_refund(rows, folder, refund_file_name, branchs)
#     is_write_refund = True
# except Exception as e:
#     logger.critical(e)

non_waiver_code_refund_iars = get_non_waiver_code_refund_iars(refund_iars, waiver_code_iars)

try:
    if refund_csv_lines:
        export_csv(refund_csv_lines, folder, csv_refund_file_name)

    if waiver_code_csv_lines:
        export_csv(waiver_code_csv_lines, folder, csv_waiver_code_file_name)

        # is_write_csv = True
except Exception as ex:
    logger.critical(ex)


count = len(non_waiver_code_refund_iars)
loop = 100
length = count/loop + (0 if count % loop == 0 else 1)

refund_rows_list = []
for i in range(0, length):
    end = count if loop * (i + 1) > count else loop * (i + 1)
    sqls = []
    for j in range(i * loop, end):
        sqls.append("TicketNumber like '%s%%'" % (non_waiver_code_refund_iars[j]["carrier"] + non_waiver_code_refund_iars[j]["documentNumber"]))

    refund_sql = "select Branch,PCC,TicketNumber,PaymentType,GdsStatus,IARStatus,WaiverCode,CreateDate from TicketRefund where " + (" or ".join(sqls))

    refund_rows = ms.ExecQuery(refund_sql)
    refund_rows_list.append(refund_rows)

try:
    export_vs(non_waiver_code_refund_iars, refund_rows_list, folder, vs_file_name)
    is_write_vs = True
except Exception as ex:
    logger.critical(ex)


mail_is_local = conf.get("email", "is_local").lower() == "true"
mail_smtp_server = conf.get("email", "smtp_server")
mail_smtp_port = conf.get("email", "smtp_port")
mail_is_enable_ssl = conf.get("email", "is_enable_ssl").lower() == "true"
mail_user = conf.get("email", "user")
mail_password = conf.get("email", "password")
mail_from_addr = conf.get("email", "from")
mail_to_addr = conf.get("email", "to_iar_refund_vs").split(';')
mail_subject = conf.get("email", "subject_iar_refund_vs")


try:
    body = 'Please check this attachments'
    mail = arc.Email(is_local=mail_is_local, smtp_server=mail_smtp_server, smtp_port=mail_smtp_port, is_enable_ssl=mail_is_enable_ssl,
                     user=mail_user, password=mail_password)

    files = []
    # if is_write_refund:
    #     files.append("file/" + refund_file_name)
    #
    # if is_write_csv:
    #     files.append("file/" + csv_file_name)

    if is_write_vs:
        files.append("file/" + vs_file_name)

    mail.send(mail_from_addr, mail_to_addr, mail_subject, body, is_html=True, files=files)
    logger.info('email sent')
except Exception as e:
    logger.critical(e)

logger.debug("<<<<<<<<<<<<<<<<<<<<<END>>>>>>>>>>>>>>>>>>>>>")
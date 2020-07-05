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
            csv_text = arc_model.get_csv(arc_name, is_this_week, ped, action, arc_number, token, viewFromDate=date_time, viewToDate=date_time,
                              dateTypeRadioButtons="entryDate", selectedTransactionType="WV")

            if csv_text:
                lines = csv_text.split('\r\n')
                for line in lines:
                    if line:
                        if lines.index(line) == 0:
                            csv_lines.append("ARC," + line)
                        else:
                            cells = line.split(',')
                            if len(cells) == 0:
                                continue

                            carrier = cells[1]
                            document_number = cells[2]

                            if not carrier and not document_number:
                                continue

                            csv_lines.append(arc_number + "," + line)
                            iars.append({"status": cells[0], "carrier": carrier, "documentNumber": document_number,
                                        "ticketType": cells[3], "FOP": cells[4], "total": cells[5], "comm": cells[6],
                                        "i": cells[7], "net": cells[8], "entryDate": cells[9], "windowDate": cells[10],
                                        "voidDate": cells[11], "esac": cells[12]})
            time.sleep(3)
            # break
    except Exception, ex:
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
sql = ('''
declare @start date
declare @end date
set @start = DATEADD(day,-1,getdate())
set @end = GETDATE()
select Branch,PCC,TicketNumber,PaymentType,GdsStatus,IARStatus,WaiverCode,CreateDate from TicketRefund
where WaiverCode<>'' 
and CreateDate>=@start
and CreateDate<@end
''')

rows = ms.ExecQuery(sql)
if len(rows) == 0:
    logger.warn("NO DATA")
    sys.exit(0)

sql_server_ticketFDA = conf.get("sqlTicketFDA", "server")
sql_database_ticketFDA = conf.get("sqlTicketFDA", "database")
sql_user_ticketFDA = conf.get("sqlTicketFDA", "user")
sql_pwd_ticketFDA = conf.get("sqlTicketFDA", "pwd")
ms_ticketFDA = arc.MSSQL(server=sql_server_ticketFDA, db=sql_database_ticketFDA, user=sql_user_ticketFDA, pwd=sql_pwd_ticketFDA)
sql_ticketFDA = ('''
select BranchCode,IataNumber,WorldspanSID,SabreSID,AmaduesSID AmadeusSID,ApolloSID,GalileoSID from Branch
where IsDelete=0
''')

branchs = ms_ticketFDA.ExecQuery(sql_ticketFDA)


thread_lock = threading.Lock()
csv_lines = []
iars = []
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
    sheet['D1'] = 'GDS Status'
    sheet['E1'] = 'IAR Status'
    sheet['F1'] = 'Payment Type'
    sheet['G1'] = 'Waiver Code'
    sheet['H1'] = 'Date'

    row_index = 2
    for r in rows:
        sheet.cell(row=row_index, column=1).value = r.Branch
        sheet.cell(row=row_index, column=2).value = get_arc_number(branchs, r.Branch, r.PCC)
        sheet.cell(row=row_index, column=3).value = r.TicketNumber[0:13]
        sheet.cell(row=row_index, column=4).value = r.GdsStatus
        sheet.cell(row=row_index, column=5).value = r.IARStatus
        payment_type = ""
        if r.PaymentType == "C":
            payment_type = "CC"
        elif r.PaymentType == "K":
            payment_type = "CK"

        sheet.cell(row=row_index, column=6).value = payment_type
        sheet.cell(row=row_index, column=7).value = r.WaiverCode
        sheet.cell(row=row_index, column=8).value = r.CreateDate
        row_index += 1

    wb.save(filename=folder + "/" + file_name)


def export_csv(csv_lines, folder, file_name):
    body = "\r\n".join(csv_lines)
    arc_model.csv_write(folder, file_name, body)


def export_vs(rows, iars, folder, file_name, branchs):
    if not arc.os.path.exists(folder):
        arc.os.makedirs(folder)

    wb = arc.Workbook()
    sheet = wb.active
    sheet['A1'] = 'Branch'
    sheet['B1'] = 'ARC #'
    sheet['C1'] = 'TKT.#'
    sheet['D1'] = 'GDS Status'
    sheet['E1'] = 'IAR Status'
    sheet['F1'] = 'Payment Type'
    sheet['G1'] = 'Waiver Code'
    sheet['H1'] = 'Date'
    sheet['I1'] = 'Result'
    sheet['J1'] = 'Remark'
    sheet['K1'] = 'STATUS'
    sheet['L1'] = 'CARRIER'
    sheet['M1'] = 'DOCUMENT #'
    sheet['N1'] = 'TT'
    sheet['O1'] = 'FOP'
    sheet['P1'] = 'TOTAL'
    sheet['Q1'] = 'COMM'
    sheet['R1'] = 'NET'
    sheet['S1'] = 'ENTRY DATE'

    row_index = 2
    for r in rows:
        sheet.cell(row=row_index, column=1).value = r.Branch
        sheet.cell(row=row_index, column=2).value = get_arc_number(branchs, r.Branch, r.PCC)
        sheet.cell(row=row_index, column=3).value = r.TicketNumber[0:13]
        sheet.cell(row=row_index, column=4).value = r.GdsStatus
        sheet.cell(row=row_index, column=5).value = r.IARStatus
        payment_type = ""
        if r.PaymentType == "C":
            payment_type = "CC"
        elif r.PaymentType == "K":
            payment_type = "CK"

        sheet.cell(row=row_index, column=6).value = payment_type
        sheet.cell(row=row_index, column=7).value = r.WaiverCode
        sheet.cell(row=row_index, column=8).value = r.CreateDate
        iar = get_iar(r.TicketNumber)
        if not iar:
            sheet.cell(row=row_index, column=9).value = "NOT FOUND"
        else:
            sheet.cell(row=row_index, column=10).value = "FOUND"
            sheet.cell(row=row_index, column=11).value = iar["status"]
            sheet.cell(row=row_index, column=12).value = iar["carrier"]
            sheet.cell(row=row_index, column=13).value = iar["documentNumber"]
            sheet.cell(row=row_index, column=14).value = iar["ticketType"]
            sheet.cell(row=row_index, column=15).value = iar["FOP"]
            sheet.cell(row=row_index, column=16).value = iar["total"]
            sheet.cell(row=row_index, column=17).value = iar["comm"]
            sheet.cell(row=row_index, column=18).value = iar["net"]
            sheet.cell(row=row_index, column=19).value = iar["entryDate"]

                            # iars.append({"status": cells[0], "carrier": carrier, "documentNumber": document_number,
                            #             "ticketType": cells[3], "FOP": cells[4], "total": cells[5], "comm": cells[6],
                            #             "i": cells[7], "net": cells[8], "entryDate": cells[9], "windowDate": cells[10],
                            #             "voidDate": cells[11], "esac": cells[12]})


        row_index += 1

    wb.save(filename=folder + "/" + file_name)


def get_iar(ticket_number):
    for iar in iars:
        if iar["carrier"] + iar["documentNumber"] == ticket_number[0:13]:
            return  iar

    return None

# export excel (data)
folder = "file"
refund_file_name = datetime.datetime.now().strftime('refund-%Y%m%d.xlsx')
csv_file_name = datetime.datetime.now().strftime('iar-%Y%m%d.csv')
vs_file_name = datetime.datetime.now().strftime('vs-%Y%m%d.xlsx')
is_write_refund = False
is_write_csv = False
is_write_vs = False
try:
    export_refund(rows, folder, refund_file_name, branchs)
    is_write_refund = True
except Exception as e:
    logger.critical(e)

try:
    if csv_lines:
        export_csv(csv_lines, folder, csv_file_name)
        is_write_csv = True
except Exception as ex:
    logger.critical(ex)

try:
    export_vs(rows, iars, folder, vs_file_name, branchs)
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
    if is_write_refund:
        files.append("file/" + refund_file_name)

    if is_write_csv:
        files.append("file/" + csv_file_name)

    if is_write_vs:
        files.append("file/" + vs_file_name)

    mail.send(mail_from_addr, mail_to_addr, mail_subject, body, is_html=True, files=files)
    logger.info('email sent')
except Exception as e:
    logger.critical(e)

logger.debug("<<<<<<<<<<<<<<<<<<<<<END>>>>>>>>>>>>>>>>>>>>>")
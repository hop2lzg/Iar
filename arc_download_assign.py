import time, datetime
import ConfigParser
import arc
# import thread
import threading


class MyThread(threading.Thread):
    def __init__(self, name, is_this_week, csv_lines, is_first_arc_number):
        threading.Thread.__init__(self)
        self.name = name
        self.is_this_week = is_this_week
        self.csv_lines = csv_lines
        self.is_first_arc_number = is_first_arc_number

    def run(self):
        # print "runnuing: " + self.name + '\n'
        thread_lock.acquire()
        try:
            execute(self.name, self.is_this_week, self.csv_lines, self.is_first_arc_number)
        except Exception as ex:
            logger.fatal(ex)
        finally:
            thread_lock.release()


def execute(name, is_this_week, csv_lines, is_first_arc_number):
    # print "start name:%s" % name
    password = conf.get("geoff", name)
    if not arc_model.execute_login(name, password):
        return

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
        if name == "gttqc02":
            conf_name = 'all'
        else:
            conf_name = name[-3:]

        arc_numbers = conf.get("arc", conf_name).split(',')
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
            # print "go to list"
            if not token:
                continue

            # print "create list"
            create_list_html = arc_model.create_list(token, ped, action, arc_number)
            token = arc_regex.create_list(create_list_html)

            if not token:
                continue
            csv_text = arc_model.get_csv(name, is_this_week, ped, action, arc_number, token, from_date, to_date,
                              transaction_type="SA", form_of_payment="CA", tick=tick)

            if csv_text:
                lines = csv_text.split('\r\n')
                for line in lines:
                    # logger.debug("ARC: %s, line: %s" % (arc_number, line))
                    if line:
                        # logger.debug("ARC: %s, has line: %s" % (arc_number, line))
                        cells = line.split(',')
                        new_line = ""
                        if len(cells) == 13:
                            new_lines = cells[:3] + cells[5:8] + cells[9:10]
                            new_line = ",".join(new_lines)
                        if lines.index(line) == 0:
                            if is_first_arc_number:
                                csv_lines.append("ARC#," + new_line)
                                is_first_arc_number = False
                        else:
                            # logger.debug("length: %d, row: %s" % (len(cells), cells))
                            if len(cells) == 0:
                                continue
                            if cells[0] == "V":
                                continue
                            # logger.debug("ADD length: %d, row: %s" % (len(cells), cells))
                            csv_lines.append(arc_number + "," + new_line)

            time.sleep(3)
    except Exception, ex:
        logger.critical(ex)
    finally:
        arc_model.iar_logout(ped, action, arcNumber)
        arc_model.logout()
    # print "over name:%s" % name


def thread_set(is_this_week, csv_lines, is_first_arc_number):
    threads = []
    # thread_lock = threading.Lock()
    for option in conf.options("geoff"):
        thread = MyThread(option, is_this_week, csv_lines=csv_lines, is_first_arc_number=is_first_arc_number)
        thread.start()
        threads.append(thread)

    for t in threads:
        t.join()


conf = ConfigParser.ConfigParser()
conf.read('../iar_update.conf')
arc_model = arc.ArcModel("arc download")
arc_regex = arc.Regex()
logger = arc_model.logger
logger.debug("<<<<<<<<<<<<<<<<<<<<<START>>>>>>>>>>>>>>>>>>>>>")
thread_lock = threading.Lock()
date_time = datetime.datetime.now()
date_week = date_time.weekday()
tick = date_time.strftime('%Y%m%d%H%M%S')
csv_lines = []
is_first_arc_number = True

is_this_week = False
if date_time.weekday() > 1:
    is_this_week = True

thread_set(is_this_week, csv_lines, is_first_arc_number)
# logger.debug("ALL lines: %s" % csv_lines)
# ##---------------------------send email
# print "start send email"
if csv_lines:
    mail_smtp_server = conf.get("email", "smtp_server")
    mail_from_addr = conf.get("email", "from")
    mail_to_addr = conf.get("email", "to_arc_download_assign").split(';')
    mail_subject = 'IAR DOWNLOAD SA CA'
    try:
        body = "\r\n".join(csv_lines)
        floder = "csv"
        file_name = "all_SA_CA.csv"
        arc_model.csv_write(floder, file_name, body)
        mail = arc.Email(smtp_server=mail_smtp_server)
        mail.send(mail_from_addr, mail_to_addr, mail_subject, "Please check the attachment.", [floder + "/" + file_name])
        logger.debug("sent email")
    except Exception as e:
        logger.critical(e)

logger.debug("<<<<<<<<<<<<<<<<<<<<<END>>>>>>>>>>>>>>>>>>>>>")
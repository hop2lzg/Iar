import time
import datetime
import ConfigParser
import arc
import threading
import hashlib
import socket
import base64
import os


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
            # logger.info("execute, is this week %s." % self.is_this_week)
            execute(self.name, self.is_this_week, self.csv_lines, self.is_first_arc_number)
        except Exception as ex:
            logger.fatal(ex)
        finally:
            thread_lock.release()


class WebSocketThread(threading.Thread):
    def __init__(self, connection):
        super(WebSocketThread, self).__init__()
        self.connection = connection

    def __send(self, data):
        self.connection.send('%c%c%s' % (0x81, len(data), data))

    def run(self):
        print 'new websocket client joined!'
        data = self.connection.recv(1024)
        # logger.info("recive data from client: %s" % data)
        re = parse_data(data)
        logger.info("recive data (parse) from client: %s." % re)
        # print re
        self.__send("LOADING")
        global is_this_week
        if re and re == "this week":
            is_this_week = True

        thread_set(is_this_week, csv_lines, is_first_arc_number)
        self.__send("OVER")

        if csv_lines:
            self.__send("CREATE FILE")
            body = "\r\n".join(csv_lines)
            file_path = download_file_path
            file_name = "all_SA_CA.csv"
            arc_model.csv_write(file_path, file_name, body)
            self.__send("download file:" + file_name)
        else:
            self.__send("SORRY FILED!")


class LoopThread(threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)

    def run(self):
        max_count = 60
        for i in range(max_count):
            # print i, is_socket_accepted
            if (i >= max_count-1) and not is_socket_accepted:
                logger.warn("Socket run long time,exit.")
                os._exit(0)
            elif is_socket_accepted:
                break
            else:
                time.sleep(1)


def parse_data(msg):
    v = ord(msg[1]) & 0x7f
    if v == 0x7e:
        p = 4
    elif v == 0x7f:
        p = 10
    else:
        p = 2
    mask = msg[p:p + 4]
    data = msg[p + 4:]

    return ''.join([chr(ord(v) ^ ord(mask[k % 4])) for k, v in enumerate(data)])


def parse_headers(msg):
    headers = {}
    header, data = msg.split('\r\n\r\n', 1)
    for line in header.split('\r\n')[1:]:
        key, value = line.split(': ', 1)
        headers[key] = value
    headers['data'] = data
    return headers


def generate_token(msg):
    key = msg + '258EAFA5-E914-47DA-95CA-C5AB0DC85B11'
    ser_key = hashlib.sha1(key).digest()
    return base64.b64encode(ser_key)


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
            # create_list_html = arc_model.create_list(token, ped, action, arc_number, selectedStatusId="",
            #                                          selectedTransactionType="SA", selectedFormOfPayment="CA",
            #                                          dateTypeRadioButtons="ped", viewFromDate=from_date, viewToDate=to_date,
            #                                          selectedNumberOfResults="20")

            create_list_html = arc_model.create_list(token=token, ped=ped, action=action, arcNumber=arc_number,
                                                     viewFromDate=from_date, viewToDate=to_date, documentNumber="",
                                                     selectedStatusId="", selectedDocumentType="",
                                                     selectedTransactionType="SA", selectedFormOfPayment="CA",
                                                     dateTypeRadioButtons="ped", selectedNumberOfResults="20")

            token = arc_regex.create_list(create_list_html)

            if not token:
                continue
            csv_text = arc_model.get_csv(name, is_this_week, ped, action, arc_number, token, from_date, to_date,
                                         selectedTransactionType="SA", selectedFormOfPayment="CA", tick=tick)

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

HOST = conf.get("socket", "host")
PORT = int(conf.get("socket", "port"))

logger.debug("HOST: %s, PORT: %d" % (HOST, PORT))

download_file_path = conf.get("download", "filePath")

is_this_week = False
# if date_time.weekday() > 1:
#     is_this_week = True

# thread_set(is_this_week, csv_lines, is_first_arc_number)

sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
sock.bind((HOST, PORT))
sock.listen(5)
is_socket_accepted = False
# print "loop thread"
loop_thread = LoopThread()
loop_thread.start()

logger.debug("accept start")
connection, address = sock.accept()
is_socket_accepted = True
logger.debug("accept end")
try:
    data = connection.recv(1024)
    headers = parse_headers(data)
    token = generate_token(headers['Sec-WebSocket-Key'])
    connection.send('\
HTTP/1.1 101 WebSocket Protocol Hybi-10\r\n\
Upgrade: WebSocket\r\n\
Connection: Upgrade\r\n\
Sec-WebSocket-Accept: %s\r\n\r\n' % token)
    thread = WebSocketThread(connection)
    thread.start()
except socket.timeout:
    logger.warn("WebSocket connection timeout")
    # print 'WebSocket connection timeout'



# logger.debug("ALL lines: %s" % csv_lines)
# ##---------------------------send email
# print "start send email"

# if csv_lines:
    # mail_smtp_server = conf.get("email", "smtp_server")
    # mail_from_addr = conf.get("email", "from")
    # mail_to_addr = conf.get("email", "to_arc_download_assign").split(';')
    # mail_subject = 'IAR DOWNLOAD SA CA'
    # try:
    #     body = "\r\n".join(csv_lines)
    #     floder = "csv"
    #     file_name = "all_SA_CA.csv"
    #     arc_model.csv_write(floder, file_name, body)
    #     mail = arc.Email(smtp_server=mail_smtp_server)
    #     mail.send(mail_from_addr, mail_to_addr, mail_subject, "Please check the attachment.", [floder + "/" + file_name])
    #     logger.debug("sent email")
    # except Exception as e:
    #     logger.critical(e)

logger.debug("<<<<<<<<<<<<<<<<<<<<<END>>>>>>>>>>>>>>>>>>>>>")
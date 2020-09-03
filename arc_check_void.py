import time
import datetime
import ConfigParser
import arc
import threading
import hashlib
import socket
import base64
import os
import json


class MyThread(threading.Thread):
    def __init__(self, section, name, tickets, is_this_week):
        threading.Thread.__init__(self)
        self.section = section
        self.name = name
        self.tickets = tickets
        self.is_this_week = is_this_week

    def run(self):
        thread_lock.acquire()
        try:
            execute(self.section, self.name, self.tickets, self.is_this_week)
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
        # print 'new websocket client joined!'
        data = self.connection.recv(1024*1024)
        # logger.info("recive data from client: %s" % data)
        re = parse_data(data)
        logger.info("recive data (parse) from client: %s." % re)
        # print re
        self.__send("LOADING")
        jsonData = json.loads(re)
        is_this_week = jsonData["IsThisWeek"]
        is_mza = jsonData["IsMZA"]
        tickets = jsonData["Iars"]
        logger.info("tickets: %s" % tickets)
        if tickets and len(tickets) > 0:
            thread_set(tickets, is_this_week, is_mza)
        self.__send("OVER")
        logger.debug("items")
        logger.debug(items)
        if items:
            self.__send("CREATE FILE")
            head_name = "ARC,AIR,TKT.#,STATUS,FOP,ERRORS,REMARK"
            lines = []
            lines.append(head_name)
            logger.debug(lines)
            for item in items:
                line = "'%s,'%s,'%s,%s,%s,%s,%s" % (item["arcNumber"], item["airlineNumber"], item["tktNumber"], ("" if item["status"] is None else item["status"]), ("" if item["FOP"] is None else item["FOP"]), ("" if item["certificates"] is None else item["certificates"]), "")
                lines.append(line)

            logger.info(lines)
            body = "\r\n".join(lines)
            file_path = download_file_path
            file_name = "iar_check_void.csv"
            logger.debug(body)
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


def execute(section, name, tickets, is_this_week):
    logger.debug("login: %s" % name)
    # print "start name:%s" % name
    password = conf.get(section, name)
    if not arc_model.execute_login(name, password):
        return

    iar_html = arc_model.iar()
    if not iar_html:
        logger.error('open iar error: %s' % name)
        return

    ped, action, arcNumber = arc_regex.iar(iar_html, is_this_week=is_this_week)
    if not action:
        logger.error('regex iar error: %s' % name)
        arc_model.logout()
        return

    list_transactions_html = arc_model.listTransactions(ped, action, arcNumber)
    token, from_date, to_date = arc_regex.listTransactions(list_transactions_html)
    if not token:
        return

    try:
        for ticket in tickets:
            item = {"arcNumber": ticket["ArcNumber"], "airlineNumber": ticket["AirlineNumber"],
                    "tktNumber": ticket["TktNumber"], "status": "NOT CHECK", "FOP": "",
                    "certificates":""}
            arc_number = ticket["ArcNumber"]
            tkt_number = ticket["TktNumber"]
            # ticketing_date = datetime.datetime.strptime(ticket["TicketingDate"], '%Y-%m-%dT%H:%M:%S')
            # entry_date = ticketing_date + datetime.timedelta(days=1)
            # from_date = (entry_date + datetime.timedelta(days=(-entry_date.weekday()))).strftime('%d%b%y').upper()
            # to_date = (entry_date + datetime.timedelta(days=(6 - entry_date.weekday()))).strftime('%d%b%y').upper()
            ped = to_date

            create_list_html = arc_model.create_list(token=token, ped=ped, action=action, arcNumber=arc_number,
                                                     viewFromDate=from_date, viewToDate=to_date, documentNumber=tkt_number,
                                                     selectedStatusId="", selectedDocumentType="",
                                                     selectedTransactionType="", selectedFormOfPayment="",
                                                     dateTypeRadioButtons="", selectedNumberOfResults="")

            if not create_list_html:
                continue

            create_list_token = arc_regex.create_list(create_list_html)
            if not create_list_token:
                continue

            if create_list_html.find("No Transactions Found") >= 0:
                item["status"] = "No Transactions Found"
            else:
                status = arc_regex.get_status(create_list_html)
                item["status"] = status
                if status != "V":
                    seqNum, documentNumber = arc_regex.search(create_list_html)
                    if not seqNum:
                        continue

                    modify_tran_html = arc_model.modifyTran(seqNum, documentNumber)
                    if not modify_tran_html:
                        continue

                    item["FOP"] = arc_regex.get_form_of_payment(modify_tran_html)
                    certificate_items = arc_regex.get_certificate_items(modify_tran_html)
                    item["certificates"] = ";".join(certificate_items)
            items.append(item)
    except Exception, ex:
        logger.critical(ex)
    finally:
        arc_model.iar_logout(ped, action, arcNumber)
        arc_model.logout()


def thread_set(tickets, is_this_week, is_mza):
    threads = []
    # thread_lock = threading.Lock()
    section = "extra"
    if is_mza:
        section = "extra-mza"

    for option in conf.options(section):
        if is_mza and option != "gttmza-bw":
            continue
        thread = MyThread(section, option, tickets, is_this_week)
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
items = []

HOST = conf.get("socket", "host")
PORT = int(conf.get("socket", "port"))

logger.debug("HOST: %s, PORT: %d" % (HOST, PORT))

download_file_path = conf.get("download", "statementPath")

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
    data = connection.recv(1024*1024)
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


logger.debug("<<<<<<<<<<<<<<<<<<<<<END>>>>>>>>>>>>>>>>>>>>>")
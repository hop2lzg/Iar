import time
import arc
import ConfigParser
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
            run(self.section, self.name, self.tickets, self.is_this_week)
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
        try:
            jsonData = json.loads(re)
            is_this_week = jsonData["IsThisWeek"]
            is_mza = jsonData["IsMZA"]
            tickets = jsonData["Iars"]
            logger.info("tickets: %s" % tickets)
            if tickets and len(tickets) > 0:
                thread_set(tickets, is_this_week, is_mza)
                logger.debug("iar execute over")
                fails = []
                for t in tickets:
                    if t["status"] is None or t["status"] == 0 or t["status"] == 4:
                        fails.append(t["TktNumber"])

                logger.warn("Fails: %s" % fails)
                logger.info("Fails Length: %d" % len(fails))
                if len(fails) == 0:
                    self.__send("success")
                else:
                    self.__send("FAILED: %s" % fails)
            else:
                self.__send("Ticket NOT FOUND")
        except Exception as ex:
            logger.fatal(ex)
            self.__send(ex.message)


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


conf = ConfigParser.ConfigParser()
conf.read('../iar_update.conf')
arc_model = arc.ArcModel("ARC UPDATE TICKET")
arc_regex = arc.Regex()
logger = arc_model.logger

logger.debug('--------------<<<START>>>--------------')
thread_lock = threading.Lock()
HOST = conf.get("socket", "host")
PORT = int(conf.get("socket", "port"))
logger.debug("HOST: %s, PORT: %d" % (HOST, PORT))
is_this_week = False

agent_codes = conf.get("certificate", "agentCodes").split(',')


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


def execute(ped, action, ticket, token, from_date, to_date):
    ticket["status"] = 0
    arc_number = ticket["ArcNumber"]
    airline_number = ticket["AirlineNumber"]
    document_number = ticket["TktNumber"]
    # result = {'void': 0, 'update': 0}
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
        ticket["status"] = 2
        return
    elif is_void_pass == 1:
        ticket["status"] = 1
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
            ticket["status"] = 3
            logger.info("UPDATE SUCCESS.")
        else:
            ticket["status"] = 4
            logger.warning('UPDATE MAY BE ERROR.')

    # return result


def run(section, user_name, tickets, is_this_week=True):
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

    for ticket in tickets:
        execute(ped, action, ticket, token,
                         from_date, to_date)
        # if result:
        #     ticket["void"] = result["void"]
        #     ticket["update"] = result["update"]

        logger.info(ticket)
    arc_model.iar_logout(ped, action, arc_number)
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


logger.debug('--------------<<<END>>>--------------')
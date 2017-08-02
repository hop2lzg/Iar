import time, datetime
import ConfigParser
import arc
# import thread
import threading


class MyThread(threading.Thread):
    def __init__(self, name, is_this_week):
        threading.Thread.__init__(self)
        self.name = name
        self.is_this_week = is_this_week

    def run(self):
        print "runnuing: " + self.name + '\n'
        thread_lock.acquire()
        try:
            execute(self.name, self.is_this_week)
        except Exception as ex:
            logger.fatal(ex)
        finally:
            thread_lock.release()


def execute(name, is_this_week):
    print "start name:%s" % name
    password = conf.get("login", name)
    login_html = arc_model.login(name, password)
    if login_html.find('You are already logged into My ARC') < 0 and login_html.find('Account Settings :') < 0:
        logger.error('login error: '+name)
        return

    iar_html = arc_model.iar()
    if not iar_html:
        logger.error('iar error: '+name)
        return

    ped, action, arcNumber = arc_regex.iar(iar_html, is_this_week)
    if not action:
        logger.error('regex iar error: '+name)
        arc_model.logout()
        return

    try:
        if name == 'mulingpeng':
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
            if not token:
                continue
            arc_model.get_csv(name, is_this_week, ped, action, arc_number, token, from_date, to_date)
            time.sleep(3)
    except Exception, ex:
        # print ex
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
if date_week != 0:
    print 'this week'
    thread_set(True)

if date_week == 0 or date_week == 1 or date_week == 2:
    print 'last week'
    timer_sleep = 0
    if date_week == 1 or date_week == 2:
        timer_sleep = 1*60*20
    timer = threading.Timer(timer_sleep, thread_set, [False])
    timer.start()

logger.debug("<<<<<<<<<<<<<<<<<<<<<END>>>>>>>>>>>>>>>>>>>>>")
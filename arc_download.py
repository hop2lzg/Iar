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
        print "Starting " + self.name + '\n'

        execute(self.name, self.is_this_week)


conf = ConfigParser.ConfigParser()
conf.read('../arc_update.conf')
# print conf.sections()
# print conf.options("login") 
# pwd = conf.get("login","muling-yww")

arc_model = arc.ArcModel("arc download")
arc_regex = arc.Regex()
logger = arc_model.logger


def execute(name, is_this_week):
    password = conf.get("login", name)
    arc_model.login(name, password)
    iar_html = arc_model.iar()
    ped, action, arcNumber = arc_regex.iar(iar_html, is_this_week)
    # print ped, action, arcNumber
    if not action:
        arc_model.logout()
        return
    try:
        if name == 'mulingpeng':
            conf_name = 'all'
        else:
            conf_name = name[-3:]
        arc_numbers = conf.get("arc", conf_name).split(',')
        if arcNumber in arc_numbers:
            arc_numbers.remove(arcNumber)
        for arc_number in arc_numbers:
            print 'Downloading ' + arc_number
            listTransactions_html = arc_model.listTransactions(ped, action, arc_number)
            token, from_date, to_date = arc_regex.listTransactions(listTransactions_html)
            if not token:
                continue
            arc_model.get_csv(name, is_this_week, ped, action, arc_number, token, from_date, to_date)
            time.sleep(3)
    except Exception, ex:
        print ex
        logger.critical(ex)
    finally:
        arc_model.iar_logout(ped, action, arcNumber)
        arc_model.logout()


def thread_set(is_this_week):
    threads = []
    for option in conf.options("login"):
        thread = MyThread(option, is_this_week)
        thread.start()
        threads.append(thread)
    # thread.start_new_thread(execute,(is_this_week,option))
    for t in threads:
        t.join()


date_time = datetime.datetime.now()
date_week = date_time.weekday()

# if date_week!=0:
# 	print 'this week'
# 	thread_set(True)
# if date_week==0 or date_week==1 or date_week==2:
# 	print 'last week'
# 	time.sleep(1*60*30)
# 	thread_set(False)

# def execute(is_this_week):
# 	for option in conf.options("login"):
# 		name=option
# 		password=conf.get("login",option)

# 		arc_model.login(name,password)
# 		iar_html=arc_model.iar()

# 		ped,action,arcNumber=arc_regex.iar(iar_html)
# 		if not action:
# 			arc_model.logout()
# 			continue

# 		try:
# 			arc_numbers=['05500445','05507073','05513826','05520502','05545783','05563983','05613495','05635814','05639255','05649125','05765476','06542082','09502964','10522374','11521436','14537891','14646015','17581351','18503306','21524952','22505851','23534803','24514571','26503945','31533003','33508333','33519032','33547544','33583454','33589544','34517840','36537502','37531152','39654591','45532885','45666574','45668574','46543291','49587775','50622154']
# 			if name=='muling-yww' or name=='muling-tvo':
# 				arc_numbers=[]
# 			elif name=='muling-aca':
# 				arc_numbers=['05617986']

# 			arc_numbers[0:0]=arcNumber

# 			for arc_number in arc_numbers:
# 				print 'Downloading '+arc_number
# 				listTransactions_html=arc_model.listTransactions(ped,action,arc_number)
# 				token,from_date,to_date=arc_regex.listTransactions(listTransactions_html)
# 				if not token:
# 					continue
# 				arc_model.get_csv(name,is_this_week,ped,action,arc_number,token,from_date,to_date)
# 				time.sleep(3)

# 		except Exception,ex:
# 			logger.critical(ex)
# 			print ex
# 		finally:
# 			arc_model.iar_logout(ped,action,arcNumber)
# 			arc_model.logout()

import arc
import ConfigParser
import sys
import datetime

arc_model=arc.ArcModel("arc remove error")
arc_regex=arc.Regex()
logger=arc_model.logger

logger.debug('--------------<<<START>>>--------------')
conf = ConfigParser.ConfigParser()
conf.read('../iar_update.conf')

#----------------------login
name="mulingpeng"
password=conf.get("login",name)
login_html=arc_model.login(name,password)
if login_html.find('You are already logged into My ARC')<0 and login_html.find('Account Settings :')<0:
	logger.error('login error')
	sys.exit(0)

# -------------------go to IAR
iar_html=arc_model.iar()
if not iar_html:
	logger.error('iar error')
	sys.exit(0)
ped,action,arcNumber=arc_regex.iar(iar_html,False)
if not action:
	logger.error('regex iar error')
	arc_model.logout()
	sys.exit(0)

listTransactions_html=arc_model.listTransactions(ped,action,arcNumber)
if not listTransactions_html:
	logger.error('listTransactions error')
	arc_model.iar_logout(ped,action,arcNumber)
	arc_model.logout()
	sys.exit(0)

token,from_date,to_date=arc_regex.listTransactions(listTransactions_html)
if not token:
	logger.error('regex listTransactions error')
	arc_model.iar_logout(ped,action,arcNumber)
	arc_model.logout()
	sys.exit(0)




# def RegexTransactionConfirmation(html):
# 	documentNumber=commission=None
# 	pattern_documentNumber = re.compile(r'<a href="/IAR/modifyTran\.do\?seqNum=\d{10}&amp;documentNumber=(\d{10})">')
# 	m_documentNumber=pattern_documentNumber.search(html)
# 	if(m_documentNumber!=None):
# 		documentNumber=m_documentNumber.group(1)
# 	pattern_commission=re.compile(r'<td width="8%" align="right" nowrap>(\d+\.\d{2}) &nbsp;</td>')
# 	m_commission=pattern_commission.search(html)
# 	if(m_commission!=None):
# 		commission=m_commission.group(1)
# 	return documentNumber,commission

def Execute(data):
	seqNum=data['seqNum']
	documentNumber=data['documentNumber']
	logger.debug('ducumentNumber: %s' % documentNumber)
	modify_html=arc_model.modifyTran(seqNum,documentNumber)
	if(modify_html==None):
		return

	voided_index=modify_html.find('Document is being displayed as view only')
	if voided_index>=0:
		data['status']=2
		return

	token,maskedFC,regex_commission,waiverCode,list_certificates=arc_regex.modifyTran(modify_html)
	logger.debug("regex commission:"+regex_commission)
	if(token==None or maskedFC==None):
		return

	financialDetails_html=arc_model.financialDetailsRemoveError(token,regex_commission,waiverCode,maskedFC,seqNum,documentNumber,list_certificates)
	if not financialDetails_html:
		return

	token=arc_regex.itineraryEndorsements(financialDetails_html)


	if token:
		transactionConfirmation_html=arc_model.transactionConfirmation(token)
		if transactionConfirmation_html:
			if transactionConfirmation_html.find('Document has been modified')>=0:
				data['status']=1
			else:
				logger.warning('Document can not modify')

def Remove(today,weekday,ped,action,arcNumber):
	listTransactions_html=arc_model.listTransactions(ped,action,arcNumber)
	if not listTransactions_html:
		logger.error('go to listTransactions_html error')
		return
	token,from_date,to_date=arc_regex.listTransactions(listTransactions_html)
	if not token:
		logger.error('regex listTransactions token error')
		return
	search_html=arc_model.searchError(ped,action,arcNumber,token,from_date,to_date)
	if not search_html:
		logger.error('go to seach error')
		return

	list_entry_date=[]
	if weekday>=2:
		list_entry_date.append((today+datetime.timedelta(days = -2)).strftime('%d%b%y').upper())
		# list_entry_date.append((today+datetime.timedelta(days = -1)).strftime('%d%b%y').upper())
	# if weekday==3 or weekday==4:
	# 	list_entry_date.append((today+datetime.timedelta(days = -3)).strftime('%d%b%y').upper())
	# elif weekday==5:
	# 	list_entry_date.append((today+datetime.timedelta(days = -3)).strftime('%d%b%y').upper())
	# 	list_entry_date.append((today+datetime.timedelta(days = -2)).strftime('%d%b%y').upper())

	entry_date='\d{2}[A-Z]{3}\d{2}'
	if list_entry_date:
		entry_date='|'.join(list_entry_date)

	list_regex_search=arc_regex.searchError(search_html,entry_date)

	if not list_regex_search:
		logger.warning('regex seach error')
		return
	

	for i in list_regex_search:

			v={}
			v['ticketNumber']=i[2]+i[0]
			v['seqNum']=i[1]
			v['documentNumber']=i[0]
			v['date']=i[3]
			v['arcNumber']=arcNumber
			v['status']=0
			
			Execute(v)

			list_data.append(v)

list_data=[]
try:
	date_time=datetime.datetime.now()
	date_week=date_time.weekday()
	date_ped=date_time+datetime.timedelta(days = (6-date_time.weekday()))

	if date_week<2:
	    date_ped=date_ped+datetime.timedelta(days = -7)
	    
	# from_date=(date_ped+datetime.timedelta(days = -6)).strftime('%d%b%y').upper()
	ped=date_ped.strftime('%d%b%y').upper()
	action="7"

	arcNumbers=conf.get("arc","all").split(',')
	arcNumbers[0:0]=["45668571"]
	# list_arc=['45668571','05500445','05507073','05513826','05520502','05545783','05563983','05613495','05635814','05639255','05649125','05765476','06542082','09502964','10522374','11521436','14537891','14646015','17581351','18503306','21524952','22505851','23534803','24514571','26503945','31533003','33508333','33519032','33547544','33583454','33589544','34517840','36537502','37531152','39654591','45532885','45666574','45668574','46543291','49587775','50622154']
	for arcNumber in arcNumbers:
		logger.debug(arcNumber)
		Remove(date_time,date_week,ped,action,arcNumber)

except Exception as e:
	# print e
	logger.critical(e)


mail_smtp_server=conf.get("email","smtp_server")
mail_from_addr=conf.get("email","from")
mail_to_addr=conf.get("email","to_remove_error").split(';')
mail_sub=conf.get("email","subject")+" remove error"

try:
	body=''

	for i in list_data:
		status='No'
		if i['status']==1:
			status='Yes'
		elif i['status']==2:
			status='Void'
		body=body+'ARC:%s  TicketNumber:%s  Date:%s Updated:%s\n' %(i['arcNumber'],i['ticketNumber'],
			i['date'],status)
	# print body
	# print send_mail(mailto_list,"iar commission updated",body)
	mail=arc.SendEmail(from_addr=mail_from_addr,to_addr=mail_to_addr
		,smtp_server=mail_smtp_server,subject=mail_sub)
	mail.send(body)
except Exception as e:
	logger.critical(e)


# print "OVER"
logger.debug('--------------<<<END>>>--------------')
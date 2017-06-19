import arc
import ConfigParser
import sys
import datetime

def Execute(post,action,token,from_date,to_date):
	arcNumber=post['ArcNumber']
	documentNumber=post['Ticket']
	date=post['IssueDate']
	is_check=False
	if post['Status']==3:
		is_check=True

	date_time=datetime.datetime.strptime(date,'%Y-%m-%d')
	ped=(date_time+datetime.timedelta(days = (6-date_time.weekday()))).strftime('%d%b%y').upper()
	logger.info("PED: "+ped+" arc: "+arcNumber+" tkt: "+documentNumber)
	search_html=arc_model.search(ped,action,arcNumber,token,from_date,to_date,documentNumber)
	if not search_html:
		return

	seqNum,documentNumber=arc_regex.search(search_html)
	if not seqNum:
		return

	modify_html=arc_model.modifyTran(seqNum,documentNumber)
	if not modify_html:
		return

	voided_index=modify_html.find('Document is being displayed as view only')
	if voided_index>=0:
		post['Status']=2
		return

	token,maskedFC,arc_commission,waiverCode,certificates=arc_regex.modifyTran(modify_html)
	logger.debug("regex commission:"+arc_commission)
	if not token:
		return

	financialDetails_html=arc_model.financialDetailsPutError(token,is_check,arc_commission,waiverCode,maskedFC,seqNum,documentNumber,certificates)
	if not financialDetails_html:
		return

	token=arc_regex.itineraryEndorsements(financialDetails_html)
	if token:
		transactionConfirmation_html=arc_model.transactionConfirmation(token)
		if transactionConfirmation_html:
			if transactionConfirmation_html.find('Document has been modified')>=0:
				post['Status']=1
			else:
				logger.warning('update may be error')

arc_model=arc.ArcModel("arc put error")
arc_regex=arc.Regex()
logger=arc_model.logger


logger.debug('--------------<<<START>>>--------------')
conf = ConfigParser.ConfigParser()
conf.read('../iar_update.conf')


#----------------------sql
logger.debug('select sql')
sql_server = conf.get("sql","server")
sql_database = conf.get("sql","database")
sql_user = conf.get("sql","user")
sql_pwd = conf.get("sql","pwd")

ms = arc.MSSQL(server=sql_server,db=sql_database,user=sql_user,pwd=sql_pwd)
sql=('''declare @t date
set @t=DATEADD(DAY,-1,GETDATE())
select [SID],TicketNumber,substring(TicketNumber,4,10) Ticket,IssueDate,ArcNumber,PaymentType from Ticket
where Status not like '[NV]%'
and IssueDate=@t
and QCStatus=2
and (TicketNumber not like '0447%' or TicketNumber not like '0457%' or TicketNumber not like '0477%' or TicketNumber not like '1347%' or TicketNumber not like '1397%' or TicketNumber not like '1477%' or TicketNumber not like '2307%' or TicketNumber not like '2697%' or TicketNumber not like '4627%' or TicketNumber not like '4697%' or TicketNumber not like '5447%' or TicketNumber not like '8377%' or TicketNumber not like '9577%')
and Id not in (
select Id from Ticket
where Status not like '[NV]%'
and FareType<>'BULK'
and IssueDate=@t
and Comm>=0
and Comm<QCComm-5
and (ISNULL(TourCode,'')='' or ((TicketNumber like '00[16]7%' or TicketNumber like '0147%' or TicketNumber like '1767%' or TicketNumber like '6077%' or TicketNumber like '0727%' or TicketNumber like '0657%' or TicketNumber like '1577%' or TicketNumber like '2357%' or TicketNumber like '5557%' or TicketNumber like '1607%') and TourCode<>''))
)
order by ArcNumber,Ticket
''')

list_data=arc_model.load()
if not list_data:
	rows = ms.ExecQuery(sql)
	if(len(rows)==0):
		sys.exit(0)
	list_data=[]
	for i in rows:
		v={}
		v['TicketNumber']=i.TicketNumber
		v['Ticket']=i.Ticket
		v['IssueDate']=str(i.IssueDate)
		v['ArcNumber']=i.ArcNumber
		v['Status']=0
		if i.PaymentType != 'C':
			v['Status']=3
		list_data.append(v)

#----------------------login
name="mulingpeng"
password=conf.get("login",name)
login_html=arc_model.login(name,password)
if login_html.find('You are already logged into My ARC')<0 and login_html.find('Account Settings :')<0:
	logger.error('login error')
	sys.exit(0)


#-------------------go to IAR
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

try:
	for data in list_data:
		# print data
		# if data['Status'] != 0 and data['Status'] != 3:
		if data['Status'] != 0:
			continue
		Execute(data,action,token,from_date,to_date)
		# time.sleep(3)
except Exception as e:
	# print e
	logger.critical(e)
finally:
	arc_model.store(list_data)




##---------------------------send email
mail_smtp_server=conf.get("email","smtp_server")
mail_from_addr=conf.get("email","from")
mail_to_addr=conf.get("email","to").split(';')
mail_sub=conf.get("email","subject")+' put error'

try:
	body=''

	for i in list_data:
		status='No'
		if i['Status']==1:
			status='Yes'
		elif i['Status']==2:
			status='Void'
		elif i['Status']==3:
			status='Check'
			
		body=body+'''<tr>
			<td>%s</td>
			<td>%s</td>
			<td>%s</td>
			<td>%s</td>
		</tr>''' % (i['ArcNumber'],i['TicketNumber'],
			i['IssueDate'],status)

		# body=body+'ARC:%s  TicketNumber:%s  Date:%s  COMM:%-6s  TourCode:%-10s  Updated:%s\n' %(i['ArcNumber'],i['TicketNumber'],
		# 	i['IssueDate'],i['QCComm'],i['QCTourCode'],status)
	# print body

	body='''<table border=1>
	<thead>
		<tr>
			<th>ARC</th>
			<th>TicketNumber</th>
			<th>Date</th>
			<th>Updated</th>
		</tr>
	</thead>
	<tbody>%s
	</tbody>
</table>
''' % body
	# print send_mail(mailto_list,"iar commission updated",body)
	mail=arc.SendEmail(from_addr=mail_from_addr,to_addr=mail_to_addr
		,smtp_server=mail_smtp_server,subject=mail_sub)
	mail.send(body)
except Exception as e:
	logger.critical(e)

arc_model.iar_logout(ped,action,arcNumber)
arc_model.logout()

print "OVER"
logger.debug('--------------<<<END>>>--------------')
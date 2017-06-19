import arc
import ConfigParser
import sys
import datetime

def Execute(post,action,token,from_date,to_date):
	arcNumber=post['ArcNumber']
	documentNumber=post['Ticket']
	date=post['IssueDate']
	commission=post['QCComm']
	tour_code=""
	if post['TourCode']:
		tour_code=post['TourCode']
	qc_tour_code=""
	if post['QCTourCode']:
		qc_tour_code=post['QCTourCode']

	is_check=False

	if post['Status']==3:
		is_check=True

	date_time=datetime.datetime.strptime(date,'%Y-%m-%d')
	ped=(date_time+datetime.timedelta(days = (6-date_time.weekday()))).strftime('%d%b%y').upper()
	
	logger.info("UPDATING PED: "+ped+" arc: "+arcNumber+" tkt: "+documentNumber)
	
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
	if not token:
		return
	post['ArcComm']=arc_commission

	financialDetails_html=arc_model.financialDetails(token,is_check,commission,waiverCode,maskedFC,seqNum,documentNumber,tour_code,qc_tour_code,certificates)
	if not financialDetails_html:
		return

	
	token,arc_tour_code,backOfficeRemarks,ticketDesignators=arc_regex.financialDetails(financialDetails_html)
	if not token:
		return
	post['ArcTourCode']=arc_tour_code

	if ticketDesignators:
		list_ticketDesignator=[]
		for ticketDesignator in ticketDesignators:
			list_ticketDesignator.append(ticketDesignator[1])
		post['TicketDesignator']='/'.join(list_ticketDesignator)
	if tour_code!=qc_tour_code:
		itineraryEndorsements_html=arc_model.itineraryEndorsements(token,qc_tour_code,backOfficeRemarks,ticketDesignators)
		if not itineraryEndorsements_html:
			return
		token=arc_regex.itineraryEndorsements(itineraryEndorsements_html)
	if token:
		transactionConfirmation_html=arc_model.transactionConfirmation(token)
		if transactionConfirmation_html:
			if transactionConfirmation_html.find('Document has been modified')>=0:
				post['Status']=1
			else:
				logger.warning('update may be error')

			# updated_ticket_number,updated_commission=RegexTransactionConfirmation(transactionConfirmation_html)
			
			# # logger.info("documentNumber:"+updated_ticket_number+" commission:"+updated_commission)
			# # print documentNumber,updated_ticket_number,type(documentNumber),type(updated_ticket_number)
			# # print commission,updated_commission,type(commission),type(updated_commission)
			# if(documentNumber==updated_ticket_number and str(commission)==updated_commission):
			# 	# Write(documentNumber)
			# 	post['Status']=1
			# else:
			# 	print 'no write'

def Check(post,action,token,from_date,to_date):
	arcNumber=post['ArcNumber']
	documentNumber=post['Ticket']
	date=post['IssueDate']
	commission=post['QCComm']
	tour_code=""
	if post['TourCode']:
		tour_code=post['TourCode']
	qc_tour_code=""
	if post['QCTourCode']:
		qc_tour_code=post['QCTourCode']

	is_check=False

	if post['Status']==3:
		is_check=True

	date_time=datetime.datetime.strptime(date,'%Y-%m-%d')
	ped=(date_time+datetime.timedelta(days = (6-date_time.weekday()))).strftime('%d%b%y').upper()
	
	logger.info("CHECK PED: "+ped+" arc: "+arcNumber+" tkt: "+documentNumber)
	
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
	if not token:
		return
	post['ArcCommUpdated']=arc_commission

	financialDetails_html=arc_model.financialDetails(token,is_check,commission,waiverCode,maskedFC,seqNum,documentNumber,tour_code,qc_tour_code,certificates,True)
	if not financialDetails_html:
		return

	
	token,arc_tour_code,backOfficeRemarks,ticketDesignators=arc_regex.financialDetails(financialDetails_html)
	if not token:
		return
	post['ArcTourCodeUpdated']=arc_tour_code

def update(datas):
	list_id=[]
	for i in list_data_sql:
		if i['Status']!=0:
			if i['QcId'] not in list_id:
				list_id.append("'"+i['QcId']+"'")
	if list_id:
		sql="update TicketQC set ARCupdated=1 where Id in (%s)" % ','.join(list_id)
		try:
			if(ms.ExecNonQuery(sql)>0):
				logger.info('update sql success')
		except Exception as e:
			logger.critical(e) 
def insert(datas):
	if not datas:
		return
	insert_sql=''
	is_updated=0
	list_id=[]
	for data in datas:
		if data['Id'] in list_id:
			continue;
		list_id.append(data['Id'])
		if data['QCComm']==data['ArcCommUpdated'] and data['QCTourCode']==data['ArcTourCodeUpdated']:
			is_updated=1
		comm=None
		try:
			comm=float(data['ArcCommUpdated'])
		except ValueError:
			comm="null"
		if not data['ArcId']:
			insert_sql=insert_sql+'''insert into IarUpdate(Id,Commission,TourCode,TicketDesignator,IsUpdated,TicketId) values (newid(),%s,'%s','%s',%d,'%s');''' %(comm,data['ArcTourCodeUpdated'],data['TicketDesignator'],is_updated,data['Id'])
		else:
			insert_sql=insert_sql+'''update IarUpdate set Commission=%s,TourCode='%s',TicketDesignator='%s',IsUpdated=%d where Id='%s';''' %(comm,data['ArcTourCodeUpdated'],data['TicketDesignator'],is_updated,data['ArcId'])

	logger.info(insert_sql)
	rowcount = ms.ExecNonQuery(insert_sql)
	if rowcount>0:
		logger.info('insert success')
	else:
		logger.error('insert error')

arc_model=arc.ArcModel("arc update commission by user")
arc_regex=arc.Regex()
logger=arc_model.logger

logger.debug('--------------<<<START>>>--------------')
logger.debug('select sql')

conf = ConfigParser.ConfigParser()
conf.read('../iar_update.conf')
sql_server = conf.get("sql","server")
sql_database = conf.get("sql","database")
sql_user = conf.get("sql","user")
sql_pwd = conf.get("sql","pwd")


ms = arc.MSSQL(server=sql_server,db=sql_database,user=sql_user,pwd=sql_pwd)


sql=('''
declare @t date
set @t=dateadd(day,-2,getdate())
select aa.*,iu.Id IarId from (
select t.Id,qc.Id qcId,t.TicketNumber,substring(t.TicketNumber,4,10) Ticket,t.IssueDate,t.ArcNumber,t.PaymentType,
t.Comm,t.TourCode,t.QCComm UpdatedComm,t.QCTourCode UpdatedTourCode from Ticket t
left join TicketQC qc
on t.Id=qc.TicketId
where qc.ARCupdated=0
and (qc.OPUser='GW' or OPLastUser='GW')
and t.QCStatus=2 and qc.AGStatus=1 and qc.OPStatus<>2
and t.CreateDate>=@t
union
select t.Id,qc.Id qcId,t.TicketNumber,substring(t.TicketNumber,4,10) Ticket,t.IssueDate,t.ArcNumber,t.PaymentType,
t.Comm,t.TourCode,qc.AGComm UpdatedComm,qc.AGTourCode UpdatedTourCode from Ticket t
left join TicketQC qc
on t.Id=qc.TicketId
where qc.ARCupdated=0
and t.QCStatus=2 and qc.AGStatus=3
and ((t.Comm<>qc.AGComm) or (t.TourCode<>qc.AGTourCode))
and qc.OPStatus<>2
and (qc.OPUser='GW' or OPLastUser='GW')
and t.CreateDate>=@t
union
select t.Id,qc.Id qcId,t.TicketNumber,substring(t.TicketNumber,4,10) Ticket,t.IssueDate,t.ArcNumber,t.PaymentType,
t.Comm,t.TourCode,qc.OPComm UpdatedComm,qc.OPTourCode UpdatedTourCode from Ticket t
left join TicketQC qc
on t.Id=qc.TicketId
where qc.ARCupdated=0
and qc.OPStatus=2
and ((t.Comm<>qc.OPComm) or (t.TourCode<>qc.OPTourCode))
and (qc.OPUser='GW' or OPLastUser='GW')
and t.CreateDate>=@t
) aa
left join IarUpdate iu
on aa.Id=iu.TicketId
order by IssueDate
''')

# rows = ms.ExecQuery(sql)
#
# logger.debug(len(rows))
# if(len(rows)==0):
# 	sys.exit(0)

# list_data_sql=[]

# for i in rows:
# 	v={}
# 	v['Id']=i.Id
# 	v['QcId']=i.qcId
# 	v['TicketNumber']=i.TicketNumber
# 	v['Ticket']=i.Ticket
# 	v['IssueDate']=str(i.IssueDate)
# 	v['ArcNumber']=i.ArcNumber
# 	v['Comm']=str(i.Comm)
# 	v['QCComm']=str(i.UpdatedComm)
# 	v['TourCode']=i.TourCode
# 	v['QCTourCode']=i.UpdatedTourCode
# 	v['ArcComm']=''
# 	v['ArcTourCode']=''
# 	v['TicketDesignator']=''
# 	v['ArcCommUpdated']=''
# 	v['ArcTourCodeUpdated']=''
# 	v['ArcId']=i.IarId
# 	v['Status']=0
# 	if i.PaymentType!='C':
# 		v['Status']=3
# 	list_data_sql.append(v)

# logger.debug(list_data_sql)
#----------------------login

name="muling-aca"
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

#
# try:
# 	for data in list_data_sql:
# 		# print data
# 		if data['Status'] != 0 and data['Status'] != 3:
# 			continue
# 		Execute(data,action,token,from_date,to_date)
# 		if data['Status'] == 1:
# 			Check(data,action,token,from_date,to_date)
# except Exception as e:
# 	print e
# 	logger.critical(e)
# finally:
# 	arc_model.store(list_data_sql)
#
# try:
# 	update(list_data_sql)
# except Exception as e:
# 	logger.critical(e)
#
# try:
# 	insert(list_data_sql)
# except Exception as e:
# 	logger.critical(e)




# mail_smtp_server=conf.get("email","smtp_server")
# mail_from_addr=conf.get("email","from")
# mail_to_addr=conf.get("email","to_update_user").split(';')
# mail_subject=conf.get("email","subject")+" by user"
# try:
# 	body=''
#
# 	for i in list_data_sql:
# 		status,is_updated=arc_model.convertStatus(i)
# 		body=body+'''<tr>
# 			<td>%s</td>
# 			<td>%s</td>
# 			<td>%s</td>
# 			<td>%s</td>
# 			<td>%s</td>
# 			<td>%s</td>
# 			<td>%s</td>
# 			<td>%s</td>
# 			<td>%s</td>
# 			<td>%s</td>
# 		</tr>''' % (i['ArcNumber'],i['TicketNumber'],
# 			i['IssueDate'],i['Comm'],i['QCComm'],i['ArcCommUpdated'],i['TourCode'],i['QCTourCode'],
# 			i['ArcTourCodeUpdated'],is_updated)
#
# 	# print body
# 	body='''<table border=1>
# 	<thead>
# 		<tr>
# 			<th>ARC</th>
# 			<th>TicketNumber</th>
# 			<th>Date</th>
# 			<th>TKCM</th>
# 			<th>QCCM</th>
# 			<th>ARCCM</th>
# 			<th>TKTC</th>
# 			<th>QCTC</th>
# 			<th>ARCTC</th>
# 			<th>Status</th>
# 		</tr>
# 	</thead>
# 	<tbody>%s
# 	</tbody></table>''' % body
# 	mail=arc.Email(smtp_server=mail_smtp_server)
# 	mail.send(mail_from_addr,mail_to_addr,mail_subject,body)
#
# 	# mail=arc.SendEmail(from_addr=mail_from_addr,to_addr=mail_to_addr
# 	# 	,smtp_server=mail_smtp_server,subject=mail_sub)
# 	# mail.send(body)
#
# 	logger.info('email sent')
# except Exception as e:
# 	logger.critical(e)

arc_model.iar_logout(ped,action,arcNumber)
arc_model.logout()

logger.debug('--------------<<<END>>>--------------')
import time
# agent_codes = []
#
# str = "GW,GW,"
#
# gw_level_agent_codes =filter(None, str.split(','))
# print gw_level_agent_codes
# op_users_sql = ""
# print [("'" + x + "'") for x in gw_level_agent_codes]
# op_users_where = ",".join([("'" + x + "'") for x in gw_level_agent_codes])
# print op_users_where
# print agent_codes
# for agent_code in agent_codes:
#     print agent_code
#     agent_code = "'" + agent_code + "'"

# print agent_codes
# print ",".join(agent_codes)
# print ",".join([("'"+x+"'") for x in agent_codes])

# value = {'name': 'zg', 'age': 10 }
# head = 1
# print value
# print "value: %s, header: %s" % (value, head)
# for tries in range(2):
#     print "at %d times" % tries
#     time.sleep(2)
#     print tries

#
# s = 0
# if not s:
#     print "not"
# else:
#     print "yes"


# filter(lambda x: x['ArcNumber'] in arc_numbers, list_data)
# agent_codes = "MJ,M1,M2,DUP,AUTOQC,AG-AGREE,AG-PENDING,OP,QC-PROFIT,AUTOERROR,QC-ERROR".split(',')
# error_codes = conf.get("error", "errorCodes")
# sub_error_codes = []
# for error_code in error_codes:
#     if error_code:
#         sub_error_codes.append(error_code[0:8])

# ticket_number = "01234567892321"
# sql = ('''
#     select t.TicketNumber,t.IssueDate,t.Comm,t.QCComm,t.QCStatus,qc.OPStatus,qc.OPComm,qc.AGStatus,qc.AGComm,iar.AuditorStatus,iar.Commission from Ticket t
# left join TicketQC qc
# on t.Id=qc.TicketId
# left join IarUpdate iar
# on t.id=iar.TicketId
# where TicketNumber like ''' + "'" + ticket_number + '''%'
# order by t.IssueDate
#     ''')
#
# print sql

for i in range(0, 3):
    print i
    result = False
    if i == 0:
        print "a"
        continue

    print "b"
    result = True
    break

print "over"
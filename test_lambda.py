agent_codes = []


list_data=[]
data_first={}
data_first['arc']="a"
data_first['status']=0
list_data.append(data_first)
data_second={}
data_second['arc']="b"
data_second['status']=0
list_data.append(data_second)
print "init"
print list_data
list_filetr=filter(lambda x:x['arc']=="b", list_data)
print "filetr"
print list_filetr
print list_data
for data in list_filetr:
    data['status']=1
print "update"
print list_filetr
print list_data
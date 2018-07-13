import re
import arc
import datetime
import ConfigParser

def read_file(file_name):
    with open(file_name, 'r') as f:
        return f.read()


def regex(html):
    pattern = re.compile(r'''        <td width="7%" align="center">(\d{3})</td>
        <td width="11%" align="left">
        
        
            <a href="/IAR/modifyTran\.do\?seqNum=(\d{10})&amp;documentNumber=(\d{10})">\d{10}</a>
            
                
        
		</td>
        <td width="4%" align="right" >(.*?) 
        </td>''')

    return pattern.findall(html)

# arc_regex = arc.Regex()
#
# html = read_file('html/modifyTran_modifyTran.html')
# print arc_regex.get_total(html)[0]


conf = ConfigParser.ConfigParser()
conf.read('../iar_update.conf')
section = "arc"
for option in conf.options(section):
    arc_numbers = conf.get(section, option).split(',')
    account_id = "muling-"
    if option == "all":
        account_id = conf.get("accounts", "all")
    else:
        account_id = account_id + option

    print account_id
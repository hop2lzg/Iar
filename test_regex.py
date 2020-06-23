import re
import arc
import datetime
import ConfigParser

def read_file(file_name):
    with open(file_name, 'r') as f:
        return f.read()


def regex(html):
    pattern = re.compile(r'<input type="text" name=".+?FormOfPayment".+?value="(.*?)"', re.IGNORECASE)
    m = pattern.search(html)
    print m.group(1)
    # return pattern.findall(html)


arc_regex = arc.Regex()

html = read_file('html/listTransactions1.do')
# print html
print arc_regex.modify_trans(html)

# regex(html)
# groups = regex(html)
# print groups[0]
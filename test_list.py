import ConfigParser
import arc
import logging
import logging.config


# li=[100,200,300,400]
# print li
# if 500 in li:
#     index=li.index(200)
#     li[index] = li[0]
#     li[0]=200
# else:
#     li.insert(0, 500)
#
# for l in li:
#     print l

# arc_regex = arc.Regex()
# conf = ConfigParser.ConfigParser()
# conf.read('../iar_update.conf')
# agent_codes = conf.get("certificate", "agentCodes").split(',')
# print agent_codes


def test(is_check_payment, certificates, certificate, agent_codes, tour_code, qc_tour_code, is_et_button=False,
         is_check_update=False):
    certificateItems = []
    if certificates:
        for i in certificates:
            if i[1] not in agent_codes:
                certificateItems.append(i[1])

    certificateItem_len = len(certificateItems)
    for i in range(0, 3):
        if i >= certificateItem_len:
            certificateItems.append("")

    is_check_payment = False
    if is_check_payment:
        certificate = ""

    if not certificate and len(certificateItems) == 3:
        certificateItems.insert(3, "")
    elif certificate and len(certificateItems) == 3:
        certificateItems.insert(0, certificate)

    values = {
        'navButton2.x': "63",
        'navButton2.y': "18",
        'certificateItem[0].value': certificateItems[0],
        'certificateItem[1].value': certificateItems[1],
        'certificateItem[2].value': certificateItems[2],
        'certificateItem[3].value': certificateItems[3],
        'error22010': "false",
        'oldDocumentAirlineCodeFI': "",
        'oldDocumentNumberFI': "",
        # 'maskedFC': maskedFC
        # 'ETButton.x':"27",
        # 'ETButton.y':"7"
    }

    if is_et_button or (not is_check_update and tour_code == qc_tour_code):
        del values['navButton2.x']
        del values['navButton2.y']
        values['ETButton.x'] = "27"
        values['ETButton.y'] = "7"

    print values


# modify_html = ""
# file_object = open('modifyTran.html')
# try:
#     modify_html = file_object.read()
# finally:
#     file_object.close()

# token, maskedFC, arc_commission, waiverCode, certificates = arc_regex.modifyTran(modify_html)
# print certificates
# test(False, certificates, "MJ", agent_codes, "IT", "IT", is_check_update=False)
# certificateItems = []
# certificateItems.append("x")
# certificateItems.append("y")
# certificateItems.append("Z")
# certificateItems.append("A")
# print "init"
# print certificateItems
# certificate_length = len(certificateItems)
# for i in range(0, 3):
#     if i >= certificate_length:
#         certificateItems.append("")
#
# print "append"
# print certificateItems
# certificate = "MJ"
# is_check_payment = False
# if is_check_payment:
#     certificate = ""
#
# if not certificate and len(certificateItems) == 3:
#     certificateItems.insert(3, "")
# elif certificate and len(certificateItems) == 3:
#     certificateItems.insert(0, certificate)
#
# print "last"
# print certificateItems
# logging.info("hi %s,agent %d" % ("zg", 10))
str = "1.l";  # Only digit in this string
f = 0
try:
    f = float(str)
except Exception as e:
    print e
    f = 0

print f

string = " abc "
print "start:>>>%s<<<" % string.strip()

post = {}
post['OPUser'] = "GA"
arc_commission_float = 2
commission_float = 3
AGStatus = 1
if post['OPUser'] != "GW" and AGStatus != 3 and arc_commission_float > commission_float:
    print "Yes"

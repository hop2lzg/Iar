import re
import arc

def read_file(file_name):
    with open(file_name, 'r') as f:
        return f.read()


def regex(html, entry_date):
    pattern = re.compile(r'''<tr align=right class="row[01]"> 
              <td width="3%"  align="center"> .+? 
              </td>
              <td width="6%" align="center">
    				 <!-- defect #810-->
    				 <input type="checkbox" name="checkBoxes" value='(\d{10}):(\d{10})'  />  
              <td  width="5%" align="center">E 
              </td>
              <td width="6%"align="center">(\d{3}) 
              </td>
            <td width="12%" align="left">

                <a href="/IAR/modifyTran\.do\?seqNum=.+?</a>


    		</td>
              <td width="14%" align="left">.+?      
              </td>         
              <td width="21%" align="left">.+? 
              </td>
              <td width="16%" align="left">(NL-PAC) 
              </td>
              <td width="10%"  align="center">(''' + entry_date + ''') 
              </td>
              <td width="10%"  align="center" nowrap>.+? 
              </td>
            </tr>''')

    return pattern.findall(html)


html = read_file('search_05520502.html')
# print html

arc_regex = arc.Regex()

certificates = arc_regex.search_error(html, "30JAN18", "M[J12]|DUP|AG-AGREE")
print certificates
# agent_codes = "MJ,M1,M2,DUP".split(',')
# is_check_payment = False
# print agent_codes
# certificate = "MJ"
# certificateItems = []
# if certificates:
#     for i in certificates:
#         if i[1] not in agent_codes:
#             certificateItems.append(i[1])
#
# certificateItem_len = len(certificateItems)
# for i in range(0, 3):
#     if i >= certificateItem_len:
#         certificateItems.append("")
#
# if is_check_payment:
#     certificate = ""
#
# if not certificate and len(certificateItems) == 3:
#     certificateItems.insert(3, "")
# elif certificate and len(certificateItems) == 3:
#     certificateItems.insert(0, certificate)
#
# print certificateItems
#
# url = "https://iar2.arccorp.com/IAR/financialDetails.do"
# values = {
#     'org.apache.struts.taglib.html.TOKEN': "token",
#     'navButton2.x': "63",
#     'navButton2.y': "18",
#     'amountCommission': "0.00",
#     'miscSupportTypeId': "",
#     'waiverCode': "waiverCode",
#     'certificateItem[0].value': certificateItems[0],
#     'certificateItem[1].value': certificateItems[1],
#     'certificateItem[2].value': certificateItems[2],
#     'certificateItem[3].value': certificateItems[3],
#     'error22010': "false",
#     'oldDocumentAirlineCodeFI': "",
#     'oldDocumentNumberFI': "",
#     'maskedFC': "maskedFC"
#     # 'ETButton.x':"27",
#     # 'ETButton.y':"7"
# }
#
# print values

# if result:
#     commission = result[0]
#     if commission == "":
#         print "empty"
#     else:
#         print "comm:%s|" % commission

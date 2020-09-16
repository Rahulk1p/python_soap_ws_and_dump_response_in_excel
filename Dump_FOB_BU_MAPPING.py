import xlsxwriter
from zeep import Client

wsdl = "http://osbpd.na.lzb.hq:80/FOBManager/ProxyService/FOBManager?wsdl"
client = Client(wsdl)
"""
You can also pass user autherntication details(username and passowrd) in case the wsdl is password protected. For this, you’ll need to create the Session object as shown below:
from zeep import Client
from zeep.transports import Transport
from requests import Session
from requests.auth import HTTPBasicAuth
wsdl = <wsdl_url>
session = Session()
session.auth = HTTPBasicAuth(<username>, <password>)
"""

row = 2;
workbook = xlsxwriter.Workbook('C:/Users/rkumar/fobbumapping.xlsx')
worksheet = workbook.add_worksheet()
response=client.service.getFOBBusinessUnitMappings()
worksheet.write(1, 1, "CROSS REFERENCE TYPE")
worksheet.write(1, 2, "BUSINESS UNIT")
worksheet.write(1, 3, "DESCRIPTION")
worksheet.write(1, 4, "FOB CODE")

#print(response.GetFOBBusinessUnitMappingsResults)
for item in response.GetFOBBusinessUnitMappingsResults:
    worksheet.write(row, 1, item["FOBType"])
    worksheet.write(row, 2, item["BusinessUnit"])
    worksheet.write(row, 3, item["FOBDescription"])
    worksheet.write(row, 4, item["FOBCode"])
    row = row + 1

workbook.close()
import netmiko
import openpyxl
#
# workbooklist=['switch 1 Ports.xlsx','Switch 2 ports .xlsx']
# for num in workbooklist:
#     workbook = openpyxl.load_workbook(filename=num)
#     sheet= workbook['Sheet1']
#     for row in sheet.iter_rows(min_row=2,min_col=2):
#         for i in row:
#             print(i.value)
#

try:
    for i in range(4):
        for j in range(4):
            mycon=netmiko.ConnectHandler(ip='192.168.60.128',port=5000,username='admin',password='cisco',device_type='cisco_ios',secret='cisco')
            mycon.enable()
            output=mycon.send_command('show running-config')
            print(output)
            mycon.disconnect()
except Exception as error:
    print(error)
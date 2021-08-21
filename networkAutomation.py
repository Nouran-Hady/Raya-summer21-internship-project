import xlrd
import initConfig as init

file = xlrd.open_workbook("Routers_data.xlsx")

sheets = file.sheet_by_name("Sheet1")

outputFile1 = open("router1ConfigText.txt", "w")
outputFile2 = open("router2ConfigText.txt", "w")

firstTime1 = 0
firstTime2 = 0
index = 1
while index < sheets.nrows:
    if sheets.row_values(index)[0] == 'R1':
        outFile = outputFile1
    else:
        outFile = outputFile2
    if not sheets.row_values(index)[0] == "":
        if firstTime1 == 0 or firstTime2 == 0:
            outFile.write(f'''
en
conf t
hostname {sheets.row_values(index)[0]}
enable secret {sheets.row_values(index)[3]}
ip domain-name local
crypto key generate rsa
2048
ip ssh version 2
username {sheets.row_values(index)[1]} password {sheets.row_values(index)[2]}
line vty 0 4
transport input telnet ssh
login local
''')
            if sheets.row_values(index)[0] == 'R1':
                firstTime1 += 1
            else:
                firstTime2 += 1
        outFile.write(f'''
interface ethernet {int(sheets.row_values(index)[4])}/0
ip address {sheets.row_values(index)[5].split('/')[0]} 255.255.255.0
''')
    index += 1

outputFile1.write(f'''     
no shutdown
exit
''')

outputFile2.write(f'''     
no shutdown
exit
''')

outputFile1.close()
outputFile2.close()

con = init.open_connection('COM2')
if con:
    init.check_init_dialog(con)
    init.run_command(con)
    with open('router1ConfigText.txt','r+') as file:
        for cmd in file:
            init.run_command(con, cmd)

output = init.read_from_console(con)
print(output)


con = init.open_connection('COM3')
if con:
    init.check_init_dialog(con)
    init.run_command(con)
    with open('router2ConfigText.txt','r+') as file:
        for cmd in file:
            init.run_command(con, cmd)

output = init.read_from_console(con)
print(output)
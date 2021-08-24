import serial, time, xlrd
def configure(name,user_name,user_pass,ip_address,index):
    list = ['COM2', 'COM1']
    with serial.Serial(port=list[index]) as console:
        if console.isOpen():
            print("Serial connection successfully")
            console.write(b'enable\n')
            time.sleep(2)
            console.write(b'configure terminal\n')
            time.sleep(2)
            console.write(b'enable secret cisco\n')
            time.sleep(2)
            console.write(b'hostname '+name.encode()+b'\n')
            time.sleep(2)
            console.write(b'ip domain-name local\n')
            time.sleep(2)
            console.write(b'crypto key generate rsa\n')
            time.sleep(2)
            console.write(b'2048\n')
            time.sleep(2)
            console.write(b'ip ssh version 2\n')
            time.sleep(2)
            console.write(b'username '+user_name.encode()+b' password '+user_pass.encode()+b'\n')
            time.sleep(2)
            console.write(b'line vty 0 4\n')
            time.sleep(2)
            console.write(b'transport input telnet ssh\n')
            time.sleep(2)
            console.write(b'login local\n')
            time.sleep(2)
            console.write(b'exit\n')
            time.sleep(2)
            console.write(b'interface VLAN 1\n')
            time.sleep(2)
            console.write(b'ip address ' +ip_address.encode()+ b' 255.255.255.0\n')
            time.sleep(2)
            console.write(b'no shutdown\n')
            time.sleep(2)
            console.write(b'exit\n')
            time.sleep(2)
            console.write(b'exit\n')
            time.sleep(10)
        else:
            print('Sorry')


workbook = xlrd.open_workbook('Switches_SSH-data.xlsx')
worksheet = workbook.sheet_by_name('Sheet1')
ind=0
i=1


while i < worksheet.nrows:
    data = worksheet.row_values(i)
    #print(data)
    back_slash = data[5].find('/')
    ip_address = data[5][0:back_slash]
    configure(data[0],data[1],data[2],ip_address,ind)
    ind+=1
    i+= 1
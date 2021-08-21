import serial, time, xlrd
def configure(name,ip_address,index):
    list = ['COM3', 'COM4']
    with serial.Serial(port=list[index]) as console:
        if console.isOpen():
            print("Serial connection successfully")
            console.write(b'enable\n')
            time.sleep(2)
            console.write(b'configure terminal\n')
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
            time.sleep(1)
            console.write(b'username admin password cisco\n')
            time.sleep(1)
            console.write(b'interface VLAN 1\n')
            time.sleep(1)
            console.write(b'ip address ' +ip_address.encode()+ b' 255.255.255.0\n')
            time.sleep(1)
            console.write(b'exit\n')
            time.sleep(1)
            console.write(b'line vty 0 4\n')
            time.sleep(1)
            console.write(b'transport input telnet ssh\n')
            time.sleep(1)
            console.write(b'login local\n')
            time.sleep(1)
            console.write(b'exit\n')
            time.sleep(1)
            console.write(b'exit\n')
            time.sleep(1)
            console.write(b'en\n')  # puts the string in binary representation
            time.sleep(1)
            console.write(b'terminal length 0\n')
            time.sleep(1)
            console.write(b'show ip int br\n')
            time.sleep(1)  # suspend execution
            numberofbytes=console.inWaiting() #calculates the number of bytes that will be shown
            data=console.read(numberofbytes)
            with open ("configuration.txt",'a') as file:
                file.write(data.decode()+'\n') #decode the data from binary to characters
            time.sleep(10)
        else:
            print('Sorry')


workbook = xlrd.open_workbook('Switches_SSH-data.xlsx')
worksheet = workbook.sheet_by_name('Sheet1')
list=[]
dict1 = {}
names=['SW1','SW2']
ind=0
i=1
j=0

while i < worksheet.nrows:
    whole = worksheet.row_values(i)[5]
    back_slash = whole.find('/')
    ip_address =whole[0:back_slash]
    list.append(ip_address)
    i+= 1

for k in names:
    dict1[k] = list[j]
    j+=1
#dict1={'SW1': 192.168.10.5, 'SW2': 10.1.1.5}
for key, value in dict1.items():
    configure(key,value,ind)
    ind+=1
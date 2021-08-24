import serial,time,xlrd,openpyxl,os

#########################
# Nihal and Hadeer#
#########################

def open_connection(port):
    console = serial.Serial(port=port)
    if console.isOpen():
        print("serial connection sucessfully")
        console.write(b"en \n")
        time.sleep(1)
        console.write(b"terminal length 0\n")
        time.sleep(1)
        console.write(b"show ip int br\n")
        time.sleep(2)
        return console
    else:
        print("Sorry! you cannot connect")
        return 0

def read_from_console(console):
    numBytes = console.inWaiting()
    print("num",numBytes)
    if numBytes:
        data = console.read(numBytes)
        return data.decode()
    else:
        return False

def run_command(console, command='\n',sleep=3):
    print("Sending Command: " + command)
    console.write(command.encode())
    time.sleep(sleep)

def check_init_dialog(console):
    run_command(console)
    output = read_from_console(console)
    print("out",output)
    if output and 'initial configuration dialog?' in output :
        run_command(console,'no')
        run_command(console,'\n',15)
        run_command(console, '\r\n')
        return True
    else:
        return False

#########################
# Eman#
#########################
# import xlrd
def configure(name,user_name,user_pass,ip_address,index):
    list = ['COM2', 'COM1']
    Ips=['192.168.10.1','10.1.1.1']
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
            console.write(b'ip default-gateway '+Ips[index].encode()+b'\n')
            time.sleep(5)
        else:
            print('Sorry')

#########################
# Nouran#
#########################



def switch_vlan(port,vlan_sheet):
    filelist = []
    interfaces = []
    porttype = []
    vlans = []
    file=open('switch port type.txt','a+')
    workbook = openpyxl.load_workbook(filename=vlan_sheet)
    sheet= workbook['Sheet1']
    for row in sheet.iter_rows(min_row=2,min_col=2):
        for i in row:
            file.write(str(i.value)+'\n')
    file.close()

    try:
            file = open('switch port type.txt', 'r')
            for i in file.readlines():
                filelist.append(i.strip('\n'))

            for i in range(3):
                for j in range(4):
                    interfaces.append('interface ethernet ' + f'{i}/{j}')

            for i in range(len(filelist)):
                if i % 2 == 0:
                    porttype.append(filelist[i])
                else:
                    vlans.append(filelist[i])
            interfaces.pop(11)
            interfaces.pop(10)
            interfaces.pop(9)
            file.close()

            with serial.Serial(port=port) as switch:
                i = 0
                for index in porttype:

                    if index.lower() == 'access':
                        vlan =f'vlan {vlans[i]}'

                        switch.write(vlan.encode()+b'\n')
                        time.sleep(3)
                        switch.write(b'exit\n')
                        time.sleep(2)
                        ethernet = interfaces[i]

                        switch.write(ethernet.encode()+b'\n')
                        time.sleep(2)

                        switch.write(b'switchport mode access\n')
                        time.sleep(2)

                        vlan=f'switchport access vlan {vlans[i]}'
                        switch.write(vlan.encode()+b'\n')
                        time.sleep(2)
                        switch.write(b'no shut\n')
                        time.sleep(2)
                        switch.write(b'exit\n')
                        i = i + 1

                    else:
                        ethernet = interfaces[i]

                        switch.write(ethernet.encode()+b'\n')
                        time.sleep(2)

                        switch.write(b'switchport trunk encapsulation dot1q\n')
                        time.sleep(2)
                        switch.write(b'switchport mode trunk\n')
                        time.sleep(2)
                        vlan=f'switchport trunk allowed Vlan {vlans[i]}'
                        switch.write(vlan.encode()+b'\n')
                        time.sleep(2)
                        switch.write(b'no shut\n')
                        time.sleep(2)
                        switch.write(b'exit\n')
                        i = i + 1


    except Exception as error:
            print(error)
    os.remove(r'switch port type.txt')

def ip_route():
    coms=['COM4','COM5']
    routes=['10.1.1.0 255.255.255.0 192.168.11.2','192.168.10.0  255.255.255.0  192.168.11.1']
    for com in range(2):
        with serial.Serial(port=coms[com]) as rout:
            rout.write(b'ip route '+routes[com].encode()+b'\n')
            time.sleep(2)
            rout.write(b'exit\n')
            time.sleep(2)
            rout.write(b'show ip route\n')
            time.sleep(2)
            print('static route number',com+1)
            numBytes = rout.inWaiting()
            data = rout.read(numBytes)
            print(data.decode())
            print("##########################")



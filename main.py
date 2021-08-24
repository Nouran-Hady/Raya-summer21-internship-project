import serial,time
import openpyxl,os
import initConfig as init
#
# filelist=[]
# interfaces=[]
# porttype=[]
# vlans=[]
#
# def switch_vlan(port,vlan_sheet):
#     file=open('switch port type.txt','a+')
#     workbook = openpyxl.load_workbook(filename=vlan_sheet)
#     sheet= workbook['Sheet1']
#     for row in sheet.iter_rows(min_row=2,min_col=2):
#         for i in row:
#             file.write(str(i.value)+'\n')
#     file.close()
#
#     try:
#             file = open('switch port type.txt', 'r')
#             for i in file.readlines():
#                 filelist.append(i.strip('\n'))
#
#             for i in range(3):
#                 for j in range(4):
#                     interfaces.append('interface ethernet ' + f'{i}/{j}')
#
#             for i in range(len(filelist)):
#                 if i % 2 == 0:
#                     porttype.append(filelist[i])
#                 else:
#                     vlans.append(filelist[i])
#             interfaces.pop(11)
#             interfaces.pop(10)
#             interfaces.pop(9)
#             file.close()
#
#             with serial.Serial(port=port) as switch:
#                 i = 0
#                 for index in porttype:
#
#                     if index.lower() == 'access':
#                         vlan =f'vlan {vlans[i]}'
#
#                         switch.write(vlan.encode()+b'\n')
#                         time.sleep(3)
#                         switch.write(b'exit\n')
#                         time.sleep(2)
#                         ethernet = interfaces[i]
#
#                         switch.write(ethernet.encode()+b'\n')
#                         time.sleep(2)
#
#                         switch.write(b'switchport mode access\n')
#                         time.sleep(2)
#                         switch.write(b'no shut\n')
#                         time.sleep(2)
#                         vlan=f'switchport access vlan {vlans[i]}'
#                         switch.write(vlan.encode()+b'\n')
#                         time.sleep(2)
#                         switch.write(b'exit\n')
#                         i = i + 1
#
#                     else:
#                         ethernet = interfaces[i]
#
#                         switch.write(ethernet.encode()+b'\n')
#                         time.sleep(2)
#                         switch.write(b'no shut\n')
#                         time.sleep(2)
#                         switch.write(b'switchport trunk encapsulation dot1q\n')
#                         time.sleep(2)
#                         switch.write(b'switchport mode trunk\n')
#                         time.sleep(2)
#                         vlan=f'switchport trunk allowed Vlan {vlans[i]}'
#                         switch.write(vlan.encode()+b'\n')
#                         time.sleep(2)
#                         switch.write(b'exit\n')
#                         i = i + 1
#
#                 numberofbytes = switch.inWaiting()
#                 data = switch.read(numberofbytes)
#                 print(data.decode())
#
#     except Exception as error:
#             print(error)
#     os.remove(r'D:\raya summer training\raya project\pythonProject\switch port type.txt')
#
#
#
# switch_vlan('COM2','switch 1 Ports.xlsx')
# switch_vlan('COM1','Switch 2 Ports .xlsx')

def reload(com):
    try:
        print("reloading...")
        with serial.Serial(port=com) as device:
            device.write(b"en\n")
            time.sleep(1)
            numBytes = device.inWaiting()
            data = device.read(numBytes)
            if 'password' in str(data.decode()).lower():
                device.write(b"cisco\n")
                time.sleep(1)
                print("password cisco")

            device.write(b"write erase\n")
            time.sleep(1)
            device.write(b"\n")
            time.sleep(1)
            device.write(b"reload\n")
            time.sleep(1)
            device.write(b"\n")
            time.sleep(20)
            numBytes = device.inWaiting()
            data = device.read(numBytes)
            if 'save?' in data.decode():
                device.write(b'no\n')
                device.write(b'\n')
                time.sleep(2)
            numBytes = device.inWaiting()
            data = device.read(numBytes)
            if 'initial configuration dialog?' in data.decode():
                    init.run_command(device, 'no')
            numBytes = device.inWaiting()
            data = device.read(numBytes)

            if 'terminate autoinstall? [yes]:' in data.decode():
                device.write(b'yes\n')
                device.write(b'\n')
                time.sleep(2)



    except Exception as err:
        print(err)

reload('COM4')
reload('COM5')








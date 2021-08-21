import serial,time


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
    console.write(command.encode() + b'\n')
    time.sleep(sleep)

def check_init_dialog(console):
    run_command(console)
    output = read_from_console(console)
    print("out",output)
    if 'initial configuration dialog?' in output:
        run_command(console,'no')
        run_command(console,'\n',15)
        run_command(console, '\r\n')
        return True
    else:
        return False


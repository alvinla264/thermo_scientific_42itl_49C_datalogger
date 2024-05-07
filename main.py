import serial.tools.list_ports
import serial.serialutil
import serial.serialwin32
from datetime import datetime as dt
def get_com_ports():
    com_ports = []
    ports = serial.tools.list_ports.comports()
    for port, desc, _ in ports:
        com_ports.append(port)
    return com_ports

def open_serial_port(com_port, baudrate):
    try:
        # Open the serial port
        ser = serial.serialwin32.Serial(com_port, baudrate)
        print(f"Serial port {com_port} opened successfully.")
        return ser
    except serial.serialutil.SerialException as e:
        print(f"Failed to open serial port {com_port}: {e}")
        return None
# Example usage:
if __name__ == "__main__":
    com_list = get_com_ports()
    instrument_serial_info = {"42i-TL" : {"com_port": "",
                                          "baudrate" : 0,
                                          "avaliable_baudrate": [1200, 2400 , 4800, 9600 , 19200 , 38400 , 57600 , 115200]},
                              "49C":  {"com_port": "",
                                          "baudrate" : 0,
                                          "avaliable_baudrate" : [1200, 2400 , 4800, 9600]}}
    for instrument in instrument_serial_info.keys():
        valid_com_port = False
        valid_baudrate = False
        while not valid_com_port:
            print("\nList of available COM ports:")
            for port in com_list:
                print(port)
            com_input = input(f"Select COM port for {instrument}: ").upper()
            if com_input not in com_list:
                print("unknown COM port\n")
            else:
                valid_com_port = True
        while not valid_baudrate:
            print("\nList of available Baudrate:")
            for baudrate in instrument_serial_info[instrument]["avaliable_baudrate"]:
                print(baudrate)
            baurdrate_input = input(f"Select Baudrate for {instrument}: ")
            if baurdrate_input.isdigit():
                baurdrate_input = int(baurdrate_input)
                if baurdrate_input not in instrument_serial_info[instrument]["avaliable_baudrate"]:
                    print("Unknown baudrate")
                else:
                    valid_baudrate = True
            else:
                print("Invalid Input")
        instrument_serial_info[instrument]["com_port"] = com_input
        instrument_serial_info[instrument]["baudrate"] = baurdrate_input
    print()
    for instrument in instrument_serial_info:
        print(instrument + ":", instrument_serial_info[instrument]["com_port"], instrument_serial_info[instrument]['baudrate'])
    ser_42i = open_serial_port(instrument_serial_info["42i-TL"]["com_port"], instrument_serial_info["42i-TL"]["baudrate"])
    if ser_42i:
        print("Success")
        ser_42i.write(b"Hello World\n")
    else:
        print("Fail")
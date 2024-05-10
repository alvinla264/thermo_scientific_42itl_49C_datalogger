'''
File: serial_func.py
Author: Alvin La
Description: This python script contains functions that gets available com ports for your system and opens serial ports
'''
import serial.tools.list_ports
import serial.serialutil
import serial.serialwin32


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
    
def get_data(ser, compound_list, id):
    results = {}
    for compound in compound_list:
        ser.write(f'{id}{compound}\r'.encode('utf-8'))
        response = ser.read_until(b'\r').decode()
        if response.find('\n') != -1:
            response = response[:response.index('\n') - 1]
        response = response.split()
        #print(response, len(response))

        if len(response) == 3:
            results[response[0]] = {"value" : float(response[1]),
                                "unit": response[2]}
    return results 


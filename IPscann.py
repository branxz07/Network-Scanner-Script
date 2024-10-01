import subprocess
import re
import socket
import requests
import time
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.utils.cell import get_column_letter
import platform
import time

def get_local_ip():
    hostname = socket.gethostname()
    local_ip = socket.gethostbyname(hostname)
    return local_ip

def get_hostname(ip):
    try:
        hostname = socket.gethostbyaddr(ip)[0]
    except socket.herror:
        hostname = None
    return hostname

def get_mac_vendor(mac):
    try:
        url = f'https://api.macvendors.com/{mac}'
        response = requests.get(url)
        if response.status_code == 200:
            return response.text
        else:
            return 'Unknown'
    except Exception as e:
        return 'Unknown'

def scan_network_with_arp():
    command = 'arp -a'
    result = subprocess.run(command, capture_output=True, text=True, shell=True)
    devices = []
    for line in result.stdout.split('\n'):
        if re.match(r'^\s*\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}', line):  # Match any IP address
            parts = line.split()
            ip = parts[0]
            mac = parts[1]
            devices.append({'ip': ip, 'mac': mac})
    return devices



def log_devices(devices, log_file):
    header = ["Time", "IP", "MAC", "Hostname", "Vendor"]
    try:
        # Load the existing workbook
        workbook = openpyxl.load_workbook(log_file)
        sheet = workbook.active
    except FileNotFoundError:
        # Create a new workbook if the file doesn't exist
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        # Set the header row
        sheet.append(header)
    
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    for device in devices:
        hostname = get_hostname(device['ip'])
        vendor = get_mac_vendor(device['mac'])
        row = [timestamp, device['ip'], device['mac'], hostname if hostname else 'Unknown', vendor if vendor else 'Unknown']
        sheet.append(row)
    
    # Adjust column widths
    columns = list(sheet.columns)
    for col_idx, col in enumerate(columns, start=1):
        max_length = max(len(str(cell.value)) for cell in col)
        sheet.column_dimensions[get_column_letter(col_idx)].width = max_length + 2
    
    # Add filter to IP column
    sheet.auto_filter.ref = 'B1:E' + str(sheet.max_row)
    
    # Retry saving the workbook if PermissionError occurs
    while True:
        try:
            workbook.save(log_file)
            print(f"Information logged to {log_file}")
            break
        except PermissionError:
            print(f"Permission denied: {log_file} is open. Retrying in 5 seconds...")
            time.sleep(5)

if __name__ == '__main__':
    print("You are currently running the Scann Script on: " + platform.system())
    if platform.system() == 'Linux':
        print('Hi linux')
    elif platform.system() == 'Windows':
        print('Hi Windows')
    else:
        print('This script does not support MAC or any other OS other than Linux and Windows 10/11')
    # Create a log file name with the current date
    log_file = f"network-log-{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    
    while True:
        local_ip = get_local_ip()
        
        print(f"Scanning network using ARP table")
        devices = scan_network_with_arp()
        
        print("Devices found:")
        for device in devices:
            hostname = get_hostname(device['ip'])
            vendor = get_mac_vendor(device['mac'])
            info = f"IP: {device['ip']}, MAC: {device['mac']}, Hostname: {hostname if hostname else 'Unknown'}, Vendor: {vendor if vendor else 'Unknown'}" #cambiar vendedor por fabricante y checar mis dispositivos para guardar los que se que son mios
            print(info)
        
        log_devices(devices, log_file)
        
        print("Waiting till next Scann")
        
        time.sleep(60)  # Wait for 60 seconds before the next scan
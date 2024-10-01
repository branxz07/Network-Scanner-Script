"""
Network Scanner Script using ARP Table.

This script scans the network and logs details
about the devices on the network, including their IP, MAC address,
hostname, and vendor/fabricant. The information is saved in an Excel file, with
columns for time, IP, MAC, hostname, and vendor.

Creator: Brandon Magaña Avalos
"""

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


def get_local_ip():
    """
    Get the local IP address of the current machine.

    Returns:
        str: The local IP address.
    """
    hostname = socket.gethostname()
    local_ip = socket.gethostbyname(hostname)
    return local_ip


def get_hostname(ip):
    """
    Get the hostname of a device by its IP address.

    Args:
        ip (str): The IP address of the device.

    Returns:
        str: The hostname if available, otherwise None.
    """
    try:
        hostname = socket.gethostbyaddr(ip)[0]
    except socket.herror:
        hostname = None
    return hostname


def get_mac_vendor(mac):
    """
    Get the vendor of a device using its MAC address.

    Args:
        mac (str): The MAC address of the device.

    Returns:
        str: The vendor name if available, otherwise 'Unknown'.
    """
    try:
        url = f'https://api.macvendors.com/{mac}'
        response = requests.get(url)
        if response.status_code == 200:
            return response.text
        else:
            return 'Unknown'
    except Exception:
        return 'Unknown'


def scan_network_with_arp():
    """
    Scan the local network using the ARP table.

    Uses the ARP table to gather IP and MAC addresses of devices on the
    network.

    Returns:
        list[dict]: A list of devices with their IP and MAC addresses.
    """
    command = 'arp -a'
    result = subprocess.run(command, capture_output=True, text=True, shell=True)
    devices = []
    
    # Parse the ARP table output
    for line in result.stdout.split('\n'):
        if re.match(r'^\s*\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}', line):
            parts = line.split()
            ip = parts[0]
            mac = parts[1]
            devices.append({'ip': ip, 'mac': mac})
    
    return devices


def log_devices(devices, log_file):
    """
    Log device information to an Excel file.

    Args:
        devices (list[dict]): A list of devices containing IP, MAC addresses.
        log_file (str): The file path to log the data.

    This function logs the timestamp, IP, MAC address, hostname, and vendor
    of each device to the specified Excel file.
    """
    header = ["Time", "IP", "MAC", "Hostname", "Vendor"]
    
    try:
        # Load the existing workbook
        workbook = openpyxl.load_workbook(log_file)
        sheet = workbook.active
    except FileNotFoundError:
        # Create a new workbook if the file doesn't exist
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(header)  # Set the header row
    
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    # Append device information to the Excel sheet
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
    
    # Retry saving the workbook in case of permission error
    while True:
        try:
            workbook.save(log_file)
            print(f"Information logged to {log_file}")
            break
        except PermissionError:
            print(f"Permission denied: {log_file} is open. Retrying in 5 seconds...")
            time.sleep(5)


if __name__ == '__main__':
    print("You are currently running the Scan Script on: " + platform.system())

    # Create a way to work for Linux and windows
    if platform.system() == 'Linux':
        print('Hi Linux')
    elif platform.system() == 'Windows':
        print('Hi Windows')
    else:
        print('This script does not support Mac or any OS other than Linux and Windows 10/11')
    
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
            info = f"IP: {device['ip']}, MAC: {device['mac']}, Hostname: {hostname if hostname else 'Unknown'}, Vendor: {vendor if vendor else 'Unknown'}"
            print(info)
        
        log_devices(devices, log_file)
        
        print("Waiting till next scan")
        time.sleep(60)  # Wait for 60 seconds before the next scan

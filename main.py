__author__ = 'LandrCod'
"""
PyPLC IP Address Grabber

Generates an Excel File with sheets for each PLC project in a directory the user selects.
"""

import xmltodict
import xlwt
import easygui
import os
from collections import defaultdict

# Global to hold all the ip addresses
third_octets = defaultdict(list)

def get_plc_ip_addresses(file_path, filename):
    """
    Takes a file path for a RSLogix 5000 Project in L5X format.  Generates a list
    of all IP Addresses for all hardware in the project IO Configuration Tree.
    Returns a list of lines in CSV Format that can be written to a CSV file or
    into an Excel Sheet.
    """
    global third_octets
    file_lines = []

    # Parse the L5X (XML) file and narrow down to the hardware modules
    with open(file_path) as xml_file:
        modules = xmltodict.parse(xml_file.read())['RSLogix5000Content']['Controller']['Modules']

    # Gather pertinent data for the CSV Columns (name, cat_no, ip_addr)
    for module in modules.values():
        for attr in module:

            if '@Name' in attr.keys():
                name = attr['@Name']
            else:
                name = "No Name"
            if '@CatalogNumber' in attr.keys():
                catalog_no = attr['@CatalogNumber']
            else:
                catalog_no = 'None'

            # The 'Ports' field may return a list or an OrderedDict
            for allports in attr['Ports'].items():
                for ports in allports:

                    # if the device has more than one port, we must iterate through the list, checking
                    # for the ip address on each port
                    if isinstance(ports, list):
                        for port in ports:
                            if port['@Type'] == 'Ethernet':
                                try:
                                    file_lines.append((name, catalog_no, port['@Address']))
                                    ip = port['@Address']
                                    octets_1_to_3 = ip[:ip.rfind('.')]
                                    third_octets[octets_1_to_3].append(ip.split('.')[3])
                                except Exception, e:
                                    print e, name

                    # otherwise, we can just check for the ethernet ip address
                    elif isinstance(ports, dict):
                        if ports['@Type'] == 'Ethernet':
                            try:
                                file_lines.append((name, catalog_no, ports['@Address']))
                                ip = ports['@Address']
                                octets_1_to_3 = ip[:ip.rfind('.')]
                                third_octets[octets_1_to_3].append(ip.split('.')[3])
                            except Exception, e:
                                print e, name

    # return a list of lines in CSV Format
    return file_lines


def find_free_ips(ip_dict):
    """
    Gets all the available IP's excluding those used anywhere in the input dict
    """
    free_ip_dict = {}
    for k, v in ip_dict.items():
        free_ips = [ip for ip in range(1, 255) if str(ip) not in v]
        free_ip_dict[k] = free_ips
    return free_ip_dict

easygui.msgbox("This program will generate an Excel file with all the IP addresses found in a folder you designate.  "
               "Save RSLogix 5000 project files as type L5X and put them all in a folder.  Select this folder in the "
               "following dialog.  When the program is done, select a location for the file to be saved.",
               ok_button="OK", title="IP Address File Generator")


# Get list of L5X files from the user selected directory
directory = easygui.diropenbox('Select your L5X Folder.', default='C:\\')
files = sorted(os.listdir(directory))

# Create a new workbook and setup cell styles
workbook = xlwt.Workbook()
title_style = xlwt.easyxf("font: bold 1")

# Create a new sheet for each L5X file and populate with the name, device type, and ip address of all hardware
for file in files:

    # Create the worksheet
    print 'Processing file:', file
    filename = file[:file.find('.')]
    sheet = workbook.add_sheet(filename)

    sheet.write(0, 0, filename, title_style)

    # Set the title
    for col, title in enumerate(["Name", "Device Type", "IP Address"]):
        sheet.write(1, col, title, title_style)

    # Get the hardware data
    file = directory + '\\' + file
    lines = get_plc_ip_addresses(file, filename)

    # Write rows
    for row, line in enumerate(lines):
        for col_idx, col in enumerate(line):
            sheet.write(row + 2, col_idx, line[col_idx])

    # Set column widths
    for col in range(5):
        sheet.col(col).width = 256 * 30

print "Finding unused IP addresses..."
# Generate Sheet with free IP addresses that aren't currently being used
free_ip_sheet = workbook.add_sheet('Free IP Addresses')
free_ips = find_free_ips(third_octets)

for col, header in enumerate(free_ips.keys()):
    free_ip_sheet.write(0, col, header)
    free_ip_sheet.col(col).width = 256 * 15
    for row, last_octet in enumerate(free_ips[header]):
        free_ip_sheet.write(row + 1, col, last_octet)

# Save the workbook
workbook.save(easygui.filesavebox('Save IP Address File', 'Save File', default='IP_Addresses.xls', filetypes=['*.xls']))

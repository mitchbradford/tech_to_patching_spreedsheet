# -----------------------------------------------------------------------
'''
Script by Mitch Bradford
Dependencies - Python 3, xlwt

Parses files containing one or more show tech of Cisco devices
and extracts port information to an Excel file to assist in creating
port patching spreedsheets for switch replacements.

You can put show tech of as many Cisco devices as you want in one file
or you can have multiple files and use wildcards

usage: python tech_to_patching_spreadsheet.py <Excel output file> <input file 1> <input file 2> etc or *.txt

Credit to Andres Gonzelez for creating the original script
https://github.com/angonz/tech2xl
'''
# -----------------------------------------------------------------------

import re, glob, sys, csv, collections
import xlwt
import time

# To be determined use
def expand(s, list):
    for item in list:
        if len(s) <= len(item):
            if s.lower() == item.lower()[:len(s)]:
                return item
    return None

# To be determined use
def expand_string(s, list):
    result = ''
    for pos, word in enumerate(s.split()):
        expanded_word = expand(word, list[pos])
        if expanded_word is not None:
            result = result + ' ' + expanded_word
        else:
            return None
    return result[1:]

# -----------------------------------------
start_time = time.time()

print("tech_to_patching_spreadsheet v1.0")

# If user hasn't provided sufficent commands when running scripts, prompt
if len(sys.argv) < 3:
    print("Usage: tech_to_patching_spreadsheet <output file> <input files>...")
    sys.exit(2)

commands = [["show"], \
            ["version", "cdp", "technical-support", "running-config", "interfaces", "diag", "inventory"], \
            ["neighbors", "status"], \
            ["detail"]]


int_types = ["Ethernet", "FastEthernet", "GigabitEthernet", "Gigabit", "TenGigabit", "Serial", "ATM", "Port-channel", "Tunnel", "Loopback"]

# Inicialized the collections.OrderedDictionary that will store all the info
systeminfo = collections.OrderedDict()
intinfo = collections.OrderedDict()
cdpinfo = collections.OrderedDict()
diaginfo = collections.OrderedDict()

#These are the fields to be extracted for the Inventory tab
systemfields = ["Name", "Model", "System ID", "Image", "Uptime"]

# Fields for Interfaces tab.  Commenting out excess noise
intfields = ["Patch Panel / Device", \
# change "Name" to switch
            "Name", \
            "Interface", \
            "Number", \
            "Description", \
            "Status", \
            "Line protocol", \
            "Switchport mode", \
            "Access vlan", \
            "Voice vlan", \
			"Native VLAN", \
			"Allowed VLAN", \
             "Last Packet Input", \
            "Duplex", \
            "Speed", \
            "Media type"]   

#diagfields = ["Name", "Slot", "Subslot", "Description", "Serial number", "Part number"]

# takes all arguments starting from 2nd
for arg in sys.argv[2:]:
    # uses glob to consider wildcards
    for file in glob.glob(arg):

        infile = open(file, "r")
        
        # This is the name of the router
        name = ''

        # Identifies the section of the file that is currently being read
        command = ''
        section = ''
        item = ''
        cdp_neighbor = ''

        take_next_line = 0

        for line in infile:

            # checks for device name in prompt
            m = re.search("^([a-zA-Z0-9][a-zA-Z0-9_\-\.]*)[#>]\s*([\w\-\_\s\b\a]*)", line)
            # avoids a false positive in the "show switch detail" or "show flash: all" section of show tech
            if m and not (command == "show switch detail" or command == "show flash: all"):

                if name == '':
                    infile.seek(0)
                else:
                    #removes all deleted chars with backspace (\b) and bell chars (\a)
                    cli = m.group(2)

                    while re.search("\b|\a", cli):
                        cli = re.sub("[^\b]\b|\a", "", cli)
                        cli = re.sub("^\b", "", cli)
                    command = expand_string(cli, commands)

                name = m.group(1)
                section = ''
                item = ''

                if name not in systeminfo.keys():
                    systeminfo[name] = collections.OrderedDict(zip(systemfields, [''] * len(systemfields)))
                    systeminfo[name]['Name'] = name

                if name not in intinfo.keys():
                    intinfo[name] = collections.OrderedDict()

#                if name not in cdpinfo.keys():
#                    cdpinfo[name] = collections.OrderedDict()

                continue

            # detects section within show tech
            m = re.search("^------------------ (.*) ------------------$", line)
            if m:
                command = m.group(1)
                section = ''
                item = ''
                continue

            # processes "show running-config" command or section of sh tech
            if command == 'show running-config':
                # extracts information as per patterns

                m = re.match("hostname ([a-zA-Z0-9][a-zA-Z0-9_\-\.]*)", line)
                if m:
                    if name == '':
                        name = m.group(1)
                        infile.seek(0)

                        section = ''
                        item = ''

                        if name not in systeminfo.keys():
                            systeminfo[name] = collections.OrderedDict(zip(systemfields, [''] * len(systemfields)))
                            systeminfo[name]['Name'] = name

                        if name not in intinfo.keys():
                            intinfo[name] = collections.OrderedDict()

#                        if name not in cdpinfo.keys():
#                            cdpinfo[name] = collections.OrderedDict()

                    continue

                m = re.match("interface (\S*)", line)
                if m:
                    section = 'interface'
                    item = m.group(1)

                    if item not in intinfo[name].keys():
                        intinfo[name][item] = collections.OrderedDict(zip(intfields, [''] * len(intfields)))
                        intinfo[name][item]['Name'] = name
                        intinfo[name][item]['Interface'] = item
                        intinfo[name][item]['Number'] = re.split('\D+', item, 1)[1]
                    continue

                if section == 'interface':

                    if line == '!':
                        section = ''
                        continue

                    m = re.match(" description (.*)", line)
                    if m:
                        intinfo[name][item]['Description'] = m.group(1)
                        continue

                    m = re.match(" switchport mode (\w*)", line)
                    if m:
                        intinfo[name][item]['Switchport mode'] = m.group(1)
                        continue

                    m = re.search(" switchport access vlan (\d+)", line)
                    if m:
                        intinfo[name][item]["Access vlan"] = m.group(1)
                        continue

                    m = re.search(" switchport voice vlan (\d+)", line)
                    if m:
                        intinfo[name][item]["Voice vlan"] = m.group(1)
                        continue

                    m = re.search(" switchport trunk native vlan (\d+)", line)
                    if m:
                        intinfo[name][item]["Native VLAN"] = m.group(1)
                        continue
                   
                    m = re.search(" switchport trunk allowed vlan (.*)", line)
                    if m:
                        intinfo[name][item]["Allowed VLAN"] = m.group(1)
                        continue

                    m = re.search("^ ip address ([\d|\.]+) ([\d|\.]+)", line)
                    if m:
                        intinfo[name][item]['IP address'] = m.group(1)
						
            # processes "show version" command or section of sh tech
            if command == 'show version' and name != '':
                # extracts information as per patterns
                m = re.search("Processor board ID (.*)", line)
                if m:
                    systeminfo[name]['System ID'] = m.group(1)
                    continue

                m = re.search("Model number\s*: (.*)", line)
                if m:
                    systeminfo[name]['Model'] = m.group(1)
                    continue

                m = re.search("^cisco (.*) processor", line)
                if m:
                    systeminfo[name]['Model'] = m.group(1)
                    continue

                m = re.search("^Cisco (.*) \(revision", line)
                if m:
                    systeminfo[name]['Model'] = m.group(1)
                    continue

                m = re.search("Motherboard serial number\s*: (.*)", line)
                if m:
                    systeminfo[name]['Mother ID'] = m.group(1)
                    continue

                m = re.search('System image file is \"flash:\/?(.*)\.bin\"', line)
                if m:
                    systeminfo[name]['Image'] = m.group(1)
                    continue

                m = re.search('System image file is \"flash:\/.*\/(.*)\.bin\"', line)
                if m:
                    systeminfo[name]['Image'] = m.group(1)
                    continue

                m = re.search('System image file is \"bootflash:(.*)\.bin\"', line)
                if m:
                    systeminfo[name]['Image'] = m.group(1)
                    continue

                m = re.search('System image file is \"sup-bootflash:(.*)\.bin\"', line)
                if m:
                    systeminfo[name]['Image'] = m.group(1)
                    continue

                m = re.search(" uptime is (.*)", line)
                if m:
                    systeminfo[name]['Uptime'] = m.group(1)
                    continue
                    
            # processes "show interfaces" command or section of sh tech
            if command == 'show interfaces' and name != '':
                # extracts information as per patterns

                m = re.search("^(\S+) is ([\w|\s]+), line protocol is (\w+)", line)
                if m:
                    item = m.group(1)
                    if item not in intinfo[name].keys():
                        intinfo[name][item] = collections.OrderedDict(zip(intfields, [''] * len(intfields)))
                        intinfo[name][item]['Name'] = name
                        intinfo[name][item]['Interface'] = item

                    intinfo[name][item]['Status'] = m.group(2)
                    intinfo[name][item]['Line protocol'] = m.group(3)
                    continue

                m = re.search("^  Last input (.*)", line)
                if m:
                    intinfo[name][item]['Last Packet Input'] = m.group()
                    continue

                m = re.search("^  Description: (.*)", line)
                if m:
                    intinfo[name][item]['Description'] = m.group(1)
                    continue

                m = re.search("(\w+) Duplex, (\d+)Mbps, link type is (\w+), media type is (.*)", line)
                if m:
                    intinfo[name][item]['Duplex'] = m.group(3) + "-" + m.group(1)
                    intinfo[name][item]['Speed'] = m.group(3) + "-" + m.group(2)
                    intinfo[name][item]['Media type'] = m.group(4)
                    continue

                m = re.search("(\w+)-duplex, (\d+)Mb/s, media type is (.*)", line)
                if m:
                    intinfo[name][item]['Duplex'] = m.group(1)
                    intinfo[name][item]['Speed'] = m.group(2)
                    intinfo[name][item]['Media type'] = m.group(3)
                    continue

            # processes "show interfaces status" command or section of sh tech
            if command == 'show interfaces status' and name != '':
                if (line[:4] != "Port"):
                    item = expand(line[:2], int_types)

                    if item is not None:
                        item = item + line[2:8].rstrip()

                        if item not in intinfo[name].keys():
                            intinfo[name][item] = collections.OrderedDict(zip(intfields, [''] * len(intfields)))
                            intinfo[name][item]['Name'] = name
                            intinfo[name][item]['Interface'] = item
                            intinfo[name][item]['Number'] = re.split('\D+', item, 1)[1]

                    m = re.search("(.+) (connected|notconnect|disabled)\s+(\S+)\s+(\S+)\s+(\S+)\s+(.*)", line[8:])
                    if m:
                        if intinfo[name][item]['Description'] == '':
                            intinfo[name][item]['Description'] = m.group(1)
                        if intinfo[name][item]['Status'] == '':
                            intinfo[name][item]['Status'] = m.group(2)
                        if intinfo[name][item]['Access vlan'] == '':
                            if m.group(3) == 'trunk':
                                intinfo[name][item]['Switchport mode'] = 'trunk'
                            elif m.group(3) == 'routed':
                                intinfo[name][item]['Switchport mode'] = 'routed'
                            else:
                                intinfo[name][item]['Access vlan'] = m.group(3)
                        intinfo[name][item]['Duplex'] = m.group(4)
                        intinfo[name][item]['Speed'] = m.group(5)
                        intinfo[name][item]['Media type'] = m.group(6)


# Writes all the information collected
# Formats header style
style_header = xlwt.easyxf('pattern: pattern solid, fore_colour light_blue;'
                              'font: colour white, bold True;')

# Writes system information
cont = len(systeminfo.keys())
print(cont, " devices")

if cont > 0:

	# Writes Inventory tab
    wb = xlwt.Workbook()
    ws_system = wb.add_sheet('Inventory')

    for i, value in enumerate(systemfields):
        ws_system.write(0, i, value, style_header)

    row = 1
    for name in systeminfo.keys():

        for col in range(0,len(systemfields)):

            ws_system.write(row, col, systeminfo[name][systemfields[col]])

        row = row + 1

    # Writes interface tab
    cont = 0
    for name in intinfo.keys():
        cont = cont + len(intinfo[name])
    print(cont, " interfaces")

    if cont > 0:
        ws_int = wb.add_sheet('Interfaces')

        for i, value in enumerate(intfields):
            ws_int.write(0, i, value, style_header)

        row = 1
        for name in intinfo.keys():
            for item in intinfo[name].keys():

                for col in range(0,len(intfields)):

                    ws_int.write(row, col, intinfo[name][item][intfields[col]])

                row = row + 1

    try:
        wb.save(sys.argv[1])
    except IOError as e:
        print("Could not write " + sys.argv[1] + ". Check if file is not open in Excel. \nError: ", e)
        sys.exit(1)

else:
    print("No device found")

print("%s seconds" %(time.time() - start_time))


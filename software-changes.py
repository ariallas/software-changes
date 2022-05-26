from re import I
from pyzabbix import ZabbixAPI
import configparser
from collections import defaultdict
from openpyxl import Workbook
import datetime
import time
import sys

ZABBIX_SERVER_URL = "http://10.23.210.12/zabbix"

# Read config.ini for credentials
config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')
login = config['credentials']['login']
password = config['credentials']['password']

# Inialize API access
zapi = ZabbixAPI(ZABBIX_SERVER_URL)
zapi.login(login, password)

# time_till = datetime.datetime.now()
time_till = datetime.datetime.strptime('26.05.2022 17:20', '%d.%m.%Y %H:%M')
# time_till = datetime.datetime.strptime('24.05.2022 17:20', '%d.%m.%Y %H:%M')
time_from = time_till - datetime.timedelta(hours=15)
earliest_trigger_time = time_from + datetime.timedelta(hours=12)
timestamp_till = int(time_till.timestamp())
timestamp_from = int(time_from.timestamp())
timestamp_earliest_trigger_time = int(earliest_trigger_time.timestamp())
print("Time from: ", time_from)
print("Time till: ", time_till)
print("Earliest trigger time: ", earliest_trigger_time)

events = zapi.event.get(time_from=timestamp_earliest_trigger_time,
                        time_till=timestamp_till,
                        object=0,
                        value=1,
                        severity=1,
                        sortfield='clock',
                        sortorder='DESC',
                        output=['clock', 'objectid'],
                        filter={'name':'Произошли изменения в пакетах, установленных в системе' },
                        selectHosts=['host'])
# Compile list of hosts with software changes
changed_hosts = [ t['hosts'][0] for t in events ]

if len(events) == 0:
    print("No active triggers founds")
    sys.exit()

print(f"Triggers({len(events)}):\n", events)

# For every host get corresponding itemid for its software list
host_ids = [ h['hostid'] for h in changed_hosts ]
items = zapi.item.get(hostids=host_ids,
                      output=['hostid', 'lastclock'],
                      sortfield='itemid',
                      with_triggers=True,
                      filter={'key_': ['ubuntu.soft', 'system.sw.packages'] },
                      selectHosts=['host'])

print(f"Items({len(items)}):\n", items)

# For every itemid get 2 latest values
item_ids = [ i['itemid'] for i in items ]
history = zapi.history.get(itemids=item_ids,
                           history=4,
                           sortfield='clock',
                           sortorder='DESC',
                           time_from=timestamp_from,
                           time_till=timestamp_till,
                           output=['itemid', 'clock', 'value', 'key_'],
                           limit=len(item_ids)*2) # <- x2 here is important
# First half of the history list should be new values, second half old ones
new_packages = history[:len(item_ids)]
old_packages = history[len(item_ids):]
new_packages.sort(key=lambda h: h['itemid'])
old_packages.sort(key=lambda h: h['itemid'])

# We have all the data, now to combine it together
# This assumes that all the lists are sorted in the same order (by itemid), which they SHOULD be
hosts = []
for index in range(len(item_ids)):
    if items[index]['itemid'] != new_packages[index]['itemid'] or \
            items[index]['itemid'] != old_packages[index]['itemid']:
        print('Lists are not sorted properly')
        print(items[index]['itemid'], new_packages[index]['itemid'], old_packages[index]['itemid'])
        sys.exit()
    host = {}
    host['hostid'] = items[index]['hostid']
    host['host']   = items[index]['hosts'][0]['host']
    host['itemid'] = items[index]['itemid']
    host['clock']  = new_packages[index]['clock']
    if items[index] == 'ubuntu.soft':
        host['new_packages'] = set(new_packages[index]['value'].split('\n'))
        host['old_packages'] = set(old_packages[index]['value'].split('\n'))
    else:
        host['new_packages'] = set( new_packages[index]['value'][6:].split(', ') ) # Assuming centOS alsways have [rpm] in front
        host['old_packages'] = set( old_packages[index]['value'][6:].split(', ') )
    hosts.append(host)
hosts.sort(key=lambda h: h['clock'], reverse=True)

# Compile lists of new and removed software via set differences
for host in hosts:
    installed = list(host['new_packages'] - host['old_packages'])
    removed = list(host['old_packages'] - host['new_packages'])
    installed.sort()
    removed.sort()
    host['installed'] = installed
    host['removed'] = removed

# Compiling a dictionary of { list of changes : list of hosts with these changes }
# to group up hosts with identical changes
host_groups = {}
for host in hosts:
    key_tuple = tuple(host['installed'] + host['removed'])
    if not key_tuple in host_groups.keys():
        host_groups[key_tuple] = []
    host_groups[key_tuple].append(host)


def output_txt(filename):
    with open(filename, 'w') as f:
        for key_tuple, hosts in host_groups.items():
            for host in hosts:
                print(host['host'], file=f)
            host = hosts[0]
            if host['installed']:
                print("\nNew packages:", file=f)
                for new_package in host['installed']:
                    print(new_package, file=f)
            if host['removed']:
                print("\nRemoved packages:", file=f)
                for removed_package in host['removed']:
                    print(removed_package, file=f)
            print('------------------------------------------', file=f)
output_txt('output.txt')

def output_xlsx(filename):
    workbook = Workbook()
    sheet = workbook.active

    top_row = 1
    for key_tuple, hosts in host_groups.items():
        sheet.cell(row=top_row, column=2).value = 'Удалённые пакеты'
        sheet.cell(row=top_row, column=3).value = 'Установленные пакеты'
        row = top_row + 1
        for host in hosts:
            sheet.cell(row=row, column=1).value = host['host']
            row += 1
        row = top_row + 1
        for package in hosts[0]['removed']:
            sheet.cell(row=row, column=2).value = package
            row += 1
        row = top_row + 1
        for package in hosts[0]['installed']:
            sheet.cell(row=row, column=3).value = package
            row += 1
        top_row += max(len(hosts), len(hosts[0]['installed']), len(hosts[0]['removed'])) + 2        

    dims = {}
    for row in sheet.rows:
        for cell in row:
            if cell.value:
                dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value)))) 
    for col, value in dims.items():
        sheet.column_dimensions[col].width = value + 5
    workbook.save(filename="output.xlsx")
output_xlsx('output.xlsx')

def output_xlsx(filename):
    workbook = Workbook()
    sheet = workbook.active

    top_row = 1
    for key_tuple, hosts in host_groups.items():
        sheet.cell(row=top_row, column=2).value = 'Удалённые пакеты'
        sheet.cell(row=top_row, column=3).value = 'Установленные пакеты'
        row = top_row + 1
        for host in hosts:
            sheet.cell(row=row, column=1).value = host['host']
            row += 1
        row = top_row + 1
        for package in hosts[0]['removed']:
            sheet.cell(row=row, column=2).value = package
            row += 1
        row = top_row + 1
        for package in hosts[0]['installed']:
            sheet.cell(row=row, column=3).value = package
            row += 1
        top_row += max(len(hosts), len(hosts[0]['installed']), len(hosts[0]['removed'])) + 2        

    dims = {}
    for row in sheet.rows:
        for cell in row:
            if cell.value:
                dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value)))) 
    for col, value in dims.items():
        sheet.column_dimensions[col].width = value
    workbook.save(filename="output.xlsx")
output_xlsx('output.xlsx')

for h in history:
    del h['value']
print(f"History({len(history)}):\n", history)
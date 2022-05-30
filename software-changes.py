from pyzabbix import ZabbixAPI
from openpyxl import Workbook
from openpyxl.styles.borders import Border, Side
from argparse import ArgumentParser
import configparser
import datetime
import sys

def make_timestamp(time):     return int(time.timestamp())
def make_datetime(timestamp): return datetime.datetime.fromtimestamp(timestamp)
def format_date(date):        return date.strftime('%d.%m %H:%M')
def parse_date(str):          return datetime.datetime.strptime(str, '%d.%m.%Y %H:%M')

# Read CMD arguments
parser = ArgumentParser()
parser.add_argument('group_id',  type=str, help='Zabbix group ID')
parser.add_argument('date_from', type=str, help='Start searching from this date')
parser.add_argument('date_till', type=str, help='Search till this date')
args = parser.parse_args()

try:
    group_id = args.group_id
    date_till = parse_date(args.date_till)
    date_from = parse_date(args.date_from)
    pass
except ValueError as e:
    print(e)
    print("Date must be in DD.MM.YYYY HH:MM format")
    sys.exit()

# Read config.ini for credentials
config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')
login = config['credentials']['login']
password = config['credentials']['password']
zabbix_server_url = config['params']['zabbix_server_url']

# Inialize API access
zapi = ZabbixAPI(zabbix_server_url)
zapi.login(login, password)

hosts = zapi.host.get(groupids=group_id,
                      output=['hostid'],
                      monitored_hosts=True)
print(f"Hosts with groupid {group_id} found: {len(hosts)}")

# Abort if no hosts were found
if len(hosts) == 0:
    print("No hosts were founds")
    sys.exit()

# For every host get corresponding itemid for its software list
host_ids = [ h['hostid'] for h in hosts ]
items = zapi.item.get(hostids=host_ids,
                      output=['hostid', 'lastclock', 'key_'],
                      sortfield='itemid',
                      with_triggers=True,
                      filter={ 'key_' : ['ubuntu.soft', 'system.sw.packages'] },
                      selectHosts=['host'])

# For every itemid get its value history
print(f"Searching for history items from {format_date(date_from)} to {format_date(date_till)}")
item_ids = [ i['itemid'] for i in items ]
history = zapi.history.get(itemids=item_ids,
                           history=4,
                           sortfield='clock',
                           sortorder='DESC',
                           time_from=make_timestamp(date_from),
                           time_till=make_timestamp(date_till),
                           output=['itemid', 'clock', 'value'])
print(f"History length: {len(history)}")
if len(history) < len(item_ids):
    print("History length is less than amount of hosts")
    sys.exit()

# First batch of history results are newest software list, last batch is the oldest one. Ignoring everything inbetween
new_packages = history[:len(item_ids)]
old_packages = history[-len(item_ids):]
# Sort them by itemid, so they are sorted the same way as items list
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
    if items[index]['key_'] == 'ubuntu.soft':
        host['new_packages'] = set(new_packages[index]['value'].split('\n'))
        host['old_packages'] = set(old_packages[index]['value'].split('\n'))
    else:                      # Assuming centOS alsways have [rpm] in front
        host['new_packages'] = set( new_packages[index]['value'][6:].split(', ') )
        host['old_packages'] = set( old_packages[index]['value'][6:].split(', ') )
    hosts.append(host)
hosts.sort(key=lambda h: h['host'])

# Compile lists of new and removed software via set differences
for host in hosts:
    installed = list(host['new_packages'] - host['old_packages'])
    removed   = list(host['old_packages'] - host['new_packages'])
    host['installed'] = sorted(installed)
    host['removed']   = sorted(removed)

# Compiling a dictionary to group up hosts with identical changes
host_groups = {}
for host in hosts:
    key_tuple = tuple(host['installed'] + host['removed'])
    if not key_tuple in host_groups.keys():
        host_groups[key_tuple] = []
    host_groups[key_tuple].append(host)


# Output to a .txt file
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
output_txt('report.txt')

# Output to an .xlsx table
def output_xlsx(filename):
    def set_border(ws, cell_range, bottom_row, top_row):
        thin = Side(border_style="thin", color="000000")
        for row in ws[cell_range]:
            for cell in row:
                if cell.row == bottom_row:
                    cell.border += Border(top=thin)
                if cell.row in (top_row, bottom_row):
                    cell.border += Border(bottom=thin)
                cell.border += Border(left=thin, right=thin)

    workbook = Workbook()
    sheet = workbook.active

    top_row = 1
    for key_tuple, hosts in host_groups.items():
        sheet.cell(row=top_row, column=1).value = 'Хосты'
        sheet.cell(row=top_row, column=2).value = 'Удаленные пакеты'
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
        old_top_row = top_row
        top_row += max(len(hosts), len(hosts[0]['installed']), len(hosts[0]['removed'])) + 2
        set_border(sheet, f"A{old_top_row}:C{top_row - 2}", old_top_row, top_row - 2)

    dims = {}
    for row in sheet.rows:
        for cell in row:
            if cell.value:
                dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value)))) 
    for col, value in dims.items():
        sheet.column_dimensions[col].width = value + 5
    workbook.save(filename=filename)
output_xlsx('report.xlsx')
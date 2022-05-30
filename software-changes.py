from pyzabbix import ZabbixAPI
from openpyxl import Workbook
from openpyxl.styles.borders import Border, Side
import configparser
import datetime
import sys

def make_timestamp(time):     return int(time.timestamp())
def make_datetime(timestamp): return datetime.datetime.fromtimestamp(timestamp)
def format_date(date):        return date.strftime('%d.%m %H:%M')

# Read config.ini for credentials and search settings
config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')
login = config['credentials']['login']
password = config['credentials']['password']

zabbix_server_url = config['params']['zabbix_server_url']
search_interval   = config.getint('params', 'search_interval')
metric_interval  = config.getint('params', 'metric_interval')

# Inialize API access
zapi = ZabbixAPI(zabbix_server_url)
zapi.login(login, password)

# Date from which to start searching for trigger events (backwards in time)
time_till = datetime.datetime.now()

# These are for tests
# time_till = datetime.datetime.strptime('28.05.2022 10:10', '%d.%m.%Y %H:%M')
# time_till = datetime.datetime.strptime('27.05.2022 10:10', '%d.%m.%Y %H:%M')
# time_till = datetime.datetime.strptime('26.05.2022 17:20', '%d.%m.%Y %H:%M')
# time_till = datetime.datetime.strptime('26.05.2022 05:20', '%d.%m.%Y %H:%M')
# time_till = datetime.datetime.strptime('25.05.2022 17:20', '%d.%m.%Y %H:%M')
# time_till = datetime.datetime.strptime('24.05.2022 17:20', '%d.%m.%Y %H:%M')
# time_till = datetime.datetime.strptime('19.05.2022 17:20', '%d.%m.%Y %H:%M')
# time_till = datetime.datetime.strptime('18.05.2022 17:20', '%d.%m.%Y %H:%M')

# Search for events starting from this time
event_from = time_till - datetime.timedelta(hours=search_interval)
print(f"Searching for trigger events from {format_date(event_from)} to {format_date(time_till)}")

# Get latest software change events
events = zapi.event.get(time_from=make_timestamp(event_from),
                        time_till=make_timestamp(time_till),
                        object=0,
                        value=1,
                        suppressed=False,
                        sortfield='clock',
                        sortorder='DESC',
                        output=['clock', 'objectid'],
                        filter={'name':'Произошли изменения в пакетах, установленных в системе'},
                        selectHosts=['host'])
# List of hosts with software changes
changed_hosts = [ t['hosts'][0] for t in events ]

# Abort if no events were found
if len(events) == 0:
    print("No software update events founds")
    sys.exit()

# Latest and oldest event timestamps, will be the same if its only one batch
latest_event_time = make_datetime(int(events[0]['clock']))
oldest_event_time = make_datetime(int(events[-1]['clock']))

# For every host get corresponding itemid for its software list
host_ids = [ h['hostid'] for h in changed_hosts ]
host_ids = list(set(host_ids)) # Remove duplicate hosts
print(f"Alert count: {len(events)}    Unique hosts count: {len(host_ids)}")

items = zapi.item.get(hostids=host_ids,
                      output=['hostid', 'lastclock', 'key_'],
                      sortfield='itemid',
                      with_triggers=True,
                      filter={ 'key_' : ['ubuntu.soft', 'system.sw.packages'] },
                      selectHosts=['host'])

# For every itemid get its value history
history_from = oldest_event_time - datetime.timedelta(hours=metric_interval + 1)
print(f"Searching for history items from {format_date(history_from)} to {format_date(latest_event_time)}")
item_ids = [ i['itemid'] for i in items ]
history = zapi.history.get(itemids=item_ids,
                           history=4,
                           sortfield='clock',
                           sortorder='DESC',
                           time_from=make_timestamp(history_from),
                           time_till=make_timestamp(latest_event_time) + 1000,
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
# hosts.sort(key=lambda h: h['clock'], reverse=True) # This sorts the same way as in Zabbix
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
    if not len(key_tuple):
        print(f"No changes found for host {host['host']}")
        continue
    if not key_tuple in host_groups.keys():
        host_groups[key_tuple] = []
    host_groups[key_tuple].append(host)

print(f"Изменение ПО в ВИФИД {latest_event_time.strftime('%d.%m')} за последние 24 часа")


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
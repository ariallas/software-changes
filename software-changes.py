from re import I
from pyzabbix import ZabbixAPI
import configparser
from collections import defaultdict
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

# Get recently fired 'Software change' triggers by trigger template ID
triggers = zapi.trigger.get(only_true=1,
                            filter={'templateid' : '412805' },
                            monitored=1,
                            active=1,
                            # sortfield='lastchange',
                            # sortorder='DESC',
                            output=['hosts', 'lastchange'],
                            selectHosts=['host'])
print("Number of hosts with changes: ", len(triggers))
# Compile list of hosts with software changes
# triggers.sort(key=lambda t: t['hosts'][0]['host'])
changed_hosts = [ t['hosts'][0] for t in triggers ]

if len(triggers) == 0:
    print("No active triggers founds")
    sys.exit()

print(f"Triggers({len(triggers)}):\n", triggers)

# For every host get corresponding itemid for its software list
host_ids = [ h['hostid'] for h in changed_hosts ]
items = zapi.item.get(hostids=host_ids,
                      output=['hostid', 'lastclock'],
                      sortfield='itemid',
                      filter={"name":"Software Ubuntu"}, # !!! This will NOT work everywhere, need some other filter
                      selectHosts=['host'])

print(f"Items({len(items)}):\n", items)

# For every itemid get 2 latest values
# This request seems kinda big brain, but should work?
item_ids = [ i['itemid'] for i in items ]
history = zapi.history.get(itemids=item_ids,
                           history=4,
                           sortfield='clock',
                           sortorder='DESC',
                           output=['itemid', 'clock', 'value'],
                           limit=len(item_ids)*2) # <- x2 here is important
# First half of the history list should be new values, second half old ones
new_packages = history[:len(item_ids)]
old_packages = history[len(item_ids):]
new_packages.sort(key=lambda h: h['itemid'])
old_packages.sort(key=lambda h: h['itemid'])

# We have all the data, now to combine it together
# This assumes that all the lists are sorted in the same order (by itemid), which they SHOULD be
hosts = []
index = 0
while index < len(item_ids):
    if items[index]['itemid'] != new_packages[index]['itemid'] or \
            items[index]['itemid'] != old_packages[index]['itemid']:
        print('Lists are not sorted properly')
        print(items[index]['itemid'], new_packages[index]['itemid'], old_packages[index]['itemid'])
        sys.exit()
    host = {}
    host['hostid'] = items[index]['hostid']
    host['host']   = items[index]['hosts'][0]['host']
    host['itemid'] = items[index]['itemid']
    host['clock']  = items[index]['lastclock']
    host['new_packages'] = set(new_packages[index]['value'].split('\n')) # Will need a different split for the centOS here
    host['old_packages'] = set(old_packages[index]['value'].split('\n'))
    hosts.append(host)
    index += 1
hosts.sort(key=lambda h: h['clock'], reverse=True)

for host in hosts:
    installed = list(host['new_packages'] - host['old_packages'])
    removed = list(host['old_packages'] - host['new_packages'])
    installed.sort()
    removed.sort()
    host['installed'] = installed
    host['removed'] = removed

# Dictionary of { list of changes : list of hosts with these changes }
# To group up hosts with identical changes
host_groups = {}
for host in hosts:
    key_tuple = tuple(host['installed'] + host['removed'])
    if not key_tuple in host_groups.keys():
        host_groups[key_tuple] = []
    host_groups[key_tuple].append(host)

# Now output what we got
with open('output.txt', 'w') as f:
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

for h in history:
    del h['value']
print(f"History({len(history)}):\n", history)
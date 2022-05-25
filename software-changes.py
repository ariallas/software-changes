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
                            sortfield='lastchange',
                            sortorder='DESC',
                            output=['hosts', 'lastchange'],
                            selectHosts=['host'])
print("Number of hosts with changes: ", len(triggers))
# Compile list of hosts with software changes
hosts = [ t['hosts'][0] for t in triggers ]

# For every host get corresponding itemid for its software list
host_ids = [ h['hostid'] for h in hosts ]
items = zapi.item.get(hostids=host_ids,
                      output=['hostid', 'lastclock'],
                      filter={"name":"Software Ubuntu"}) # !!! This will NOT work everywhere, need some other filter
items = sorted(items, key=lambda d: d['lastclock'], reverse=True)

# For every itemid get 2 latest values
# This request seems kinda big brain, but should work?
item_ids = [ i['itemid'] for i in items ]
history = zapi.history.get(itemids=item_ids,
                           history=4,
                           sortfield="clock",
                           sortorder="DESC",
                           output=['itemid', 'clock', 'value'],
                           limit=len(item_ids)*2) # <- x2 here is important
# First half of the history list should be new values, second half old ones
new_packages = history[:len(item_ids)]
old_packages = history[len(item_ids):]

print(hosts)
print(items)

# We have all the data, now to combine it together
# This assumes that all the lists are sorted in the same order, which they SHOULD be
index = 0
while index < len(item_ids):
    if hosts[index]['hostid'] != items[index]['hostid'] or \
        items[index]['itemid'] != new_packages[index]['itemid'] or \
        items[index]['itemid'] != old_packages[index]['itemid']:
        print('Lists are not sorted properly')
        print(hosts[index]['hostid'], items[index]['hostid'], items[index]['itemid'], 
        new_packages[index]['itemid'], old_packages[index]['itemid'])
        sys.exit()
    hosts[index]['itemid'] = items[index]['itemid']
    hosts[index]['new_packages'] = set(new_packages[index]['value'].split('\n')) # Will need a different split for the centOS here
    hosts[index]['old_packages'] = set(old_packages[index]['value'].split('\n'))
    index += 1

# Just output what we got
with open('output.txt', 'w') as f:
    for host in hosts:
        installed = list(host['new_packages'] - host['old_packages'])
        removed = list(host['old_packages'] - host['new_packages'])

        f.write(host['host'] + '\n')
        if installed:
            f.write("New:\n")
            for new_package in installed:
                f.write(new_package)
                f.write('\n')

        if removed:
            f.write("Removed:\n")
            for removed_package in removed:
                f.write(removed_package)
                f.write('\n')
        f.write('\n')

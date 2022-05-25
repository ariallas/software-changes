from pyzabbix import ZabbixAPI
import configparser
import datetime
import time
import sys

ZABBIX_SERVER_URL = "http://10.23.210.12/zabbix"

# Read config.ini for credentials
config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')
login = config['credentials']['login']
password = config['credentials']['login']

# Inialize API access
zapi = ZabbixAPI(ZABBIX_SERVER_URL)
zapi.login(login, password)


# Get recently fired software change triggers
# triggers = zapi.trigger.get(only_true=1,
#                             filter={'templateid':'412805'},
#                             skipDependent=1,
#                             monitored=1,
#                             active=1,
#                             output='extend',
#                             expandDescription=1,
#                             selectHosts=['host'])
# for t in triggers:
#     print(t)


# items = zapi.item.get(hostids="11695",
#                       filter={"name":"Software Ubuntu"})
# print(items[0])

history = zapi.history.get(itemids="655944",
                           history=4,
                           sortfield="clock",
                           sortorder="DESC",
                           limit=2)
new_packages = set(history[0]["value"].split('\n'))
old_packages = set(history[1]["value"].split('\n'))
print("New packages: ", list(new_packages - old_packages))
print("Removed packages: ", list(old_packages - new_packages))
# for h in history:
#     print(h["clock"])


# triggers = zapi.trigger.get(host="prod-vdnhpromet-01.vifid.ru")
# print(triggers)
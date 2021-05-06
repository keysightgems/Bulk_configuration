# Bulk Config Assistant POC
This repository is regarding the config assistant which will do bulk config.

# Key Features
•	Config Assistant for bulk config

•	Create default excel file with min arguments.

•	Save config as Json

# Required python packages
pip install pandas

pip install xlsxwriter

# Import the packages
from create_excel import API_Sheet

from bulk_config import BulkConfig

# Run script

To generate excel sheet with default parameters

excel=API_Sheet()

excelName=excel.create_excel()

To generate excel sheet with custom parameters

Implemented custom parameters: 

	build_information, base, physical, devicegroup, ipv4_ethernet, ipv6_ethernet, ipv4_bgp, ipv6_bgp, ipv4_loopback, bgp_capabilities, ipv4_ospf, ipv6_ospf, isis, networkgroup, igmp_host, igmp_querier, traffic, packet_editor.

excel=API_Sheet()

excelName=excel.create_excel(parameter=dict(list(value)))
# Example:
	devicegroup = {'Topology':['SI-FANOUT-SW11'], 'Device Group':['SI-FANOUT-SW11_Vlan2000'], 'Multiplier':[1], 'Vlan Header':[2000]}
	
	For Example: If 'Vlan Header' is multivalue please send the value like below.
	
		For increment: 'Vlan Header':['increment;2001;1']
	
		For singleValue: 'Vlan Header':[2000]
		
		For valueList: 'Vlan Header':['valuelist;2001;2010;2011;2012;2020;2021;2022']

	ipv4_ethernet = {'Device Group':['SI-FANOUT-SW11_Vlan2000']}

# Create bulk config
tgnObj = BulkConfig('127.0.0.1', clearConfig=True)

tgnObj.bulk_config(excelName)

tgnObj.create_trafficitems(excelName)


from obtain_data_from_excel import excelReader
from datetime import datetime
from ixnetwork_restpy import SessionAssistant
import json,re

class BulkConfig():
    def __init__(self,apiServerIp, clearConfig):
        session_assistant = SessionAssistant(IpAddress=apiServerIp,
            UserName='admin', Password='admin',
            LogLevel=SessionAssistant.LOGLEVEL_INFO,
            ClearConfig=clearConfig)
        self.ixnetwork = session_assistant.Ixnetwork
        self.config = []
        self.vportList = []
        self.portList = []

    def add_port_to_list(self, chassis_ip, linecard, port, topology_name, port_name):
        info = {'chassis': chassis_ip, 'line_card': linecard, 'port': port,
                'top_name': topology_name, 'port_name': port_name}
        self.portList.append(info)
        return self.portList

    def generate_vport_list(self, portList):
        for vportidx, vport_dict in enumerate(portList, 1):
            info = {'vport': f'vport{vportidx}', 'name': vport_dict['port_name'],
                    'top_name': vport_dict['top_name'], 'top_idx': vportidx}
            self.vportList.append(info)
        return self.vportList

    def config_multivalueObj(self, deviceGroupValue, stackPath, stackAttribute):

        if not isinstance(deviceGroupValue, int):
            stackvaluelist = deviceGroupValue.split(';')
            if stackvaluelist[0] == 'valuelist':
                stackvaluelist.pop(0)
                stackvalueinfo = {
                    "xpath": '"/multivalue[@source = ' + stackPath + " " +stackAttribute+"']/valueList",
                    "values": stackvaluelist
                    }
                self.config.append({key: value for key, value in stackvalueinfo.items()})
            elif stackvaluelist[0] == 'increment':
                stackvalueinfo = {"xpath": '"/multivalue[@source = ' + stackPath + " " +stackAttribute+"']/counter",
                                      "start": stackvaluelist[1], "step": stackvaluelist[2], "direction": 'increment'}
                self.config.append({key: value for key, value in stackvalueinfo.items()})
        if isinstance(deviceGroupValue, int):
            stackvalueinfo = {"xpath": '"/multivalue[@source = ' + stackPath + " " + stackAttribute + "']/singleValue",
                              "value": deviceGroupValue}
            self.config.append({key: value for key, value in stackvalueinfo.items()})
        if isinstance(deviceGroupValue, str) and ";" not in deviceGroupValue:
            stackvalueinfo = {"xpath": '"/multivalue[@source = ' + stackPath + " " +stackAttribute+"']/singleValue", "value": deviceGroupValue}
            self.config.append({key: value for key, value in stackvalueinfo.items()})

    def bulk_config(self, workbook_name):
        # Initiates the excelReader class to obtain all data from the input file
        excel = excelReader(workbook_name)
        Worksheet_Dict = excel.get_worksheet_data()
        status_dict = excel.obtain_status()
        if 'Physical' in status_dict:
            if 'Physical' in Worksheet_Dict:
                for port in Worksheet_Dict['Physical']:
                    self.add_port_to_list(port['Chassis IP'], port['Linecard Number'],
                                                port['Port Number'], port['Topology Name'], port['Port Name'])
                if self.portList != []:
                    vportList = self.generate_vport_list(self.portList)
                for idx, vport in enumerate(vportList):
                    devicegroupIndex = 1
                    # If set to True it will take the information within the pysical worksheet
                    if status_dict['Physical'] == True:
                        vportidx = idx + 1
                        # Run the create_topology function which creates the physical port to vport mapping
                        vportinfo = {'xpath': ('/vport[' + str(vportidx) + ']'),
                                     'name': vport['name']}
                        self.config.append({key: value for key, value in vportinfo.items()})
                        topoinfo = {'xpath': ('/topology[' + str(vportidx) + ']'),
                                    'name': vport['top_name'],
                                    'ports': [('/vport[' + str(vportidx) + ']')]}
                        self.config.append({key: value for key, value in topoinfo.items()})
                        # If set to True it will take the information within the device groupe worksheet
                        if status_dict['Devicegroup'] == True:
                            # Loops through each device group defined within the worksheet
                            for row in Worksheet_Dict['Devicegroup']:
                                # Verifiy that the topology matches one mapped to a port created within the physical sheets
                                if row['Topology'] == self.portList[idx]['top_name']:
                                    # Creates a dictionary containing the device group information
                                    if topoinfo['name'] == row['Topology']:
                                        value = (topoinfo['xpath'] + '/deviceGroup[' + str(devicegroupIndex) + ']')
                                        if 'Multiplier' in row:
                                            multiplier = row['Multiplier']
                                        else:
                                            multiplier = 1
                                        devicegroupinfo = {'xpath': value,
                                                           'multiplier': multiplier,
                                                           'name': row['Device Group']}
                                        self.config.append({key: value for key, value in devicegroupinfo.items()})
                                        ethernetinfo = {'xpath': (devicegroupinfo['xpath'] + '/ethernet[1]'),
                                                        'name': ('Ethernet ' + str(devicegroupIndex))}
                                        self.config.append({key: value for key, value in ethernetinfo.items()})
                                        # Adding vlan info to the Ethernet
                                        if 'Device Group' in row:
                                            if row['Device Group'] == devicegroupinfo['name']:
                                                if 'Vlan Header' in row:
                                                    if row['Vlan Header']:
                                                        enablevlaninfo = {"xpath": '"/multivalue[@source = ' + ethernetinfo[
                                                            'xpath'] + " enableVlans']/singleValue",
                                                                          "value": "true"}
                                                        self.config.append({key: value for key, value in enablevlaninfo.items()})
                                                        vlanxpathinfo = {"xpath": ethernetinfo['xpath'] + "/vlan[1]"}
                                                        self.config.append({key: value for key, value in vlanxpathinfo.items()})
                                                        # Vlan with random user defined values
                                                        if not isinstance(row['Vlan Header'], int):
                                                            vlanlist = row['Vlan Header'].split(';')
                                                            if vlanlist[0] == 'valuelist':
                                                                vlanlist.pop(0)
                                                                vlanIDinfo = {"xpath": '"/multivalue[@source = ' + vlanxpathinfo['xpath'] + " vlanId']/valueList",
                                                                              "values": vlanlist}
                                                                self.config.append({key: value for key, value in vlanIDinfo.items()})
                                                            if vlanlist[0] == 'increment':
                                                                vlanIDinfo = {"xpath": '"/multivalue[@source = ' + vlanxpathinfo['xpath'] + " vlanId']/counter",
                                                                              "start": vlanlist[1], "step": vlanlist[2], "direction": 'increment'}
                                                                self.config.append({key: value for key, value in vlanIDinfo.items()})
                                                        # Single vlan ID json creation
                                                        if isinstance(row['Vlan Header'], int):
                                                            vlanIDinfo = {"xpath": '"/multivalue[@source = ' + vlanxpathinfo[
                                                                'xpath'] + " vlanId']/singleValue",
                                                                          "value": row['Vlan Header']}
                                                            self.config.append({key: value for key, value in vlanIDinfo.items()})
                                                    if 'VLAN Count' in row:
                                                        if row['VLAN Count']:
                                                            vlanCountinfo = {"xpath": ethernetinfo['xpath'],
                                                                          "vlanCount": row['VLAN Count']}
                                                            self.config.append({key: value for key, value in vlanCountinfo.items()})

                                        # Verify IPv4 data presence and add stack
                                        if 'IPv4_Ethernet' in status_dict:
                                            if status_dict['IPv4_Ethernet'] == True:
                                                if 'IPv4_Ethernet' in Worksheet_Dict:
                                                    for device_group_name in Worksheet_Dict['IPv4_Ethernet']:
                                                        # devicegroupinfo['name'] = devicegroupinfo['name'].replace(" ", "")
                                                        if device_group_name['Device Group'] == devicegroupinfo['name']:
                                                            ipv4info = {'xpath': (ethernetinfo['xpath'] + '/ipv4[1]'),
                                                                        'name': ('ipv4_ ' + str(devicegroupIndex))}
                                                            self.config.append({key: value for key, value in ipv4info.items()})
                                                            if 'Address' in device_group_name:
                                                                self.config_multivalueObj(device_group_name['Address'], ipv4info['xpath'], 'address')
                                                            if 'Prefix' in device_group_name:
                                                                ipv4prefix = {"xpath": '"/multivalue[@source = ' + ipv4info['xpath'] + " prefix']/singleValue",
                                                                              "value": device_group_name['Prefix']
                                                                              }
                                                                self.config.append({key: value for key, value in ipv4prefix.items()})
                                                            if 'Gateway IP' in device_group_name:
                                                                self.config_multivalueObj(device_group_name['Gateway IP'], ipv4info['xpath'], 'gatewayIp')
                                                            if 'Resolve Gateway' in device_group_name:
                                                                if device_group_name['Resolve Gateway'] == 'yes':
                                                                    resolvegwvalue = True
                                                                else:
                                                                    resolvegwvalue = False
                                                                resolvegwinfo = {
                                                                    "xpath": '"/multivalue[@source = ' + ipv4info['xpath'] + " resolveGateway']/singleValue",
                                                                    "value": resolvegwvalue
                                                                }
                                                                self.config.append({key: value for key, value in resolvegwinfo.items()})

                                        # Verify IPv6 data presence and add stack
                                        if 'IPv6_Ethernet' in status_dict:
                                            if status_dict['IPv6_Ethernet'] == True:
                                                ipv6 = dict()
                                                if 'IPv6_Ethernet' in Worksheet_Dict:
                                                    for device_group_name in Worksheet_Dict['IPv6_Ethernet']:
                                                        # devicegroupinfo['name'] = devicegroupinfo['name'].replace(" ", "")
                                                        if device_group_name['Device Group'] == devicegroupinfo['name']:
                                                            ipv6.update({"xpath": (ethernetinfo['xpath'] + '/ipv6[1]')})
                                                            self.config.append({key: value for key, value in ipv6.items()})
                                                            if 'Address' in device_group_name:
                                                                self.config_multivalueObj(device_group_name['Address'], ipv6['xpath'], 'address')
                                                            if 'Prefix' in device_group_name:
                                                                ipv6prefix = {"xpath": '"/multivalue[@source = ' + ipv6['xpath'] + " prefix']/singleValue",
                                                                              "value": device_group_name['Prefix']}
                                                                self.config.append({key: value for key, value in ipv6prefix.items()})
                                                            if 'Gateway IP' in device_group_name:
                                                                self.config_multivalueObj(device_group_name['Gateway IP'], ipv6['xpath'], 'gatewayIp')
                                                            if 'Resolve Gateway' in device_group_name:
                                                                if device_group_name['Resolve Gateway'] == 'yes':
                                                                    resolvegwvalue = True
                                                                else:
                                                                    resolvegwvalue = False
                                                                resolvegwinfo = {
                                                                    "xpath": '"/multivalue[@source = ' + ipv6['xpath'] + " resolveGateway']/singleValue",
                                                                    "value": resolvegwvalue
                                                                }
                                                                self.config.append({key: value for key, value in resolvegwinfo.items()})
                                        bgpCapAttributes = {'IPV4 Unicast': 'capabilityIpV4Unicast',
                                                            'IPv4 Multicast': 'capabilityIpV4Multicast',
                                                            'IPv4 MPLS VPN': 'capabilityIpV4MplsVpn', \
                                                            'VPLS': 'capabilityVpls',
                                                            'Route Refresh': 'capabilityRouteRefresh',
                                                            'Route Constraint': 'capabilityRouteConstraint', \
                                                            'IPV6 Unicast': 'capabilityIpV6Unicast',
                                                            'IPv6 Multicast': 'capabilityIpV6Multicast',
                                                            'IPv6 MPLS VPN': 'capabilityIpV6MplsVpn'}
                                        if 'IPv4_Loopback_BGP' in status_dict:
                                            if status_dict['IPv4_Loopback_BGP'] == True:
                                                loopback = dict()
                                                loopbackAttributes = {'Loopback Adress':'address', 'Prefix':'prefix', 'Peer IP':'dutIp', 'Type':'type', 'AS':'localAs2Bytes',\
                                                                      'Hold Timer':'holdTimer', 'Keepalive':'keepaliveTimer', 'Authentication':'authentication', 'Key':'md5Key'}
                                                if 'IPv4_Loopback_BGP' in Worksheet_Dict:
                                                    for device_group_name in Worksheet_Dict['IPv4_Loopback_BGP']:
                                                        # devicegroupinfo['name'] = devicegroupinfo['name'].replace(" ", "")
                                                        if 'Device Group' in device_group_name:
                                                            if device_group_name['Device Group'] == devicegroupinfo['name']:
                                                                loopback.update({"xpath": (devicegroupinfo['xpath'] + '/ipv4Loopback[1]')})
                                                                self.config.append({key: value for key, value in loopback.items()})
                                                                device_group_name.pop('Device Group')
                                                                for loopbackKey in device_group_name:
                                                                    if loopbackKey == 'Loopback Adress' or loopbackKey == 'Prefix':
                                                                        self.config_multivalueObj(device_group_name[loopbackKey], loopback['xpath'], loopbackAttributes[loopbackKey])
                                                                    else:
                                                                        self.config_multivalueObj(device_group_name[loopbackKey],
                                                                                                  loopback['xpath']+'/bgpIpv4Peer[1]',
                                                                                                  loopbackAttributes[loopbackKey])

                                            if 'BGP_Capabilities' in status_dict:
                                                if status_dict['BGP_Capabilities'] == True:
                                                    if 'BGP_Capabilities' in Worksheet_Dict:
                                                        for device_group_name in Worksheet_Dict['BGP_Capabilities']:
                                                            # devicegroupinfo['name'] = devicegroupinfo['name'].replace(" ", "")
                                                            if 'Device Group' in device_group_name:
                                                                if device_group_name['Device Group'] == devicegroupinfo['name']:
                                                                    loopback.update({"xpath": (devicegroupinfo['xpath'] + '/ipv4Loopback[1]/bgpIpv4Peer[1]')})
                                                                    self.config.append({key: value for key, value in loopback.items()})
                                                                    device_group_name.pop('Device Group')
                                                                    for bgpKey in device_group_name:
                                                                        self.config_multivalueObj(device_group_name[bgpKey], loopback['xpath'], bgpCapAttributes[bgpKey])

                                        # Verify BGP data presence and add stack
                                        if 'IPv4_BGP' in status_dict:
                                            if status_dict['IPv4_BGP'] == True:
                                                bgpv4 = dict()
                                                bgpv4Attributes = {'Dut IP':'dutIp', 'BGP Id':'bgpId', 'Type':'type', 'Local AS':'localAs2Bytes', 'Enable As 4bytes':'enable4ByteAs', 'Local AS 4byte':'localAs4Bytes', \
                                                                   'As Mode':'asSetMode', 'Enable BFD':'enableBfdRegistration', 'BFD Mode':'modeOfBfdOperations', 'Hold Timer':'holdTimer', 'Config Keepalive':'configureKeepaliveTimer', 'Keepalive':'keepaliveTimer', \
                                                                   'Update Interval':'updateInterval', 'TTL':'ttl', 'Authentication':'authentication', 'Key':'md5Key', 'Flap':'flap',
                                                                   'Uptime in Seconds':'uptimeInSec', 'Downtime in Seconds':'downtimeInSec'}
                                                if 'IPv4_BGP' in Worksheet_Dict:
                                                    for device_group_name in Worksheet_Dict['IPv4_BGP']:
                                                        # devicegroupinfo['name'] = devicegroupinfo['name'].replace(" ", "")
                                                        if 'Device Group' in device_group_name:
                                                            if device_group_name['Device Group'] == devicegroupinfo['name']:
                                                                bgpv4.update({"xpath": (ethernetinfo['xpath'] + '/ipv4[1]' + '/bgpIpv4Peer[1]')})
                                                                self.config.append({key: value for key, value in bgpv4.items()})
                                                                device_group_name.pop('Device Group')
                                                                for bgpKey in device_group_name:
                                                                    self.config_multivalueObj(device_group_name[bgpKey], bgpv4['xpath'], bgpv4Attributes[bgpKey])
                                                if 'BGP_Capabilities' in status_dict:
                                                    if status_dict['BGP_Capabilities'] == True:
                                                        if 'BGP_Capabilities' in Worksheet_Dict:
                                                            for device_group_name in Worksheet_Dict['BGP_Capabilities']:
                                                                # devicegroupinfo['name'] = devicegroupinfo['name'].replace(" ", "")
                                                                if 'Device Group' in device_group_name:
                                                                    if device_group_name['Device Group'] == devicegroupinfo['name']:
                                                                        bgpv4.update({"xpath": (ethernetinfo['xpath'] + '/ipv4[1]' + '/bgpIpv4Peer[1]')})
                                                                        self.config.append({key: value for key, value in bgpv4.items()})
                                                                        device_group_name.pop('Device Group')
                                                                        for bgpKey in device_group_name:
                                                                            self.config_multivalueObj(device_group_name[bgpKey], bgpv4['xpath'], bgpCapAttributes[bgpKey])

                                        if 'IPv6_BGP' in status_dict:
                                            if status_dict['IPv6_BGP'] == True:
                                                bgpv6 = dict()
                                                bgpv6Attributes = {'Dut IP':'dutIp', 'Type':'type', 'Local AS':'localAs2Bytes', 'Enable As 4bytes':'enable4ByteAs', 'Local AS 4byte':'localAs4Bytes',\
                                                                   'Hold Timer':'holdTimer', 'Config Keepalive':'configureKeepaliveTimer', 'Keepalive':'keepaliveTimer', 'Authentication':'authentication',\
                                                                   'Key':'md5Key', 'As Mode':'asSetMode', 'Enable BFD':'enableBfdRegistration', 'BFD Mode':'modeOfBfdOperations', 'Flap':'flap',
                                                                   'Uptime in Seconds':'uptimeInSec', 'Downtime in Seconds':'downtimeInSec'}
                                                if 'IPv6_BGP' in Worksheet_Dict:
                                                    for device_group_name in Worksheet_Dict['IPv6_BGP']:
                                                        # devicegroupinfo['name'] = devicegroupinfo['name'].replace(" ", "")
                                                        if 'Device Group' in device_group_name:
                                                            if device_group_name['Device Group'] == devicegroupinfo['name']:
                                                                bgpv6.update({"xpath": (ethernetinfo['xpath'] + '/ipv6[1]' + '/bgpIpv6Peer[1]')})
                                                                self.config.append({key: value for key, value in bgpv6.items()})
                                                                device_group_name.pop('Device Group')
                                                                for bgpKey in device_group_name:
                                                                    self.config_multivalueObj(device_group_name[bgpKey], bgpv6['xpath'],
                                                                                              bgpv6Attributes[bgpKey])
                                                if 'BGP_Capabilities' in status_dict:
                                                    if status_dict['BGP_Capabilities'] == True:
                                                        if 'BGP_Capabilities' in Worksheet_Dict:
                                                            for device_group_name in Worksheet_Dict['BGP_Capabilities']:
                                                                # devicegroupinfo['name'] = devicegroupinfo['name'].replace(" ", "")
                                                                if 'Device Group' in device_group_name:
                                                                    if device_group_name['Device Group'] == devicegroupinfo['name']:
                                                                        bgpv6.update({"xpath": (ethernetinfo['xpath'] + '/ipv4[1]' + '/bgpIpv4Peer[1]')})
                                                                        self.config.append({key: value for key, value in bgpv6.items()})
                                                                        device_group_name.pop('Device Group')
                                                                        for bgpKey in device_group_name:
                                                                            self.config_multivalueObj(device_group_name[bgpKey], bgpv6['xpath'], bgpCapAttributes[bgpKey])

                                        if 'IPv4_OSPF' in status_dict:
                                            if status_dict['IPv4_OSPF'] == True:
                                                ospfv4 = dict()
                                                ospfv4Attributes = {'Neighbor IP':'neighborIp', 'Area':'areaId', 'Network Type':'networkType', 'Hello Timers':'helloInterval',\
                                                                    'Dead Timers':'deadInterval', 'Routing Metric':'metric', 'Validate Receive MTU':'validateRxMtu', 'MTU':'maxMtu',\
                                                                    'Authentication':'authentication', 'Key Id':'md5KeyId', 'Key':'md5Key'}
                                                if 'IPv4_OSPF' in Worksheet_Dict:
                                                    for device_group_name in Worksheet_Dict['IPv4_OSPF']:
                                                        # devicegroupinfo['name'] = devicegroupinfo['name'].replace(" ", "")
                                                        if 'Device Group' in device_group_name:
                                                            if device_group_name['Device Group'] == devicegroupinfo['name']:
                                                                ospfv4.update({"xpath": (ethernetinfo['xpath'] + '/ipv4[1]' + '/ospfv2[1]')})
                                                                self.config.append({key: value for key, value in ospfv4.items()})
                                                                device_group_name.pop('Device Group')
                                                                for ospfKey in device_group_name:
                                                                    self.config_multivalueObj(device_group_name[ospfKey], ospfv4['xpath'], ospfv4Attributes[ospfKey])

                                        if 'IPv6_OSPF' in status_dict:
                                            if status_dict['IPv6_OSPF'] == True:
                                                ospfv6 = dict()
                                                ospfv6Attributes = {'Neighbor IP': 'neighborIp', 'Area': 'areaId',
                                                                    'Network Type': 'networkType', 'Hello Timers': 'helloInterval', \
                                                                    'Dead Timers': 'deadInterval', 'Link Metric': 'linkMetric', 'Authentication Algo': 'authAlgo',\
                                                                    'Authentication': 'authentication', 'SA Id': 'saId', 'Key': 'md5Key'}
                                                if 'IPv6_OSPF' in Worksheet_Dict:
                                                    for device_group_name in Worksheet_Dict['IPv6_OSPF']:
                                                        # devicegroupinfo['name'] = devicegroupinfo['name'].replace(" ", "")
                                                        if 'Device Group' in device_group_name:
                                                            if device_group_name['Device Group'] == devicegroupinfo['name']:
                                                                ospfv6.update({"xpath": (ethernetinfo['xpath'] + '/ipv6[1]' + '/ospfv3[1]')})
                                                                self.config.append({key: value for key, value in ospfv6.items()})
                                                                device_group_name.pop('Device Group')
                                                                for ospfKey in device_group_name:
                                                                    self.config_multivalueObj(device_group_name[ospfKey],
                                                                                              ospfv6['xpath'], ospfv6Attributes[ospfKey])
                                        if 'ISIS' in status_dict:
                                            if status_dict['ISIS'] == True:
                                                isis = dict()
                                                isisAttributes = {'Interface Metric': 'interfaceMetric', 'Weight':'weight', \
                                                                    'Enable Hold Time': 'enableConfiguredHoldTime', 'Configured Hold Time': 'configuredHoldTime', 'Enable 3WayHandshake':'enable3WayHandshake', \
                                                                    'Enable MT': 'enableMT', 'Enable Adj Sid': 'enableAdjSID', 'Adj Sid': 'adjSID', 'Enable BFD':'enableBfdRegistration',\
                                                                    'Ipv6 Metric': 'ipv6MTMetric', 'Network Type': 'networkType', 'Level Type': 'levelType', 'Level 1 Hello Interval':'level1HelloInterval', 'Level 1 Dead Interval':'level1DeadInterval',\
                                                                    'Max Sl Msd':'maxSlMsd', 'Level 2 Hello Interval':'level2HelloInterval', 'Level 2 Dead Interval':'level2DeadInterval', 'Authentication Type':'authType', 'Key':'circuitTranmitPasswordOrMD5Key'}
                                                if 'ISIS' in Worksheet_Dict:
                                                    for device_group_name in Worksheet_Dict['ISIS']:
                                                        # devicegroupinfo['name'] = devicegroupinfo['name'].replace(" ", "")
                                                        if 'Device Group' in device_group_name:
                                                            if device_group_name['Device Group'] == devicegroupinfo['name']:
                                                                isis.update({"xpath": (ethernetinfo['xpath'] + '/isisL3[1]')})
                                                                self.config.append({key: value for key, value in isis.items()})
                                                                device_group_name.pop('Device Group')
                                                                for isisKey in device_group_name:
                                                                    self.config_multivalueObj(device_group_name[isisKey],
                                                                                              isis['xpath'], isisAttributes[isisKey])

                                        if 'IGMP_Host' in status_dict:
                                            if status_dict['IGMP_Host'] == True:
                                                igmpHost = dict()
                                                igmpHostGroupAttributes = {'Start Group Address':'startMcastAddr', 'Group Address Incr':'mcastAddrIncr', 'Group Address Count':'mcastAddrCnt', 'Source Mode':'sourceMode'}
                                                igmpHostSourceAttributes = {'Start Source Address': 'startUcastAddr', 'Source Address Incr': 'ucastAddrIncr',
                                                                           'Source Address Count': 'ucastSrcAddrCnt'}
                                                if 'IGMP_Host' in Worksheet_Dict:
                                                    for device_group_name in Worksheet_Dict['IGMP_Host']:
                                                        # devicegroupinfo['name'] = devicegroupinfo['name'].replace(" ", "")
                                                        if 'Device Group' in device_group_name:
                                                            if device_group_name['Device Group'] == devicegroupinfo['name']:
                                                                igmpHost.update({"xpath": (ethernetinfo['xpath'] + '/ipv4[1]' + '/igmpHost[1]')})
                                                                self.config.append({key: value for key, value in igmpHost.items()})
                                                                device_group_name.pop('Device Group')
                                                                if 'No Of Group Ranges' in device_group_name:
                                                                    groupRanges = {"xpath": (ethernetinfo['xpath'] + '/ipv4[1]' + '/igmpHost[1]'),
                                                                                  "noOfGrpRanges": device_group_name['No Of Group Ranges']}
                                                                    self.config.append({key: value for key, value in groupRanges.items()})
                                                                if 'Join/Leave Multiplier' in device_group_name:
                                                                    groupRanges = {"xpath": (ethernetinfo['xpath'] + '/ipv4[1]' + '/igmpHost[1]'),
                                                                                  "jlMultiplier": device_group_name['Join/Leave Multiplier']}
                                                                    self.config.append({key: value for key, value in groupRanges.items()})
                                                                if 'Version' in device_group_name:
                                                                    self.config_multivalueObj(device_group_name['Version'], igmpHost['xpath'], 'versionType')
                                                                if 'No Of Source Ranges' in device_group_name:
                                                                    sourceRanges = {"xpath": (ethernetinfo['xpath'] + '/ipv4[1]' + '/igmpHost[1]/igmpMcastIPv4GroupList'),
                                                                                  "noOfSrcRanges": device_group_name['No Of Source Ranges']}
                                                                    self.config.append({key: value for key, value in sourceRanges.items()})
                                                                for igmpHostKey in device_group_name:
                                                                    try:
                                                                        self.config_multivalueObj(device_group_name[igmpHostKey],
                                                                                                  igmpHost['xpath']+'/igmpMcastIPv4GroupList', igmpHostGroupAttributes[igmpHostKey])
                                                                    except:
                                                                        pass

                                                                for igmpHostKey in device_group_name:
                                                                    try:
                                                                        self.config_multivalueObj(device_group_name[igmpHostKey],
                                                                                                  igmpHost['xpath']+'/igmpMcastIPv4GroupList/igmpUcastIPv4SourceList', igmpHostSourceAttributes[igmpHostKey])
                                                                    except:
                                                                        pass

                                        if 'IGMP_Querier' in status_dict:
                                            if status_dict['IGMP_Querier'] == True:
                                                igmpQuerier = dict()
                                                igmpQuerierAttributes = {'Version':'versionType', 'Query Count':'startupQueryCount', 'Query Interval':'generalQueryInterval', 'Router Alert':'routerAlert', 'Robustness':'robustnessVariable',\
                                                                         'Query Response Interval':'generalQueryResponseInterval','Transmission Count':'specificQueryTransmissionCount'}
                                                if 'IGMP_Querier' in Worksheet_Dict:
                                                    for device_group_name in Worksheet_Dict['IGMP_Querier']:
                                                        # devicegroupinfo['name'] = devicegroupinfo['name'].replace(" ", "")
                                                        if 'Device Group' in device_group_name:
                                                            if device_group_name['Device Group'] == devicegroupinfo['name']:
                                                                igmpQuerier.update({"xpath": (ethernetinfo['xpath'] + '/ipv4[1]' + '/igmpQuerier[1]')})
                                                                self.config.append({key: value for key, value in igmpQuerier.items()})
                                                                device_group_name.pop('Device Group')
                                                                for igmpQuerierKey in device_group_name:
                                                                    self.config_multivalueObj(device_group_name[igmpQuerierKey],
                                                                                              igmpQuerier['xpath'], igmpQuerierAttributes[igmpQuerierKey])

                                        if 'Network_Group' in status_dict:
                                            if status_dict['Network_Group'] == True:
                                                networkGroup = dict()
                                                if 'Network_Group' in Worksheet_Dict:
                                                    for device_group_name in Worksheet_Dict['Network_Group']:
                                                        # devicegroupinfo['name'] = devicegroupinfo['name'].replace(" ", "")
                                                        if 'Device Group' in device_group_name:
                                                            if device_group_name['Device Group'] == devicegroupinfo['name']:
                                                                networkGroup.update({"xpath": (devicegroupinfo['xpath'] + '/networkGroup[1]')})
                                                                self.config.append({key: value for key, value in networkGroup.items()})
                                                                if 'Name' in device_group_name:
                                                                    networkGroupName = {"xpath": (devicegroupinfo['xpath'] + '/networkGroup[1]'),
                                                                        "name": device_group_name['Name']}
                                                                    self.config.append({key: value for key, value in networkGroupName.items()})
                                                                if 'Multiplier' in device_group_name:
                                                                    networkGroupMul = {"xpath": (devicegroupinfo['xpath'] + '/networkGroup[1]'),
                                                                        "multiplier": device_group_name['Multiplier']}
                                                                    self.config.append({key: value for key, value in networkGroupMul.items()})
                                                                if 'IP Version' in device_group_name:
                                                                    if ';' in device_group_name['IP Version']:
                                                                        ipVersionlist = device_group_name['IP Version'].split(';')
                                                                    else:
                                                                        ipVersionlist = [device_group_name['IP Version']]
                                                                    for ipVersion in ipVersionlist:
                                                                        if ipVersion == "ipv4":
                                                                            if 'Ipv4 Address' in device_group_name and device_group_name['Ipv4 Address'] != None:
                                                                                self.config_multivalueObj(device_group_name['Ipv4 Address'],
                                                                                                          devicegroupinfo['xpath'] + '/networkGroup[1]'+'/ipv4PrefixPools[1]', 'networkAddress')
                                                                            if 'Ipv4 Prefix' in device_group_name and device_group_name['Ipv4 Prefix'] != None:
                                                                                self.config_multivalueObj(device_group_name['Ipv4 Prefix'],
                                                                                                          devicegroupinfo['xpath'] + '/networkGroup[1]'+'/ipv6PrefixPools[1]', 'prefixLength')
                                                                            if 'Step' in device_group_name:
                                                                                self.config_multivalueObj(device_group_name['Step'],
                                                                                    devicegroupinfo['xpath'] + '/networkGroup[1]'+'/ipv4PrefixPools[1]', 'prefixAddrStep')
                                                                            if 'Address Count' in device_group_name:
                                                                                self.config_multivalueObj(device_group_name['Address Count'],
                                                                                    devicegroupinfo['xpath'] + '/networkGroup[1]'+'/ipv4PrefixPools[1]', 'numberOfAddressesAsy')
                                                                            if 'Protocol' in device_group_name:
                                                                                if device_group_name['Protocol'] == "ospfv2":
                                                                                    protocol = {"xpath": (devicegroupinfo['xpath'] + '/networkGroup[1]'+'/ipv4PrefixPools[1]/connector'),
                                                                                                  "connectedTo": (ethernetinfo['xpath'] + '/ipv4[1]' + '/ospfv2[1]')}
                                                                                    self.config.append({key: value for key, value in protocol.items()})
                                                                                if device_group_name['Protocol'] == "bgpv4":
                                                                                    protocol = {"xpath": (devicegroupinfo['xpath'] + '/networkGroup[1]'+'/ipv4PrefixPools[1]/connector'),
                                                                                                  "connectedTo": (ethernetinfo['xpath'] + '/ipv4[1]' + '/bgpIpv4Peer[1]')}
                                                                                    self.config.append({key: value for key, value in protocol.items()})
                                                                                if device_group_name['Protocol'] == "isis":
                                                                                    protocol = {"xpath": (devicegroupinfo['xpath'] + '/networkGroup[1]'+'/ipv4PrefixPools[1]/connector'),
                                                                                                  "connectedTo": (ethernetinfo['xpath'] + '/ipv4[1]' + '/isisL3[1]')}
                                                                                    self.config.append({key: value for key, value in protocol.items()})

                                                                        if ipVersion == "ipv6" and len(ipVersionlist) != 2:
                                                                            if 'Ipv6 Address' in device_group_name and device_group_name['Ipv6 Address'] != None:
                                                                                self.config_multivalueObj(device_group_name['Ipv6 Address'],
                                                                                                          devicegroupinfo['xpath'] + '/networkGroup[1]'+'/ipv6PrefixPools[1]', 'networkAddress')
                                                                            if 'Ipv6 Prefix' in device_group_name and device_group_name['Ipv6 Prefix'] != None:
                                                                                self.config_multivalueObj(device_group_name['Ipv6 Prefix'],
                                                                                                          devicegroupinfo['xpath'] + '/networkGroup[1]'+'/ipv6PrefixPools[1]', 'prefixLength')
                                                                            if 'Step' in device_group_name:
                                                                                self.config_multivalueObj(device_group_name['Step'],
                                                                                    devicegroupinfo['xpath'] + '/networkGroup[1]'+'/ipv6PrefixPools[1]', 'prefixAddrStep')
                                                                            if 'Address Count' in device_group_name:
                                                                                self.config_multivalueObj(device_group_name['Address Count'],
                                                                                    devicegroupinfo['xpath'] + '/networkGroup[1]'+'/ipv6PrefixPools[1]', 'numberOfAddressesAsy')
                                                                            if 'Protocol' in device_group_name:
                                                                                if device_group_name['Protocol'] == "ospfv3":
                                                                                    protocol = {"xpath": (devicegroupinfo['xpath'] + '/networkGroup[1]'+'/ipv6PrefixPools[1]/connector'),
                                                                                                  "connectedTo": (ethernetinfo['xpath'] + '/ipv6[1]' + '/ospfv3[1]')}
                                                                                    self.config.append({key: value for key, value in protocol.items()})
                                                                                if device_group_name['Protocol'] == "bgpv6":
                                                                                    protocol = {"xpath": (devicegroupinfo['xpath'] + '/networkGroup[1]'+'/ipv6PrefixPools[1]/connector'),
                                                                                                  "connectedTo": (ethernetinfo['xpath'] + '/ipv6[1]' + '/bgpIpv6Peer[1]')}
                                                                                    self.config.append({key: value for key, value in protocol.items()})
                                                                                if device_group_name['Protocol'] == "isis":
                                                                                    protocol = {"xpath": (devicegroupinfo['xpath'] + '/networkGroup[1]'+'/ipv6PrefixPools[1]/connector'),
                                                                                                  "connectedTo": (ethernetinfo['xpath'] + '/ipv6[1]' + '/isisL3[1]')}
                                                                                    self.config.append({key: value for key, value in protocol.items()})
                                                                        else:
                                                                            if 'Ipv6 Address' in device_group_name and device_group_name['Ipv6 Address'] != None:
                                                                                self.config_multivalueObj(device_group_name['Ipv6 Address'],
                                                                                                          devicegroupinfo['xpath'] + '/networkGroup[2]'+'/ipv6PrefixPools[1]', 'networkAddress')
                                                                            if 'Ipv6 Prefix' in device_group_name and device_group_name['Ipv6 Prefix'] != None:
                                                                                self.config_multivalueObj(device_group_name['Ipv6 Prefix'],
                                                                                                          devicegroupinfo['xpath'] + '/networkGroup[2]'+'/ipv6PrefixPools[1]', 'prefixLength')
                                                                            if 'Step' in device_group_name:
                                                                                self.config_multivalueObj(device_group_name['Step'],
                                                                                    devicegroupinfo['xpath'] + '/networkGroup[2]'+'/ipv6PrefixPools[1]', 'prefixAddrStep')
                                                                            if 'Address Count' in device_group_name:
                                                                                self.config_multivalueObj(device_group_name['Address Count'],
                                                                                    devicegroupinfo['xpath'] + '/networkGroup[2]'+'/ipv6PrefixPools[1]', 'numberOfAddressesAsy')
                                                                            if 'Protocol' in device_group_name:
                                                                                if device_group_name['Protocol'] == "ospfv3":
                                                                                    protocol = {"xpath": (devicegroupinfo['xpath'] + '/networkGroup[2]'+'/ipv6PrefixPools[1]/connector'),
                                                                                                  "connectedTo": (ethernetinfo['xpath'] + '/ipv6[1]' + '/ospfv3[1]')}
                                                                                    self.config.append({key: value for key, value in protocol.items()})
                                                                                if device_group_name['Protocol'] == "bgpv6":
                                                                                    protocol = {"xpath": (devicegroupinfo['xpath'] + '/networkGroup[2]'+'/ipv6PrefixPools[1]/connector'),
                                                                                                  "connectedTo": (ethernetinfo['xpath'] + '/ipv6[1]' + '/bgpIpv6Peer[1]')}
                                                                                    self.config.append({key: value for key, value in protocol.items()})
                                                                                if device_group_name['Protocol'] == "isis":
                                                                                    protocol = {"xpath": (devicegroupinfo['xpath'] + '/networkGroup[2]'+'/ipv6PrefixPools[1]/connector'),
                                                                                                  "connectedTo": (ethernetinfo['xpath'] + '/ipv6[1]' + '/isisL3[1]')}
                                                                                    self.config.append({key: value for key, value in protocol.items()})

                                        devicegroupIndex = devicegroupIndex + 1

        self.ixnetwork.ResourceManager.ImportConfig(json.dumps(self.config), True)

    def create_trafficitems(self, workbook_name):
        excel = excelReader(workbook_name)
        status_dict = excel.obtain_status()
        Worksheet_Dict = excel.get_worksheet_data()
        if 'Traffic' in status_dict:
            if status_dict['Traffic'] == True:
                print("Generating Traffic Items")
                endpointSetDict = dict()
                configElementDict = dict()
                traffic_items = dict()
                configElementDict = dict()
                frameSizeDict = dict()
                frameRateDict = dict()

                traffic = []
                if 'Traffic' in Worksheet_Dict:
                    for traffic_item in Worksheet_Dict['Traffic']:
                        if not traffic_item:
                            break
                        self.ixnetwork.info(f"Create Traffic Item {traffic_item['Traffic name']}")
                        if traffic_item['bi-directional'].lower() == 'yes':
                            bidi = True
                        else:
                            bidi = False
                        trafficitemNumber = Worksheet_Dict['Traffic'].index(traffic_item) + 1
                        endpointIndex = 1
                        replacementString = self.ixnetwork.href
                        replacementString = replacementString.rstrip('/')
                        destinationXpath = []
                        sourceXpath = []
                        if 'enable' in traffic_item:
                            if traffic_item['enable'] != '':
                                Enabled = traffic_item['enable']
                            else:
                                Enabled = 'true'
                        else:
                            Enabled = 'true'
                        traffic_items.update({"xpath": "/traffic/trafficItem[" + str(trafficitemNumber) + "]",
                                              "name": traffic_item['Traffic name'],
                                              "biDirectional": bidi, "enabled": Enabled,
                                              "trafficType": traffic_item['Type']})
                        traffic.append({key: value for key, value in traffic_items.items()})
                        if traffic_item['Type'].lower() == "ipv4":
                            if 'Source' not in traffic_item or traffic_item['Source'] == '':
                                pass
                            else:
                                if ';' in traffic_item['Source']:
                                    sourcexpathList = traffic_item['Source'].split(';')
                                else:
                                    sourcexpathList = [traffic_item['Source']]
                                for sourcePath in sourcexpathList:
                                    try:
                                        sourcexpath = self.ixnetwork.Topology.find().DeviceGroup.find(
                                            Name='^' + sourcePath + '$').Ethernet.find().Ipv4.find().href
                                        sourcexpath = sourcexpath.replace(replacementString, '')
                                        if 'ipv4/2' in sourcexpath:
                                            sourcexpath = sourcexpath.replace('ipv4/2','ipv4/1')
                                        sourceXpath.append(re.sub(r'/(\d+)', r'[\1]', sourcexpath))
                                    except:
                                        pass
                            if 'Destination' not in traffic_item or traffic_item['Destination'] == '':
                                pass
                            else:
                                if ';' in traffic_item['Destination']:
                                    destxpathList = traffic_item['Destination'].split(';')
                                else:
                                    destxpathList = [traffic_item['Destination']]
                                for destPath in destxpathList:
                                    try:
                                        destxpath = self.ixnetwork.Topology.find().DeviceGroup.find(
                                            Name='^' + destPath + '$').Ethernet.find().Ipv4.find().href
                                        destxpath = destxpath.replace(replacementString, '')
                                        if 'ipv4/2' in destxpath:
                                            destxpath = destxpath.replace('ipv4/2','ipv4/1')
                                        destinationXpath.append(re.sub(r'/(\d+)', r'[\1]', destxpath))
                                    except:
                                        pass
                        if traffic_item['Type'].lower() == "ipv6":
                            if 'Source' not in traffic_item or traffic_item['Source'] == '':
                                pass
                            else:
                                if ';' in traffic_item['Source']:
                                    sourcexpathList = traffic_item['Source'].split(';')
                                else:
                                    sourcexpathList = [traffic_item['Source']]
                                for sourcePath in sourcexpathList:
                                    try:
                                        sourcexpath = self.ixnetwork.Topology.find().DeviceGroup.find(
                                            Name='^' + sourcePath + '$').Ethernet.find().Ipv6.find().href
                                        sourcexpath = sourcexpath.replace(replacementString, '')
                                        if 'ipv6/2' in sourcexpath:
                                            sourcexpath = sourcexpath.replace('ipv6/2','ipv6/1')
                                        sourceXpath.append(re.sub(r'/(\d+)', r'[\1]', sourcexpath))
                                    except:
                                        pass
                            if 'Destination' not in traffic_item or traffic_item['Destination'] == '':
                                pass
                            else:
                                if ';' in traffic_item['Destination']:
                                    destxpathList = traffic_item['Destination'].split(';')
                                else:
                                    destxpathList = [traffic_item['Destination']]
                                for destPath in destxpathList:
                                    try:
                                        destxpath = self.ixnetwork.Topology.find().DeviceGroup.find(
                                            Name='^' + destPath + '$').Ethernet.find().Ipv6.find().href
                                        destxpath = destxpath.replace(replacementString, '')
                                        if 'ipv6/2' in destxpath:
                                            destxpath = destxpath.replace('ipv6/2','ipv6/1')
                                        destinationXpath.append(re.sub(r'/(\d+)', r'[\1]', destxpath))
                                    except:
                                        pass
                        if destinationXpath != [] and sourceXpath != []:
                            endpointSetDict.update(
                                {"xpath": (traffic_items['xpath'] + "/endpointSet[" + str(endpointIndex) + "]"),
                                 "multicastDestinations": [],
                                 "name": "EndpointSet" + str(endpointIndex),
                                 "sources": sourceXpath,
                                 "destinations": destinationXpath
                                 })
                            traffic.append({key: value for key, value in endpointSetDict.items()})

                        if 'framesize' in traffic_item and traffic_item['framesize'] != '':
                            frameSizeDict.update({"xpath": traffic_items['xpath'] + "/configElement[1]/frameSize",
                                                  "fixedSize": traffic_item['framesize']})
                            traffic.append({key: value for key, value in frameSizeDict.items()})
                        if 'rate' in traffic_item and traffic_item['rate'] != '':
                            frameRateDict.update({"xpath": traffic_items['xpath'] + "/configElement[1]/frameRate",
                                                  "rate": traffic_item['rate']})
                        traffic.append({key: value for key, value in frameRateDict.items()})
                        if 'tracking' in traffic_item and traffic_item['rate'] != '':
                            trackinginfo = dict()
                            trackinginfo.update({'xpath': traffic_items['xpath'] + "/tracking",
                                                 "trackBy": [
                                                     "sourceDestEndpointPair0",
                                                     "trackingenabled0",
                                                     "vlanVlanId0"
                                                 ]})
                            traffic.append({key: value for key, value in trackinginfo.items()})
                        # Need to add config element stack ipv4, ipv6,tcp,udp
                        configElementStack = []
                        if status_dict['packet_editor'] == True:
                            layer4Index = 4
                            for packetInfo in Worksheet_Dict['packet_editor']:
                                if packetInfo['Traffic name'] == traffic_items['name']:
                                    if 'Type' in packetInfo:
                                        if ';' in packetInfo['Type']:
                                            prtotocolTypeList = [packetType.lower() for packetType in packetInfo['Type'].split(';')]
                                        else:
                                            prtotocolTypeList = [packetInfo['Type'].lower()]
                                        for protocolType in prtotocolTypeList:

                                        # if packetInfo['Type'].lower() == 'udp':
                                            udpXpathDict = dict()
                                            udpSourcePortDict = dict()
                                            udpDestinationPortDict = dict()
                                            udpFields = []
                                            if 'Source Port' in packetInfo and 'Destination Port' in packetInfo:
                                                if packetInfo['Source Port'] != '' and packetInfo['Destination Port'] != '':
                                                    udpPortsAuto = 'true'
                                                sourcePort = packetInfo['Source Port']
                                                destinationPort = packetInfo['Destination Port']
                                            else:
                                                udpPortsAuto = 'false'
                                                sourcePort = 63
                                                destinationPort = 123
                                            udpSourcePortDict.update({"xpath": traffic_items['xpath'] + "/configElement[1]" + "/stack[@alias = " + protocolType + "-" + str(layer4Index + prtotocolTypeList.index(protocolType)) + "]/field[@alias = " + protocolType +".header.srcPort-1]",
                                                                      "singleValue": sourcePort,
                                                                      "fieldValue": 'Default',
                                                                      "stepValue": "1",
                                                                      "valueType": "increment", "auto": udpPortsAuto,
                                                                      "startValue": sourcePort,
                                                                      "countValue": "1"})
                                            udpFields.append({key: value for key, value in udpSourcePortDict.items()})
                                            udpDestinationPortDict.update({"xpath": traffic_items['xpath'] + "/configElement[1]" + "/stack[@alias = "+ protocolType + "-" + str(layer4Index + prtotocolTypeList.index(protocolType)) + "]/field[@alias = "+ protocolType +".header.dstPort-2]",
                                                                           "singleValue": destinationPort,
                                                                           "fieldValue": 'Default',
                                                                           "stepValue": "1", "auto": udpPortsAuto,
                                                                           "valueType": "increment",
                                                                           "startValue": destinationPort, "countValue": "1"})
                                            udpFields.append({key: value for key, value in udpDestinationPortDict.items()})
                                            udpXpathDict.update({"xpath": traffic_items['xpath'] + "/configElement[1]" + "/stack[@alias = "+ protocolType + "-" + str(layer4Index + prtotocolTypeList.index(protocolType)) + "]",
                                                                 "field": udpFields})
                                            configElementStack.append({key: value for key, value in udpXpathDict.items()})

                        configElementDict.update({"xpath": traffic_items['xpath'] + "/configElement[1]",
                                                  "crc": "goodCrc",
                                                  "preambleCustomSize": 8,
                                                  "enableDisparityError": 'false',
                                                  "preambleFrameSizeMode": "auto",
                                                  "destinationMacMode": "manual",
                                                  "stack": configElementStack
                                                  })
                        traffic.append({key: value for key, value in configElementDict.items()})
        self.ixnetwork.ResourceManager.ImportConfig(json.dumps(traffic), False)


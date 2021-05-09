import pandas as pd

class API_Sheet():
    def __init__(self, build_information='default', base='default', physical='default', devicegroup='default', ipv4_ethernet='default',
                 ipv6_ethernet='default', ipv4_bgp='default', ipv6_bgp='default', ipv4_loopback='default', bgp_capabilities='default', ipv4_ospf='default',
                 ipv6_ospf='default', isis='default', networkgroup='default', igmp_host='default', igmp_querier='default', traffic='default', packet_editor='default'):
        self.Build_information = build_information
        self.Base = base
        self.Physical = physical
        self.Devicegroup = devicegroup
        self.IPv4_Ethernet = ipv4_ethernet
        self.IPv6_Ethernet = ipv6_ethernet
        self.IPv4_BGP = ipv4_bgp
        self.IPv6_BGP = ipv6_bgp
        self.IPv4_Loopback = ipv4_loopback
        self.BGP_Capabilities = bgp_capabilities
        self.IPv4_OSPF = ipv4_ospf
        self.IPv6_OSPF = ipv6_ospf
        self.ISIS = isis
        self.Network_Group = networkgroup
        self.IGMP_Host = igmp_host
        self.IGMP_Querier = igmp_querier
        self.Traffic = traffic
        self.packet_editor = packet_editor

    def create_excel(self,excelname='Generate_IxNetwork_Config.xlsx'):
        if self.Build_information == 'default':
            build_data = pd.DataFrame({'Include': ['yes']*17, 'Description':['Base Variables','Assigning Physical Ports and Topology','Assigning Device Groups to Topologies',\
                                                                             'Configuring IPv4 and Ethernet Information','Configuring IPv6 and Ethernet Information','IPv4 BGP Configuration','IPv6 BGP Configuration',\
                                                                             'IPv4 BGP on Loopback Configuation','BGP Capabilities','IPv4 OSPF Configuration','IPv6 OSPF Configuration','ISIS Configuration','Network Group Configuration',\
                                                                             'IGMP Senders group Configuration','IGMP Receivers Group Configuration','Traffic Flow setup','Packet editor'],\
                                       'Worksheet':['Base','Physical','Devicegroup','IPv4_Ethernet','IPv6_Ethernet','IPv4_BGP','IPv6_BGP','IPv4_Loopback_BGP','BGP_Capabilities',\
                                                    'IPv4_OSPF','IPv6_OSPF','ISIS','Network_Group','IGMP_Host','IGMP_Querier','Traffic','packet_editor'], 'Status':['Completed']*17})
        else:
            build_data = pd.DataFrame(self.Build_information)
        if self.Base == 'default':
            base_data = pd.DataFrame({'Platform':['windows'], 'API Server IP':['127.0.0.1'], 'API Server Port':[11009],'Username':['admin'],\
                                      'Password':['admin'],'Licensing Server IP':['10.39.70.159'],'License Mode':['perpetual'],'License Tier':['tier3'],\
                                      'Debug Mode':['False'],'Force  Port Ownership':['False']})

        else:
            base_data = pd.DataFrame(self.Base)
        if self.Physical == 'default':
            physical_data = pd.DataFrame({'Chassis IP':['10.39.70.159', '10.39.70.159'],'Linecard Number':[1, 1], 'Port Number':[1, 2], 'Port Name':['Ethernet - 001', 'Ethernet - 002'], 'Topology Name':['Topology 1', 'Topology 2']})
        else:
            physical_data = pd.DataFrame(self.Physical)
        if self.Devicegroup == 'default':
            devicegroup_data = pd.DataFrame({'Topology':['Topology 1', 'Topology 2'], 'Device Group':['Device Group 1', 'Device Group 2'], 'Multiplier':[1, 1], \
                                             'Vlan Header':[1, 1]})
        else:
            devicegroup_data = pd.DataFrame(self.Devicegroup)
        if self.IPv4_Ethernet == 'default':
            ipv4_data = pd.DataFrame({'Device Group':['Device Group 1', 'Device Group 2']})
        else:
            ipv4_data = pd.DataFrame(self.IPv4_Ethernet)
        if self.IPv6_Ethernet == 'default':
            ipv6_data = pd.DataFrame({'Device Group':['Device Group 1', 'Device Group 2']})
        else:
            ipv6_data = pd.DataFrame(self.IPv6_Ethernet)
        if self.IPv4_BGP == 'default':
            bgpv4_data = pd.DataFrame({'Device Group':['Device Group 1', 'Device Group 2']})
        else:
            bgpv4_data = pd.DataFrame(self.IPv4_BGP)
        if self.IPv6_BGP == 'default':
            bgpv6_data = pd.DataFrame({'Device Group':['Device Group 1', 'Device Group 2']})
        else:
            bgpv6_data = pd.DataFrame(self.IPv6_BGP)
        if self.BGP_Capabilities == 'default':
            bgpcap_data = pd.DataFrame({'Device Group':['Device Group 1', 'Device Group 2']})
        else:
            bgpcap_data = pd.DataFrame(self.BGP_Capabilities)
        if self.IPv4_Loopback == 'default':
            loopback_data = pd.DataFrame({'Device Group': ['Device Group 1', 'Device Group 2']})
        else:
            loopback_data = pd.DataFrame(self.IPv4_Loopback)
        if self.IPv4_OSPF == 'default':
            ospfv2_data = pd.DataFrame({'Device Group': ['Device Group 1', 'Device Group 2']})
        else:
            ospfv2_data = pd.DataFrame(self.IPv4_OSPF)
        if self.IPv6_OSPF == 'default':
            ospfv3_data = pd.DataFrame({'Device Group': ['Device Group 1', 'Device Group 2']})
        else:
            ospfv3_data = pd.DataFrame(self.IPv6_OSPF)
        if self.ISIS == 'default':
            isis_data = pd.DataFrame({'Device Group': ['Device Group 1', 'Device Group 2']})
        else:
            isis_data = pd.DataFrame(self.ISIS)
        if self.Network_Group == 'default':
            network_data = pd.DataFrame({'Device Group': ['Device Group 1', 'Device Group 2'], 'Name': ['Network Group 1', 'Network Group 2'], 'IP Version': ['ipv4', 'ipv4'], 'Multiplier':[1, 1], 'Protocol':['ospfv2', 'ospfv2']})
        else:
            network_data = pd.DataFrame(self.Network_Group)
        if self.IGMP_Host == 'default':
            host_data = pd.DataFrame({'Device Group': ['Device Group 1', 'Device Group 2']})
        else:
            host_data = pd.DataFrame(self.IGMP_Host)
        if self.IGMP_Querier == 'default':
            querier_data = pd.DataFrame({'Device Group': ['Device Group 1', 'Device Group 2']})
        else:
            querier_data = pd.DataFrame(self.IGMP_Querier)
        if self.Traffic == 'default':
            traffic_data = pd.DataFrame({'Traffic name': ['Device Group 1'], 'Type':['ipv4'], 'bi-directional': ['yes'], 'Source': ['Device Group 1'], 'Destination': ['Device Group 2']})
        else:
            traffic_data = pd.DataFrame(self.Traffic)
        if self.packet_editor == 'default':
            packet_data = pd.DataFrame({'Traffic name': ['Device Group 1'], 'Type':'TCP;UDP'})
        else:
            packet_data = pd.DataFrame(self.packet_editor)

        writer = pd.ExcelWriter(excelname, engine='xlsxwriter')
        build_data.to_excel(writer, sheet_name='Build_Information', index=False)
        base_data.to_excel(writer, sheet_name='Base', index=False)
        physical_data.to_excel(writer, sheet_name='Physical', index=False)
        devicegroup_data.to_excel(writer, sheet_name='Devicegroup', index=False)
        ipv4_data.to_excel(writer, sheet_name='IPv4_Ethernet', index=False)
        ipv6_data.to_excel(writer, sheet_name='IPv6_Ethernet', index=False)
        bgpv4_data.to_excel(writer, sheet_name='IPv4_BGP', index=False)
        bgpv6_data.to_excel(writer, sheet_name='IPv6_BGP', index=False)
        bgpcap_data.to_excel(writer, sheet_name='BGP_Capabilities', index=False)
        loopback_data.to_excel(writer, sheet_name='IPv4_Loopback_BGP', index=False)
        ospfv2_data.to_excel(writer, sheet_name='IPv4_OSPF', index=False)
        ospfv3_data.to_excel(writer, sheet_name='IPv6_OSPF', index=False)
        isis_data.to_excel(writer, sheet_name='ISIS', index=False)
        network_data.to_excel(writer, sheet_name='Network_Group', index=False)
        host_data.to_excel(writer, sheet_name='IGMP_Host', index=False)
        querier_data.to_excel(writer, sheet_name='IGMP_Querier', index=False)
        traffic_data.to_excel(writer, sheet_name='Traffic', index=False)
        packet_data.to_excel(writer, sheet_name='packet_editor', index=False)

        writer.save()

        return excelname
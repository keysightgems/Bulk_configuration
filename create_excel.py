import pandas as pd

class API_Sheet():
    def __init__(self, build_information=None, base=None, physical=None, devicegroup=None, ipv4_ethernet=None,
                 ipv6_ethernet=None, ipv4_bgp=None, ipv6_bgp=None, ipv4_loopback=None, bgp_capabilities=None, ipv4_ospf=None,
                 ipv6_ospf=None, isis=None, dhcp_ipv4=None, ldp=None, dhcp_ipv6=None, dhcp_serverv4=None, dhcp_serverv6=None, networkgroup=None, igmp_host=None, igmp_querier=None, mld_host=None, mld_querier=None, traffic=None, packet_editor=None, **kwargs):
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
        self.DHCP_Ipv4 = dhcp_ipv4
        self.DHCP_Ipv6 = dhcp_ipv6
        self.LDP = ldp
        self.DHCP_Serverv4 = dhcp_serverv4
        self.DHCP_Serverv6 = dhcp_serverv6
        self.Network_Group = networkgroup
        self.IGMP_Host = igmp_host
        self.IGMP_Querier = igmp_querier
        self.MLD_Querier = mld_querier
        self.MLD_Host = mld_host
        self.Traffic = traffic
        self.packet_editor = packet_editor

    def create_excel(self,excelname='Generate_IxNetwork_Config.xlsx'):
        if self.Build_information == 'default' or self.Build_information == None:
            build_data = pd.DataFrame({'Include': ['yes']*24, 'Description':['Base Variables','Assigning Physical Ports and Topology','Assigning Device Groups to Topologies',\
                                                                             'Configuring IPv4 and Ethernet Information','Configuring IPv6 and Ethernet Information','IPv4 BGP Configuration','IPv6 BGP Configuration',\
                                                                             'IPv4 BGP on Loopback Configuation','BGP Capabilities','IPv4 OSPF Configuration','IPv6 OSPF Configuration','ISIS Configuration','Network Group Configuration',\
                                                                             'IGMP Senders group Configuration','IGMP Receivers Group Configuration','Traffic Flow setup','Packet editor', 'DHCP_Ipv4', 'DHCP_Ipv6', 'DHCP_Serverv4', 'DHCP_Serverv6', 'LDP', 'MLD_Host', 'MLD_Querier'],\
                                       'Worksheet':['Base','Physical','Devicegroup','IPv4_Ethernet','IPv6_Ethernet','IPv4_BGP','IPv6_BGP','IPv4_Loopback_BGP','BGP_Capabilities',\
                                                    'IPv4_OSPF','IPv6_OSPF','ISIS','Network_Group','IGMP_Host','IGMP_Querier','Traffic','packet_editor','DHCP_Ipv4','DHCP_Ipv6','DHCP_Serverv4','DHCP_Serverv6', 'LDP', 'MLD_Host', 'MLD_Querier'], 'Status':['Completed']*24})
            # build_data = pd.DataFrame(self.buildData)
        else:
            if self.Build_information != None:
                build_data = pd.DataFrame(self.Build_information)
        if self.Base == 'default':
            base_data = pd.DataFrame({'Platform':['windows'], 'API Server IP':['127.0.0.1'], 'API Server Port':[11009],'Username':['admin'],\
                                      'Password':['admin'],'Licensing Server IP':['10.39.70.159'],'License Mode':['perpetual'],'License Tier':['tier3'],\
                                      'Debug Mode':['False'],'Force  Port Ownership':['False']})

        else:
            if self.Base != None:
                base_data = pd.DataFrame(self.Base)
        if self.Physical == 'default':
            physical_data = pd.DataFrame({'Chassis IP':['10.39.70.159', '10.39.70.159'],'Linecard Number':[1, 1], 'Port Number':[1, 2], 'Port Name':['Ethernet - 001', 'Ethernet - 002'], 'Topology Name':['Topology 1', 'Topology 2']})
        else:
            if self.Physical != None:
                physical_data = pd.DataFrame(self.Physical)
        if self.Devicegroup == 'default':
            devicegroup_data = pd.DataFrame({'Topology':['Topology 1', 'Topology 2'], 'Device Group':['Device Group 1', 'Device Group 2'], 'Multiplier':[1, 1], \
                                             'Vlan Header':[1, 1]})
        else:
            if self.Devicegroup != None:
                devicegroup_data = pd.DataFrame(self.Devicegroup)
        if self.IPv4_Ethernet == 'default':
            ipv4_data = pd.DataFrame({'Device Group':['Device Group 1', 'Device Group 2']})
        else:
            if self.IPv4_Ethernet != None:
                ipv4_data = pd.DataFrame(self.IPv4_Ethernet)
        if self.IPv6_Ethernet == 'default':
            ipv6_data = pd.DataFrame({'Device Group':['Device Group 1', 'Device Group 2']})
        else:
            if self.IPv6_Ethernet != None:
                ipv6_data = pd.DataFrame(self.IPv6_Ethernet)
        if self.IPv4_BGP == 'default':
            bgpv4_data = pd.DataFrame({'Device Group':['Device Group 1', 'Device Group 2']})
        else:
            if self.IPv4_BGP != None:
                bgpv4_data = pd.DataFrame(self.IPv4_BGP)
        if self.IPv6_BGP == 'default':
            bgpv6_data = pd.DataFrame({'Device Group':['Device Group 1', 'Device Group 2']})
        else:
            if self.IPv6_BGP != None:
                bgpv6_data = pd.DataFrame(self.IPv6_BGP)
        if self.BGP_Capabilities == 'default':
            bgpcap_data = pd.DataFrame({'Device Group':['Device Group 1', 'Device Group 2']})
        else:
            if self.BGP_Capabilities != None:
                bgpcap_data = pd.DataFrame(self.BGP_Capabilities)
        if self.IPv4_Loopback == 'default':
            loopback_data = pd.DataFrame({'Device Group': ['Device Group 1', 'Device Group 2']})
        else:
            if self.IPv4_Loopback != None:
                loopback_data = pd.DataFrame(self.IPv4_Loopback)
        if self.IPv4_OSPF == 'default':
            ospfv2_data = pd.DataFrame({'Device Group': ['Device Group 1', 'Device Group 2']})
        else:
            if self.IPv4_OSPF != None:
                ospfv2_data = pd.DataFrame(self.IPv4_OSPF)
        if self.IPv6_OSPF == 'default':
            ospfv3_data = pd.DataFrame({'Device Group': ['Device Group 1', 'Device Group 2']})
        else:
            if self.IPv6_OSPF != None:
                ospfv3_data = pd.DataFrame(self.IPv6_OSPF)
        if self.ISIS == 'default':
            isis_data = pd.DataFrame({'Device Group': ['Device Group 1', 'Device Group 2']})
        else:
            if self.ISIS != None:
                isis_data = pd.DataFrame(self.ISIS)

        if self.DHCP_Ipv4 == 'default':
            dhcp_data = pd.DataFrame({'Device Group': ['Device Group 1']})
        else:
            if self.DHCP_Ipv4 != None:
                dhcp_data = pd.DataFrame(self.DHCP_Ipv4)
        if self.DHCP_Serverv4 == 'default':
            dhcpserverv4_data = pd.DataFrame({'Device Group': ['Device Group 2']})
        else:
            if self.DHCP_Serverv4 != None:
                dhcpserverv4_data = pd.DataFrame(self.DHCP_Serverv4)
        if self.DHCP_Ipv6 == 'default':
            dhcpv6_data = pd.DataFrame({'Device Group': ['Device Group 1', 'Device Group 2']})
        else:
            if self.DHCP_Ipv6 != None:
                dhcpv6_data = pd.DataFrame(self.DHCP_Ipv6)
        if self.DHCP_Serverv6 == 'default':
            dhcpserverv6_data = pd.DataFrame({'Device Group': ['Device Group 2']})
        else:
            if self.DHCP_Serverv6 != None:
                dhcpserverv6_data = pd.DataFrame(self.DHCP_Serverv6)
        if self.LDP == 'default':
            ldp_data = pd.DataFrame({'Device Group': ['Device Group 1', 'Device Group 2'], 'IP Version':['ipv4','ipv6']})
        else:
            if self.LDP != None:
                ldp_data = pd.DataFrame(self.LDP)
        if self.Network_Group == 'default':
            network_data = pd.DataFrame({'Device Group': ['Device Group 1', 'Device Group 2'], 'Name': ['Network Group 1', 'Network Group 2'], 'IP Version': ['ipv4', 'ipv4'], 'Multiplier':[1, 1], 'Protocol':['ospfv2', 'ospfv2']})
        else:
            if self.Network_Group != None:
                network_data = pd.DataFrame(self.Network_Group)
        if self.IGMP_Host == 'default':
            host_data = pd.DataFrame({'Device Group': ['Device Group 1']})
        else:
            if self.IGMP_Host != None:
                host_data = pd.DataFrame(self.IGMP_Host)
        if self.MLD_Host == 'default':
            mldhost_data = pd.DataFrame({'Device Group': ['Device Group 1']})
        else:
            if self.MLD_Host != None:
                mldhost_data = pd.DataFrame(self.MLD_Host)

        if self.IGMP_Querier == 'default':
            querier_data = pd.DataFrame({'Device Group': ['Device Group 2']})
        else:
            if self.IGMP_Querier != None:
                querier_data = pd.DataFrame(self.IGMP_Querier)

        if self.MLD_Querier == 'default':
            mldquerier_data = pd.DataFrame({'Device Group': ['Device Group 2']})
        else:
            if self.MLD_Querier != None:
                mldquerier_data = pd.DataFrame(self.MLD_Querier)
        if self.Traffic == 'default':
            traffic_data = pd.DataFrame({'Traffic name': ['Device Group 1'], 'Type':['ipv4'], 'bi-directional': ['yes'], 'Source': ['Device Group 1'], 'Destination': ['Device Group 2']})
        else:
            if self.Traffic != None:
                traffic_data = pd.DataFrame(self.Traffic)
        if self.packet_editor == 'default':
            packet_data = pd.DataFrame({'Traffic name': ['Device Group 1'], 'Type':'TCP'})
        else:
            if self.packet_editor != None:
                packet_data = pd.DataFrame(self.packet_editor)

        writer = pd.ExcelWriter(excelname, engine='xlsxwriter')
        if self.Build_information != None or self.Build_information == None:
            build_data.to_excel(writer, sheet_name='Build_Information', index=False)
        if self.Base != None:
            base_data.to_excel(writer, sheet_name='Base', index=False)
        if self.Physical != None:
            physical_data.to_excel(writer, sheet_name='Physical', index=False)
        if self.Devicegroup != None:
            devicegroup_data.to_excel(writer, sheet_name='Devicegroup', index=False)
        if self.IPv4_Ethernet != None:
            ipv4_data.to_excel(writer, sheet_name='IPv4_Ethernet', index=False)
        if self.IPv6_Ethernet != None:
            ipv6_data.to_excel(writer, sheet_name='IPv6_Ethernet', index=False)
        if self.IPv4_BGP != None:
            bgpv4_data.to_excel(writer, sheet_name='IPv4_BGP', index=False)
        if self.IPv6_BGP != None:
            bgpv6_data.to_excel(writer, sheet_name='IPv6_BGP', index=False)
        if self.BGP_Capabilities != None:
            bgpcap_data.to_excel(writer, sheet_name='BGP_Capabilities', index=False)
        if self.IPv4_Loopback != None:
            loopback_data.to_excel(writer, sheet_name='IPv4_Loopback_BGP', index=False)
        if self.IPv4_OSPF != None:
            ospfv2_data.to_excel(writer, sheet_name='IPv4_OSPF', index=False)
        if self.IPv6_OSPF != None:
            ospfv3_data.to_excel(writer, sheet_name='IPv6_OSPF', index=False)
        if self.ISIS != None:
            isis_data.to_excel(writer, sheet_name='ISIS', index=False)
        if self.DHCP_Ipv4 != None:
            dhcp_data.to_excel(writer, sheet_name='DHCP_Ipv4', index=False)
        if self.DHCP_Serverv4 !=None:
            dhcpserverv4_data.to_excel(writer, sheet_name='DHCP_Serverv4', index=False)
        if self.DHCP_Ipv6 != None:
            dhcpv6_data.to_excel(writer, sheet_name='DHCP_Ipv6', index=False)
        if self.DHCP_Serverv6 !=None:
            dhcpserverv6_data.to_excel(writer, sheet_name='DHCP_Serverv6', index=False)
        if self.LDP !=None:
            ldp_data.to_excel(writer, sheet_name='LDP', index=False)
        if self.Network_Group != None:
            network_data.to_excel(writer, sheet_name='Network_Group', index=False)
        if self.IGMP_Host != None:
            host_data.to_excel(writer, sheet_name='IGMP_Host', index=False)
        if self.MLD_Host != None:
            mldhost_data.to_excel(writer, sheet_name='MLD_Host', index=False)
        if self.IGMP_Querier != None:
            querier_data.to_excel(writer, sheet_name='IGMP_Querier', index=False)
        if self.MLD_Querier != None:
            mldquerier_data.to_excel(writer, sheet_name='MLD_Querier', index=False)
        if self.Traffic != None:
            traffic_data.to_excel(writer, sheet_name='Traffic', index=False)
        if self.packet_editor != None:
            packet_data.to_excel(writer, sheet_name='packet_editor', index=False)

        writer.save()

        return excelname
# Generate nodes sheet

def gen_nodes_sheet(rc, xls_wb):
    # Get nodes config
    nodes = rc.cluster.list_nodes()

    # Create and format 1st sheet for nodes config
    ws_nodes = xls_wb.active
    ws_nodes.title = 'Nodes'
    ws_nodes.column_dimensions['A'].width = 1
    ws_nodes.row_dimensions[1].height = 5
    ws_nodes.column_dimensions['B'].width = 25
    ws_nodes.column_dimensions['C'].width = 25
    ws_nodes.column_dimensions['D'].width = 25
    ws_nodes.column_dimensions['E'].width = 25

    # Create title for table

    act_row = 2
    ws_nodes.cell(row=act_row, column=2, value='Node Name').style = 'style_title'
    ws_nodes.cell(row=act_row, column=3, value='Model Number').style = 'style_title'
    ws_nodes.cell(row=act_row, column=4, value='Serial Number').style = 'style_title'
    ws_nodes.cell(row=act_row, column=5, value='Status').style = 'style_title'
    act_row = act_row + 1

    # Filling table with nodes information

    for node in nodes:
        ws_nodes.cell(row=act_row, column=2, value=node['node_name']).style = 'style_normal'
        ws_nodes.cell(row=act_row, column=3, value=node['serial_number']).style = 'style_normal'
        ws_nodes.cell(row=act_row, column=4, value=node['model_number']).style = 'style_normal'
        ws_nodes.cell(row=act_row, column=5, value=node['node_status']).style = 'style_normal'
        act_row = act_row + 1

# Generate disks sheet


def gen_disks_sheet(rc, xls_wb):
    # Get disks config
    disks = rc.cluster.get_cluster_slots_status()

    # Create and format 2nd sheet for disks config
    ws_slot = xls_wb.create_sheet(title="Disks")
    ws_slot.column_dimensions['A'].width = 1
    ws_slot.row_dimensions[1].height = 5
    ws_slot.column_dimensions['B'].width = 10
    ws_slot.column_dimensions['C'].width = 25
    ws_slot.column_dimensions['D'].width = 25
    ws_slot.column_dimensions['E'].width = 10
    ws_slot.column_dimensions['F'].width = 10
    ws_slot.column_dimensions['G'].width = 10
    ws_slot.column_dimensions['H'].width = 10
    ws_slot.column_dimensions['I'].width = 10
    ws_slot.column_dimensions['J'].width = 10

    # Create title for table

    act_row = 2
    ws_slot.cell(row=act_row, column=2, value='Node id').style = 'style_title'
    ws_slot.cell(row=act_row, column=3, value='Disk Model').style = 'style_title'
    ws_slot.cell(row=act_row, column=4, value='Serial Number').style = 'style_title'
    ws_slot.cell(row=act_row, column=5, value='Disk Type').style = 'style_title'
    ws_slot.cell(row=act_row, column=6, value='Drive Bay').style = 'style_title'
    ws_slot.cell(row=act_row, column=7, value='Raw Capacity (GB)').style = 'style_title'
    ws_slot.cell(row=act_row, column=8, value='Slot').style = 'style_title'
    ws_slot.cell(row=act_row, column=9, value='Slot Type').style = 'style_title'
    ws_slot.cell(row=act_row, column=10, value='State').style = 'style_title'
    act_row = act_row + 1

    # Filling table with nodes information

    for disk in disks:
        # Convert disk raw capacity in GB ( /1024 /1024 /1024)
        raw_capa = int(disk['raw_capacity']) / 1024 / 1024 / 1024
        ws_slot.cell(row=act_row, column=2, value=disk['node_id']).style = 'style_normal'
        ws_slot.cell(row=act_row, column=3, value=disk['disk_model']).style = 'style_normal'
        ws_slot.cell(row=act_row, column=4, value=disk['disk_serial_number']).style = 'style_normal'
        ws_slot.cell(row=act_row, column=5, value=disk['disk_type']).style = 'style_normal'
        ws_slot.cell(row=act_row, column=6, value=disk['drive_bay']).style = 'style_normal'
        ws_slot.cell(row=act_row, column=7, value=raw_capa).style = 'style_normal'
        ws_slot.cell(row=act_row, column=8, value=disk['slot']).style = 'style_normal'
        ws_slot.cell(row=act_row, column=9, value=disk['slot_type']).style = 'style_normal'
        ws_slot.cell(row=act_row, column=10, value=disk['state']).style = 'style_normal'
        act_row = act_row + 1

# Generate network sheet


def gen_net_sheet(rc, xls_wb):
    # Get interface config
    interfaces = rc.network.list_interfaces()
    # Get network config
    networks = rc.network.list_networks(1)

    # Create and format 3rd sheet for network config
    ws_net = xls_wb.create_sheet(title="Networks")
    ws_net.column_dimensions['A'].width = 1
    ws_net.row_dimensions[1].height = 5
    ws_net.column_dimensions['B'].width = 15
    ws_net.column_dimensions['C'].width = 20

    act_row = 2

    # Only one config but API sends back a list

    for interface in interfaces:
        # Create first table with global interface config (bonding, gateway...) /v2/network/interfaces
        ws_net.merge_cells(start_row=act_row, start_column=2, end_row=act_row, end_column=3)
        ws_net.cell(row=act_row, column=2, value='Global Network settings').style = 'style_title'
        act_row = act_row + 1

        # Bond name
        ws_net.cell(row=act_row, column=2, value='Interface name').style = 'style_title'
        ws_net.cell(row=act_row, column=3, value=interface['name']).style = 'style_normal'
        act_row = act_row + 1

        # Default Gateway ipv4
        ws_net.cell(row=act_row, column=2, value='Default Gateway').style = 'style_title'
        ws_net.cell(row=act_row, column=3, value=interface['default_gateway']).style = 'style_normal'
        act_row = act_row + 1

        # Default Gateway ipv6 if defined
        if interface['default_gateway_ipv6'] != '':
            ws_net.cell(row=act_row, column=2, value='Default Gateway (IPv6)').style = 'style_title'
            ws_net.cell(row=act_row, column=3, value=interface['default_gateway_ipv6']).style = 'style_normal'
            act_row = act_row + 1

        # Bonding mode
        ws_net.cell(row=act_row, column=2, value='Bonding mode').style = 'style_title'
        ws_net.cell(row=act_row, column=3, value=interface['bonding_mode']).style = 'style_normal'
        act_row = act_row + 1

        # MTU
        ws_net.cell(row=act_row, column=2, value='MTU').style = 'style_title'
        ws_net.cell(row=act_row, column=3, value=interface['mtu']).style = 'style_normal'
        act_row = act_row + 1

        # Create a separation line between tables
        act_row = act_row + 1

    # Create a separate table for each network

    for network in networks:
        # Create title for each table containing network name in merged cells
        ws_net.merge_cells(start_row=act_row, start_column=2, end_row=act_row, end_column=3)
        ws_net.cell(row=act_row, column=2, value=network['name']).style = 'style_title'
        act_row = act_row + 1

        # Fill table with network information
        # Get all ip ranges
        ws_net.cell(row=act_row, column=2, value='IP ranges').style = 'style_title'
        ip_range = ''
        for ip_ranges in network['ip_ranges']:
            ip_range = ip_range + ip_ranges + "\n"
        ip_range = ip_range.rstrip("\n")
        ws_net.cell(row=act_row, column=3, value=ip_range).style = 'style_normal'
        act_row = act_row + 1

        # Netmask information
        ws_net.cell(row=act_row, column=2, value='Netmask').style = 'style_title'
        ws_net.cell(row=act_row, column=3, value=network['netmask']).style = 'style_normal'
        act_row = act_row + 1

        # Get all floating ip ranges if exists
        if network['floating_ip_ranges']:
            ws_net.cell(row=act_row, column=2, value='Floating IPs').style = 'style_title'
            float_ips = ''
            for ip_ranges in network['floating_ip_ranges']:
                float_ips = float_ips + ip_ranges + "\n"
            float_ips = float_ips.rstrip("\n")
            ws_net.cell(row=act_row, column=3, value=float_ips).style = 'style_normal'
            act_row = act_row + 1

        # Assignation method
        ws_net.cell(row=act_row, column=2, value='Assignation').style = 'style_title'
        ws_net.cell(row=act_row, column=3, value=network['assigned_by']).style = 'style_normal'
        act_row = act_row + 1

        # Get DNS Servers if exists
        if network['dns_servers']:
            ws_net.cell(row=act_row, column=2, value='DNS Servers').style = 'style_title'
            dns = ''
            for servers in network['dns_servers']:
                dns = dns + servers + "\n"
            dns = dns.rstrip("\n")
            ws_net.cell(row=act_row, column=3, value=dns).style = 'style_normal'
            act_row = act_row + 1

        # Get DNS search domains
        if network['dns_search_domains']:
            ws_net.cell(row=act_row, column=2, value='Search domains').style = 'style_title'
            dns_dom = ''
            for domain in network['dns_search_domains']:
                dns_dom = dns_dom + domain + "\n"
            dns_dom = dns_dom.rstrip("\n")
            ws_net.cell(row=act_row, column=3, value=dns_dom).style = 'style_normal'
            act_row = act_row + 1

        # VLAN Id
        ws_net.cell(row=act_row, column=2, value='VLAN id').style = 'style_title'
        ws_net.cell(row=act_row, column=3, value=network['vlan_id']).style = 'style_normal'
        act_row = act_row + 1

        # Create a thin separation line between tables
        ws_net.row_dimensions[act_row].height = 5
        act_row = act_row + 1

# Generate Active Directory Sheet


def gen_ad_sheet(rc, xls_wb):
    # Get Active Directory status
    ad = rc.ad.list_ad()

    # Create and format 4th sheet for Active Directory Config
    ws_ad = xls_wb.create_sheet(title="Active Directory")
    ws_ad.column_dimensions['A'].width = 1
    ws_ad.row_dimensions[1].height = 5
    ws_ad.column_dimensions['B'].width = 20
    ws_ad.column_dimensions['C'].width = 40
    ws_ad.column_dimensions['D'].width = 20
    ws_ad.column_dimensions['E'].width = 30
    ws_ad.column_dimensions['F'].width = 25
    ws_ad.column_dimensions['G'].width = 40
    ws_ad.column_dimensions['H'].width = 25

    act_row = 2

    # Create tables only if cluster is joined, else indicate status as NOT_IN_DOMAIN

    if ad['status'] != "NOT_IN_DOMAIN":
        # AD status
        ws_ad.cell(row=act_row, column=2, value='AD Status').style = 'style_title'
        ws_ad.cell(row=act_row, column=3, value=ad['status']).style = 'style_normal'
        act_row = act_row + 1

        # AD domain name
        ws_ad.cell(row=act_row, column=2, value='Domain').style = 'style_title'
        ws_ad.cell(row=act_row, column=3, value=ad['domain']).style = 'style_normal'
        act_row = act_row + 1

        # OU if specified
        if ad['ou']:
            ws_ad.cell(row=act_row, column=2, value='OU').style = 'style_title'
            ws_ad.cell(row=act_row, column=3, value=ad['ou']).style = 'style_normal'
            act_row = act_row + 1

        # POSIX Attributes
        ws_ad.cell(row=act_row, column=2, value='Use POSIX Attr').style = 'style_title'
        ws_ad.cell(row=act_row, column=3, value=ad['use_ad_posix_attributes']).style = 'style_normal'
        act_row = act_row + 1

        # NETBIOS Domain
        ws_ad.cell(row=act_row, column=2, value='NETBIOS Domain').style = 'style_title'
        ws_ad.cell(row=act_row, column=3, value=ad['domain_netbios']).style = 'style_normal'
        act_row = act_row + 1

        # Create a thin separation line between tables
        ws_ad.row_dimensions[act_row].height = 5
        act_row = act_row + 1

        # Create a new table for each node
        # Create title for this table
        ws_ad.cell(row=act_row, column=2, value='Node ID').style = 'style_title'
        ws_ad.cell(row=act_row, column=3, value='Bind URI').style = 'style_title'
        ws_ad.cell(row=act_row, column=4, value='KDC Address').style = 'style_title'
        ws_ad.cell(row=act_row, column=5, value='Bind Domain').style = 'style_title'
        ws_ad.cell(row=act_row, column=6, value='Bind Account').style = 'style_title'
        ws_ad.cell(row=act_row, column=7, value='Base DN').style = 'style_title'
        ws_ad.cell(row=act_row, column=8, value='Health').style = 'style_title'
        act_row = act_row + 1

        # Filling table with information from all nodes
        for node in ad['ldap_connection_states']:

            # Get node status
            ws_ad.cell(row=act_row, column=2, value=node['node_id']).style = 'style_normal'

            # Get Servers information
            for server in node['servers']:
                # Get Bind URI
                if server['bind_uri']:
                    ws_ad.cell(row=act_row, column=3, value=server['bind_uri']).style = 'style_normal'
                # Get KDC Address
                if server['kdc_address']:
                    ws_ad.cell(row=act_row, column=4, value=server['kdc_address']).style = 'style_normal'

            # Get binding information
            ws_ad.cell(row=act_row, column=5, value=node['bind_domain']).style = 'style_normal'
            ws_ad.cell(row=act_row, column=6, value=node['bind_account']).style = 'style_normal'

            # Get All Base DN
            base_dn = ''
            for dn in node['base_dn_vec']:
                base_dn = base_dn + dn + "\n"
            base_dn = base_dn.rstrip("\n")
            ws_ad.cell(row=act_row, column=7, value=base_dn).style = 'style_normal'

            # Get Health
            ws_ad.cell(row=act_row, column=8, value=node['health']).style = 'style_normal'
            act_row = act_row + 1
    else:
        ws_ad.cell(row=act_row, column=2, value='AD Status').style = 'style_title'
        ws_ad.cell(row=act_row, column=3, value=ad['status']).style = 'style_normal'

# Generate Quota Sheet


def gen_quota_sheet(rc, xls_wb):
    # Get all quotas
    all_quotas = rc.quota.get_all_quotas_with_status()

    # Create and format 5th sheet for Quotas
    ws_quota = xls_wb.create_sheet(title="Quotas")
    ws_quota.column_dimensions['A'].width = 1
    ws_quota.row_dimensions[1].height = 5
    ws_quota.column_dimensions['B'].width = 15
    ws_quota.column_dimensions['C'].width = 40
    ws_quota.column_dimensions['D'].width = 20
    ws_quota.column_dimensions['E'].width = 20
    act_row = 2

    # API sends a list with quotas and paging, we don't use paging item
    for quotas in all_quotas:
        # Just get info from quotas, not paging
        if quotas['quotas']:
            print('ok')
            ws_quota.cell(row=act_row, column=2, value='Quota ID').style = 'style_title'
            ws_quota.cell(row=act_row, column=3, value='Path').style = 'style_title'
            ws_quota.cell(row=act_row, column=4, value='Limit (GB)').style = 'style_title'
            ws_quota.cell(row=act_row, column=5, value='Used (GB)').style = 'style_title'
            act_row = act_row + 1

            # Get information about all quotas and fill table
            for quota in quotas['quotas']:
                ws_quota.cell(row=act_row, column=2, value=quota['id']).style = 'style_normal'
                ws_quota.cell(row=act_row, column=3, value=quota['path']).style = 'style_normal'

                # Convert Limit from bytes to gigabytes --> / 1 000 000 000
                limit_in_gb = int(quota['limit']) / 1000000000
                ws_quota.cell(row=act_row, column=4, value=limit_in_gb).style = 'style_normal'

                # Convert capacity usage from bytes to gigabytes --> / 1 000 000 000
                used_in_gb = int(quota['capacity_usage']) / 1000000000
                ws_quota.cell(row=act_row, column=5, value=used_in_gb).style = 'style_normal'

                act_row = act_row + 1
        else:
            ws_quota.cell(row=act_row, column=2, value='Quotas').style = 'style_title'
            ws_quota.cell(row=act_row, column=3, value='No quotas defined on this cluster').style = 'style_normal'

# Generate SMB Sheet


def gen_smb_sheet(rc, xls_wb):
    # Get all smb shares
    shares = rc.smb.smb_list_shares()

    # Create and format 6th sheet for SMB Shares
    ws_smb = xls_wb.create_sheet(title="SMB Shares")
    ws_smb.column_dimensions['A'].width = 1
    ws_smb.row_dimensions[1].height = 5
    ws_smb.column_dimensions['B'].width = 30
    ws_smb.column_dimensions['C'].width = 50
    ws_smb.column_dimensions['D'].width = 35
    act_row = 2
    # Generate a table for each share
    for share in shares:
        # Create title for each table containing share name in merged cells
        ws_smb.merge_cells(start_row=act_row, start_column=2, end_row=act_row, end_column=4)
        ws_smb.cell(row=act_row, column=2, value=f"Share : {share['share_name']}").style = 'style_title'
        ws_smb.cell(row=act_row, column=3).style = 'style_title'
        ws_smb.cell(row=act_row, column=4).style = 'style_title'
        act_row = act_row + 1

        # Get Path
        ws_smb.cell(row=act_row, column=2, value='Path').style = 'style_title'
        ws_smb.merge_cells(start_row=act_row, start_column=3, end_row=act_row, end_column=4)
        ws_smb.cell(row=act_row, column=3, value=share['fs_path']).style = 'style_normal'
        ws_smb.cell(row=act_row, column=4).style = 'style_normal'
        act_row = act_row + 1

        # Get Description
        ws_smb.cell(row=act_row, column=2, value='Description').style = 'style_title'
        ws_smb.merge_cells(start_row=act_row, start_column=3, end_row=act_row, end_column=4)
        ws_smb.cell(row=act_row, column=3, value=share['description']).style = 'style_normal'
        ws_smb.cell(row=act_row, column=4).style = 'style_normal'
        act_row = act_row + 1

        # Get Access Based Enumeration
        ws_smb.cell(row=act_row, column=2, value='ABE enabled').style = 'style_title'
        ws_smb.merge_cells(start_row=act_row, start_column=3, end_row=act_row, end_column=4)
        ws_smb.cell(row=act_row, column=3, value=share['access_based_enumeration_enabled']).style = 'style_normal'
        ws_smb.cell(row=act_row, column=4).style = 'style_normal'
        act_row = act_row + 1

        # Get Default File Creation Mode
        ws_smb.cell(row=act_row, column=2, value='Default File Creation Mode').style = 'style_title'
        ws_smb.merge_cells(start_row=act_row, start_column=3, end_row=act_row, end_column=4)
        ws_smb.cell(row=act_row, column=3, value=share['default_file_create_mode']).style = 'style_normal'
        ws_smb.cell(row=act_row, column=4).style = 'style_normal'
        act_row = act_row + 1

        # Get Default Directory Creation Mode
        ws_smb.cell(row=act_row, column=2, value='Default Directory Creation Mode').style = 'style_title'
        ws_smb.merge_cells(start_row=act_row, start_column=3, end_row=act_row, end_column=4)
        ws_smb.cell(row=act_row, column=3, value=share['default_directory_create_mode']).style = 'style_normal'
        ws_smb.cell(row=act_row, column=4).style = 'style_normal'
        act_row = act_row + 1

        # Get Encryption settings
        ws_smb.cell(row=act_row, column=2, value='Requires Encryption').style = 'style_title'
        ws_smb.merge_cells(start_row=act_row, start_column=3, end_row=act_row, end_column=4)
        ws_smb.cell(row=act_row, column=3, value=share['require_encryption']).style = 'style_normal'
        ws_smb.cell(row=act_row, column=4).style = 'style_normal'
        act_row = act_row + 1

        # Get all permissions for trustees
        ws_smb.merge_cells(start_row=act_row, start_column=2, end_row=act_row, end_column=4)
        ws_smb.cell(row=act_row, column=2, value="Permissions").style = 'style_title'
        ws_smb.cell(row=act_row, column=3).style = 'style_title'
        ws_smb.cell(row=act_row, column=4).style = 'style_title'
        act_row = act_row + 1
        ws_smb.cell(row=act_row, column=2, value="Type").style = 'style_title'
        ws_smb.cell(row=act_row, column=3, value="Trustee").style = 'style_title'
        ws_smb.cell(row=act_row, column=4, value="Rights").style = 'style_title'
        act_row = act_row + 1

        for perm in share['permissions']:
            # Get trustee name from its sid
            trustee = ''

            # First check if sid refers to wellknown Everyone (S-1-1-0)
            if perm['trustee']['sid'] == "S-1-1-0":
                trustee = "Everyone"
            # If trustee is not everyone, check if it is a local trustee
            # Check if trustee is a local user
            if trustee == '' and perm['trustee']['domain'] == "LOCAL":
                all_users = rc.users.list_users()
                for user in all_users:
                    if user['sid'] == perm['trustee']['sid']:
                        trustee = perm['trustee']['domain'] + "\\" + user['name']

            # If trustee is not a local user check if it is a local group
            if trustee == '' and perm['trustee']['domain'] == "LOCAL":
                all_groups = rc.groups.list_groups()
                for group in all_groups:
                    if group['sid'] == perm['trustee']['sid']:
                        trustee = perm['trustee']['domain'] + "\\" + group['name']

            # If trustee is not Everyone or a local user, get SID information from Active Directory
            if trustee == '' and perm['trustee']['domain'] == "ACTIVE_DIRECTORY":
                ids = rc.ad.sid_to_ad_account(perm['trustee']['sid'])
                trustee = perm['trustee']['domain'] + "\\" + ids['name']

            # If trustee has been found, use its name, else use auth_id as trustee
            if trustee != '':
                # Convert Rights from list to string
                rights = ''
                for right in perm['rights']:
                    rights = rights + right + ', '
                rights = rights.rstrip(", ")
                ws_smb.cell(row=act_row, column=2, value=perm['type']).style = 'style_normal'
                ws_smb.cell(row=act_row, column=3, value=trustee).style = 'style_normal'
                ws_smb.cell(row=act_row, column=4, value=rights).style = 'style_normal'
                act_row = act_row + 1
            else:
                # Convert Rights from list to string
                rights = ''
                for right in perm['rights']:
                    rights = rights + right + ', '
                rights = rights.rstrip(", ")
                ws_smb.cell(row=act_row, column=2, value=perm['type']).style = 'style_normal'
                ws_smb.cell(row=act_row, column=3, value=perm['trustee']['auth_id']).style = 'style_normal'
                ws_smb.cell(row=act_row, column=4, value=rights).style = 'style_normal'
                act_row = act_row + 1

        # Get all network permissions if exists

        ws_smb.merge_cells(start_row=act_row, start_column=2, end_row=act_row, end_column=4)
        ws_smb.cell(row=act_row, column=2, value="Network permissions").style = 'style_title'
        ws_smb.cell(row=act_row, column=3).style = 'style_title'
        ws_smb.cell(row=act_row, column=4).style = 'style_title'
        act_row = act_row + 1
        ws_smb.cell(row=act_row, column=2, value="Type").style = 'style_title'
        ws_smb.cell(row=act_row, column=3, value="IP Ranges").style = 'style_title'
        ws_smb.cell(row=act_row, column=4, value="Rights").style = 'style_title'
        act_row = act_row + 1

        for net_perm in share['network_permissions']:
            # Format address_ranges to string with CR/LF between ranges
            ranges = ''
            for iprange in net_perm['address_ranges']:
                ranges = ranges + iprange + "\n"
            ranges = ranges.rstrip("\n")

            # Convert Rights from list to string
            rights = ''
            for right in net_perm['rights']:
                rights = rights + right + ', '
            rights = rights.rstrip(", ")
            # Fill table with network permissions
            if ranges:
                ws_smb.cell(row=act_row, column=2, value=net_perm['type']).style = 'style_normal'
                ws_smb.cell(row=act_row, column=3, value=ranges).style = 'style_normal'
                ws_smb.cell(row=act_row, column=4, value=rights).style = 'style_normal'
                act_row = act_row + 1

        # Create a thin separation line between tables
        ws_smb.merge_cells(start_row=act_row, start_column=2, end_row=act_row, end_column=4)
        ws_smb.row_dimensions[act_row].height = 20
        act_row = act_row + 1

# Generate NFS Sheet


def gen_nfs_sheet(rc, xls_wb):
    # Get all NFS exports
    exports = rc.nfs.nfs_list_exports()

    # Create and format 7th sheet for NFS Exports
    ws_nfs = xls_wb.create_sheet(title="NFS Exports")
    ws_nfs.column_dimensions['A'].width = 1
    ws_nfs.row_dimensions[1].height = 5
    ws_nfs.column_dimensions['B'].width = 30
    ws_nfs.column_dimensions['C'].width = 10
    ws_nfs.column_dimensions['D'].width = 50
    act_row = 2

    # Generate a table for each export
    for export in exports:

        # Generate Table Title with export name (export_path)
        ws_nfs.merge_cells(start_row=act_row, start_column=2, end_row=act_row, end_column=4)
        ws_nfs.cell(row=act_row, column=2, value=f"Export : {export['export_path']}").style = 'style_title'
        ws_nfs.cell(row=act_row, column=3).style = 'style_title'
        ws_nfs.cell(row=act_row, column=4).style = 'style_title'
        act_row = act_row + 1

        # Get Path
        ws_nfs.cell(row=act_row, column=2, value='Path').style = 'style_title'
        ws_nfs.merge_cells(start_row=act_row, start_column=3, end_row=act_row, end_column=4)
        ws_nfs.cell(row=act_row, column=3, value=export['fs_path']).style = 'style_normal'
        ws_nfs.cell(row=act_row, column=4).style = 'style_normal'
        act_row = act_row + 1

        # Get Description
        ws_nfs.cell(row=act_row, column=2, value='Description').style = 'style_title'
        ws_nfs.merge_cells(start_row=act_row, start_column=3, end_row=act_row, end_column=4)
        ws_nfs.cell(row=act_row, column=3, value=export['description']).style = 'style_normal'
        ws_nfs.cell(row=act_row, column=4).style = 'style_normal'
        act_row = act_row + 1

        # Get Hosts restrictions
        ws_nfs.merge_cells(start_row=act_row, start_column=2, end_row=act_row, end_column=4)
        ws_nfs.cell(row=act_row, column=2, value="Hosts Restrictions").style = 'style_title'
        ws_nfs.cell(row=act_row, column=3).style = 'style_title'
        ws_nfs.cell(row=act_row, column=4).style = 'style_title'
        act_row = act_row + 1

        ws_nfs.cell(row=act_row, column=2, value="Allowed IPs").style = 'style_title'
        ws_nfs.cell(row=act_row, column=3, value="Read-only").style = 'style_title'
        ws_nfs.cell(row=act_row, column=4, value="User mapping").style = 'style_title'
        act_row = act_row + 1

        for restric in export['restrictions']:
            # Format IP ranges / network to fit Excel cell (1 range per line)
            allowed_ips = ''
            for host_restric in restric['host_restrictions']:
                allowed_ips = allowed_ips + host_restric + "\n"
            allowed_ips = allowed_ips.rstrip("\n")
            # Replace empty range by * in Excel cell
            if allowed_ips == '':
                allowed_ips = "*"
            ws_nfs.cell(row=act_row, column=2, value=allowed_ips).style = 'style_normal'

            # Get Read-Only parameter
            ws_nfs.cell(row=act_row, column=3, value=restric['read_only']).style = 'style_normal'

            # Get User mapping

            # Case with no mapping
            if restric['user_mapping'] == "NFS_MAP_NONE":
                ws_nfs.cell(row=act_row, column=4, value="no mapping").style = 'style_normal'

            # Case with mapping all users
            if restric['user_mapping'] == "NFS_MAP_ALL":
                map_rule = "Map all users to " + restric['map_to_user']['id_value']
                ws_nfs.cell(row=act_row, column=4, value=map_rule).style = 'style_normal'

            # Case with mapping root user
            if restric['user_mapping'] == "NFS_MAP_ROOT":
                map_rule = "Map root user to " + restric['map_to_user']['id_value']
                ws_nfs.cell(row=act_row, column=4, value=map_rule).style = 'style_normal'
            act_row = act_row + 1

            # Create a thin separation line between tables
            ws_nfs.merge_cells(start_row=act_row, start_column=2, end_row=act_row, end_column=4)
            ws_nfs.row_dimensions[act_row].height = 15
            act_row = act_row + 1

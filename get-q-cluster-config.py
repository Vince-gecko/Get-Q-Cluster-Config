'''
Get Qumulo Cluster Config

Prompt for cluster IP, username and password

Will use Qumulo API to retrieve information about cluster config such as :
 - Hardware / disks used
 - Network configuration
 - Active Directory config
 - Quotas
 - SMB Shares
 ...
Then generate an Excel spreadsheet with all these information
'''

__Author__ = "Vincent Lamy"
__version__ = "2021.06.01"

# import statements
from qumulo.rest_client import RestClient
import openpyxl
from generate_sheets import *
from xl_styles import *
import getpass
import sys

# Variables

# Prompt for Qumulo Cluster Access
api_hostname = input("Qumulo Cluster IP : ")
api_user = input("Username : ")
api_password = getpass.getpass()


# Excel file name19
wb_file = 'qumulo-config.xlsx'

# login to qumulo API

rc = RestClient(api_hostname, 8000)
rc.login(api_user, api_password)

# Create Excel spreadsheet and add table Style defined in xl_styles.py
xls_wb = openpyxl.Workbook()
xls_wb.add_named_style(style_title)
xls_wb.add_named_style(style_normal)


# Generate Excel sheets
gen_nodes_sheet(rc, xls_wb)
gen_disks_sheet(rc, xls_wb)
gen_net_sheet(rc, xls_wb)
gen_ad_sheet(rc, xls_wb)
gen_quota_sheet(rc, xls_wb)
gen_smb_sheet(rc, xls_wb)
gen_nfs_sheet(rc, xls_wb)
xls_wb.save(wb_file)
rc.close()

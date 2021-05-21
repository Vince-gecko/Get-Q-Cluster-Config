# Get-Q-Cluster-Config

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

# Prerequisite
 - Python 3.8
 - Qumulo_api
 - openpyxl
 - Qumulo API version 2

# Usage
python get-q-cluster-config.py

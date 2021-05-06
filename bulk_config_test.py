from create_excel import API_Sheet
from bulk_config import BulkConfig
from datetime import datetime

# excel=API_Sheet()
# excelName=excel.create_excel(DHCP={[10]*100})

tgnObj = BulkConfig('127.0.0.1', clearConfig=True)
startTime = datetime.now()
# tgnObj.bulk_config(excelName)
tgnObj.bulk_config('Generate_IxNetwork_Config-original-mohan.xlsx')
tgnObj.create_trafficitems('Generate_IxNetwork_Config-original-mohan.xlsx')
endTime = datetime.now()
print("Start Time:", startTime)
print("End Time:", endTime.now())
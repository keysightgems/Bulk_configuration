from create_excel import API_Sheet
from bulk_config import BulkConfig

excel=API_Sheet()
excelName=excel.create_excel()

tgnObj = BulkConfig('127.0.0.1', clearConfig=True)
tgnObj.bulk_config(excelName)
tgnObj.create_trafficitems(excelName)

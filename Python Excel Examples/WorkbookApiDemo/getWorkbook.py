import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.workbook_api.WorkbookApi(configuration)

name = "GetWorkbook.xlsx"
storage = ""
password = ""
folder = "ExcelDocument"
result=api.get_workbook(name, password=password, storage=storage, folder=folder)
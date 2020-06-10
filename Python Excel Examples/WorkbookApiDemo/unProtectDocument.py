import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.workbook_api.WorkbookApi(configuration)

name = "UnProtectDocument.xlsx"
storage = ""
folder = "ExcelDocument"
protection = spirecloudexcel.models.WorkbookProtectionRequest()
protection.password = "123"
protection.protection_type = "READONLY"  # Available type: ALL, READONLY, STRUCTURE, WINDOWS
api.un_protect_document(name, protection, folder=folder, storage=storage)
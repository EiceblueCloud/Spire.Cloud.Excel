import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "ProtectWorksheet.xlsx"
storage = ""
folder = "ExcelDocument"
sheetName = "Sheet1"
protectParameter = spirecloudexcel.models.ProtectSheetParameter()
protectParameter.password = "123"
protectParameter.allow_inserting_row = "false"
api.unprotect_worksheet(name, sheetName, protectParameter, storage=storage, folder=folder)
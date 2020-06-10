import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "RenameWorksheet.xlsx"
folder = "ExcelDocument";
storage = ""
sheetName = "Sheet1"
newName = "newSheetName"
api.rename_worksheet(name, sheetName, newName, folder=folder, storage=storage)
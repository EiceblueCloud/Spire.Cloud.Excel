import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "DeleteWorksheet.xlsx"
folder = "ExcelDocument";
storage = "";
sheetName = "Sheet1";
api.delete_worksheet(name, sheetName, folder=folder, storage=storage)
import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "ChangeVisibilityWorksheet.xlsx"
storage = ""
folder = "ExcelDocument"
sheetName = "Sheet1"
isVisble = 'true'
api.change_visibility_worksheet(name, sheetName, is_visible=isVisble, folder=folder, storage=storage)
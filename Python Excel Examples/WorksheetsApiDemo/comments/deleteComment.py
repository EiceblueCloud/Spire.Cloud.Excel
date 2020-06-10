import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "DeleteComment.xlsx"
storage = ""
folder = "ExcelDocument"
sheet_name = "Sheet1"
cellName = "B5"
api.delete_comment(name, sheet_name, cellName, folder=folder, storage=storage)
import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "GetComment.xlsx"
storage = ""
folder = "ExcelDocument"
sheet_name = "Sheet1"
cellName = "E17"
result = api.get_comment(name, sheet_name, cellName, storage=storage, folder=folder)
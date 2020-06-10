import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "GetComments.xlsx"
storage = ""
folder = "ExcelDocument"
sheet_name = "Sheet1"
result = api.get_comments(name, sheet_name, storage=storage, folder=folder)
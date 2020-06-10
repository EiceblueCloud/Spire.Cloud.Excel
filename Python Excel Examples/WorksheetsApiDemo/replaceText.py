import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "ReplaceText.xlsx"
folder = "ExcelDocument";
storage = ""
sheetName = "Sheet1"
oldValue = "North America"
newValue = "北美"
matchCase = 'true'
api.replace_text(name, sheetName, oldValue,newValue, match_case = matchCase,folder=folder, storage=storage)
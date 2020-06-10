import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "SearchText.xlsx"
folder = "ExcelDocument";
storage = ""
sheetName = "Sheet1"
text = "America"
matchCase = "true"  # Default value is false
result = api.search_text(name, sheetName, text, match_case=matchCase, folder=folder, storage=storage)
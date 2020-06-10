import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "GetCalculateFormula_COUNT.xlsx"
folder = "ExcelDocument"
storage = ""
sheetName = "Sheet1";
formula = "=COUNT(D2:D4,E5:E8)";
result = api.calculate_formula(name, sheetName, formula, folder=folder, storage=storage)
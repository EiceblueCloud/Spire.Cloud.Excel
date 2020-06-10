import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "GetCalculateFormula_SUM.xlsx"
folder = "ExcelDocument"
storage = ""
sheetName = "Sheet1";
formula = "=SUM(D8:D9)";
result = api.calculate_formula(name, sheetName, formula, folder=folder, storage=storage)
import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "GetCalculateFormula_IF.xlsx"
folder = "ExcelDocument"
storage = ""
sheetName = "Sheet1";
formula ="=IF(D3<D9,true,false)";
result = api.calculate_formula(name, sheetName, formula, folder=folder, storage=storage)
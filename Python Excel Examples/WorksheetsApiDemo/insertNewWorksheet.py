import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "InsertNewWorksheet.xlsx"
folder = "ExcelDocument";
storage = ""
sheettype = "NormalWorksheet"  # Can be NormalWorksheet, ChartSheet
index = 2
newsheetname = "SheetInsert"
api.insert_new_worksheet(name, newsheetname, index, sheettype, folder=folder, storage=storage)
import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "AutofitRows.xlsx"
password = ""
storage = ""
folder = "ExcelDocument"
sheet_name = "Sheet1"
start_Row = 17
end_Row = 23
only_auto = "true"
api.autofit_rows(name, sheet_name, start_Row, end_Row, only_auto=only_auto, folder=folder, storage=storage)
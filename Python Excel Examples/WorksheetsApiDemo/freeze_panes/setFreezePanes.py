import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "SetFreezePanes.xlsx"
storage = ""
folder = "ExcelDocument"
sheet_name = "Sheet1"
freezedRows = 3
freezedColumns = 3
api.set_freeze_panes(name, sheet_name, freezedRows, freezedColumns, storage=storage, folder=folder)
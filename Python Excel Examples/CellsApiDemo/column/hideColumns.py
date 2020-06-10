import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "HideColumns_1.xlsx"
sheetName = "Sheet4"
start_column = 2
total_columns = 4
storage = ""
folder = "/Cells/Column/"
api.hide_columns(name, sheet_name=sheetName, start_column=start_column,
                      total_columns=total_columns,
                      folder=folder,
                      storage=storage)
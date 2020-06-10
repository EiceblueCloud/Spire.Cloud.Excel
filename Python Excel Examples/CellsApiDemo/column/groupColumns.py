import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "GroupColumns_1.xlsx"
sheetName = "Sheet4"
first_index = 1
last_index = 4
hide = "true"
storage = ""
folder = "/Cells/Column/"
api.group_columns(name, sheet_name=sheetName, first_index=first_index,
                       last_index=last_index, hide=hide,
                       folder=folder,
                       storage=storage)

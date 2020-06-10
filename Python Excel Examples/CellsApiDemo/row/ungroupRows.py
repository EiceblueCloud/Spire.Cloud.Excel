import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "UngroupRows_1.xlsx"
sheetName = "Sheet4"
first_index = 5
last_index = 7
is_all = "false"
storage = ""
folder = "/Cells/Row/"
api.ungroup_rows(name, sheet_name=sheetName, first_index=first_index, last_index=last_index, is_all=is_all,folder=folder, storage=storage)
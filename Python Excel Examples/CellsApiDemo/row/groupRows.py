import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "GroupRows_1.xlsx"
sheetName = "Sheet4"
first_index = 2
last_index = 6
hide = "true"
storage = ""
folder = "/Cells/Row/"
api.group_rows(name, sheet_name=sheetName, first_index=first_index, last_index=last_index, hide=hide,
                    folder=folder,
                    storage=storage)

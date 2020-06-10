import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "SetRowHeight_1.xlsx"
sheetName = "Sheet1"
row_index = 3
height = 26
storage = ""
folder = "/Cells/Row/"
api.set_row_height(name, sheet_name=sheetName, row_index=row_index, height=height,folder=folder,storage=storage)
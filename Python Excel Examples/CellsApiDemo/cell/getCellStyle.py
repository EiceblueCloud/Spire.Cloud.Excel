import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "GetCellStyle_1.xlsx"
folder = "ExcelDocument"
storage = ""
sheetName = "Sheet4"
cellName = "A4"
result = api.get_cell_style(name, sheetName, cellName, folder=folder, storage=storage)
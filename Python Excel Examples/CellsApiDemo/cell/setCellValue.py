import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)

name = "SetCellValue_1.xlsx"
folder = "/Cells/Cell/Value/"
storage = ""
sheetName = "Sheet1"
cellName = "C20"
value = "1"
type = "string"
formula = ""
api.set_cell_value(name, sheetName, cellName, value, type=type, formula=formula, folder=folder, storage=storage)
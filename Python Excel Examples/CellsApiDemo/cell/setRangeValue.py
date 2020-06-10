import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "SetRangeValue_1.xlsx"
folder = "/Cells/Cell/Value/"
storage = ""
sheetName = "Sheet1"
cellArea = "A1:C3"
value = "250.5"
type = "double"
api.set_range_value(name, sheetName, cellArea, value, type, folder=folder, storage=storage)
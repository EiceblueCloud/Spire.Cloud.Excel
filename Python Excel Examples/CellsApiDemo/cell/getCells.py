import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)

name = "GetCells_1.xlsx"
folder = "ExcelDocument"
storage = ""
sheetName = "Sheet1"
offset = 0
count = 0
result = api.get_cells(name, sheetName, offest=offset, count=count, folder=folder, storage=storage)

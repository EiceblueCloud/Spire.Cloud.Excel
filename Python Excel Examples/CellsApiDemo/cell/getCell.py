import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "GetCell_1.xlsx"
sheetName = "Sheet1"
cellOrMethodName = "A2"
folder = "ExcelDocument"
storage = ""
result = api.get_cell(name, sheetName, cellOrMethodName, folder=folder, storage=storage)

import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "GetRows_1.xlsx"
sheetName = "Sheet1"
storage = ""
folder = "/ExcelDocument/"
result = api.get_rows(name, sheet_name=sheetName, folder=folder, storage=storage)

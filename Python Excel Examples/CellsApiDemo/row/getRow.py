import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "GetRow_1.xlsx"
sheetName = "Sheet4"
row_index = 2
storage = ""
folder = "/ExcelDocument/"

result = api.get_row(name, sheet_name=sheetName, row_index=row_index, folder=folder, storage=storage)
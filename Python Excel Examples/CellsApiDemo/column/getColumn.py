import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "GetColumn_1.xlsx"
sheetName = "Sheet4"
column_index = 2  # index starts 1
storage = ""
folder = "/ExcelDocument/"
result = api.get_column(name, sheet_name=sheetName, column_index=column_index, folder=folder, storage=storage)
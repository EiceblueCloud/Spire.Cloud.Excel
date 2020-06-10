import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "DeleteRow.xlsx"
sheetName = "Sheet2"
row_index = 1
storage = ""
folder = "/Cells/Row/"
api.delete_row(name, sheet_name=sheetName, row_index=row_index, folder=folder, storage=storage)


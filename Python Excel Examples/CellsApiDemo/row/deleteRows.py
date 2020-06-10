import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "DeleteRows_1.xlsx"
sheetName = "Sheet4"
startRow = 2
total_rows = 2
storage = ""
folder = "/Cells/Row/"
api.delete_rows(name, sheet_name=sheetName, startrow=startRow, total_rows=total_rows,folder=folder, storage=storage)
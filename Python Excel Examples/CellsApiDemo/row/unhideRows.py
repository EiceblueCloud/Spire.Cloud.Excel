import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "UnhideRows_1.xlsx"
sheetName = "Sheet5"
startRow = 15
total_rows = 4
rowHeight = 30 #when not setting row height, it will use the previous row height
storage = ""
folder = "/Cells/Row/"
api.unhide_rows(name, sheet_name=sheetName, startrow=startRow, total_rows=total_rows, height=rowHeight, folder=folder, storage=storage)

import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "MergeCells_1.xlsx"
folder = "/Cells/Cell/Merge/"
storage = ""
sheetName = "Sheet4"
startRow = 2
startColumn = 3
totalRows = 4
totalColumns = 4
api.merge_cells(name, sheetName, startRow, startColumn, totalRows, totalColumns, folder=folder, storage=storage)
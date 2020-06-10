import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "UnMergeCells_1.xlsx"
folder = "/Cells/Cell/Merge/"
storage = ""
sheetName = "Sheet2"
startRow = 3
startColumn = 1
totalRows = 1
totalColumns = 2
api.unmerge_cells(name, sheetName, startRow, startColumn, totalRows, totalColumns, folder=folder, storage=storage)

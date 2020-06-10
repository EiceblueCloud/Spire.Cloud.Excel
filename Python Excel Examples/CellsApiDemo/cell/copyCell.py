import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl= "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)

name = "test.xlsx"
storage = ""
folder = "/input/"
worksheet = "Sheet1"  # Source
sheetName = "Sheet2"  # Destination
cellName = "A1:A10"  # Source
destCellName = "F20:F30"
row = 1  # index starts at 1
column = 1  # index starts at 1
#When cellName has value, the value of cellName will be copied to destCellName,
#When cellName = "" , the value of the cell that row and column point to will be copied to destCellName
api.copy_cell(name, destCellName, sheetName, worksheet, source_cell_name=cellName, source_row=row, source_column=column,
                   folder=folder, storage=storage)
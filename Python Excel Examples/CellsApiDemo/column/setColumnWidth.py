import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "SetColumnWidth_1.xlsx"
sheetName = "Sheet1"
column_index = 2
width = 40
storage = ""
folder = "/Cells/Column/"
api.set_column_width(name, sheet_name=sheetName, column_index=column_index,
                          width=width,
                          folder=folder,
                          storage=storage)
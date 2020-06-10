import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "UnhideColumns_1.xlsx"
sheetName = "Sheet5"
start_column = 2
total_columns = 1
column_width = 30.5 #when not setting column width, it will use the previous column width
storage = ""
folder = "/Cells/Column/"
api.unhide_columns(name, sheet_name=sheetName, startcolumn=start_column,
                   total_columns=total_columns, width=column_width,
                   folder=folder,
                   storage=storage)
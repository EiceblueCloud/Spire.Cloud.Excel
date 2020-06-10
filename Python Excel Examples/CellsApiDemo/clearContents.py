import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "ClearContents_1.xlsx"
folder = "/Cells/Clear/"
storage = ""
sheetName = "Sheet4"
range = "A1:G19"
start_row = None
start_column = None
end_row = None
end_column = None

#range = None
#start_row = 13
#start_column = 1
#end_row = 13
#end_column = 1

api.clear_contents(name, sheetName, range=range, start_row=start_row, start_column=start_column, end_row=end_row, end_column=end_column,
                        folder=folder, storage=storage)

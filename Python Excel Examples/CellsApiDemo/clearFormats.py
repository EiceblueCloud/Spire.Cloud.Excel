import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "ClearFormats_1.xlsx"
folder = "/Cells/Clear/"
storage = ""
sheetName = "Sheet4"
range = "A1:G19" #range=None
start_row = None  #4
start_column = None #1
end_row = None #13
end_column = None #20
api.clear_formats(name, sheetName, range=range, start_row=start_row, start_column=start_column,
                       end_row=end_row, end_column=end_column, folder=folder, storage=storage)


import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "DeleteColumns_1.xlsx"
sheetName = "Sheet4"
columnIndex = 1
columns = 2
storage = ""
folder = "/Cells/Column/"
api.delete_columns(name, sheet_name=sheetName, columnIndex=columnIndex,
                        columns=columns,
                        folder=folder,
                        storage=storage)
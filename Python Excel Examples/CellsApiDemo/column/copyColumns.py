import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "CopyColumns_1.xlsx"
sheetName = "Sheet4"
source_column_index = 2
destination_column_index = 3
column_number = 2
dest_sheet = "Sheet2"
storage = ""
folder = "/Cells/Column/"
api.copy_columns(name, sheet_name=sheetName, source_column_index=source_column_index,
                      destination_column_index=destination_column_index, column_number=column_number,dest_sheet=dest_sheet,
                      folder=folder,
                      storage=storage)
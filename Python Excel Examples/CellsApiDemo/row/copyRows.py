import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "CopyRows_1.xlsx"
sheetName = "Sheet4"
source_row_index = 1
destination_row_index = 5
row_number = 3
dest_sheet = "Sheet2"
password = ""
storage = ""
folder = "/Cells/Row/"
api.copy_rows(name, sheet_name=sheetName, source_row_index=source_row_index,
                   destination_row_index=destination_row_index, row_number=row_number, dest_sheet=dest_sheet,
                   folder=folder,
                   storage=storage)

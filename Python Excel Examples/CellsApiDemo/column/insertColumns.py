import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "InsertColumns_1.xlsx"
sheetName = "Sheet1"
column_index = 2  #index starts 1
columns = 2
format_as = "FormatAsBefore"  # FormatAsBefore FormatAsAfter FormatDefault.(default to FormatDefault)
storage = ""
folder = "/Cells/Column/"
api.insert_columns(name, sheet_name=sheetName, column_index=column_index,
                        columns=columns,format_as=format_as,
                        folder=folder,
                        storage=storage)
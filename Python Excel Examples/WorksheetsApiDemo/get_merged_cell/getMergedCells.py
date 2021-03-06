import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "GetMergedCells.xlsx"
storage = ""
folder = "ExcelDocument"
sheet_name = "Sheet1"
mergedCellIndex = 0
result = api.get_merged_cells(name, sheet_name, folder=folder, storage=storage)
import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.ranges_api.RangesApi(configuration)

name = "MoveRange.xlsx"
folder = "Ranges"
storage = ""
sheetName = "Sheet1"
Range = spirecloudexcel.models.Range()
Range.first_row = 1
Range.first_column = 1
Range.row_count = 1
Range.column_count = 1
destRow = 25  # Start value is 0
destColumn = 3  # Start value is 0
api.move_range(name, sheetName, Range, destRow, destColumn, storage = storage, folder = folder)
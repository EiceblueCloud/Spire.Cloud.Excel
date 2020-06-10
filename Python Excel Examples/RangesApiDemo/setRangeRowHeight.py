import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.ranges_api.RangesApi(configuration)

name = "SetRangeRowHeight.xlsx"
folder = "Ranges"
storage = ""
sheetName = "Sheet1"
Range = spirecloudexcel.models.Range()
Range.first_row = 1
Range.first_column = 1
Range.row_count = 10
Range.column_count = 5
value = 20
api.set_range_row_height(name, sheetName, Range, value, storage = storage, folder = folder)
import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.ranges_api.RangesApi(configuration)

name = "SetRangeValue.xlsx"
folder = "Ranges"
storage = ""
sheetName = "Sheet1"
Range = spirecloudexcel.models.Range()
Range.first_row = 2
Range.first_column = 1
Range.row_count = 2
Range.column_count = 5
value = "3/4"
isConverted = "true"
setStyle = "true"
api.set_range_value(name, sheetName, Range, value, is_converted=isConverted, set_style=setStyle, folder=folder, storage=storage)
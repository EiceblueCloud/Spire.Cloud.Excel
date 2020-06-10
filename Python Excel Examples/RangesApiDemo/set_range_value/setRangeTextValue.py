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
range = spirecloudexcel.models.Range()
range.first_row = 1
range.first_column = 1
range.row_count = 2
range.column_count = 5
value = "Hello, E-iceblue"
isConverted = "true"
setStyle = "false"
api.set_range_value(name, sheetName, range, value, is_converted=isConverted, set_style=setStyle, folder=folder, storage=storage)
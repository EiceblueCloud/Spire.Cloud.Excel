import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.ranges_api.RangesApi(configuration)

name = "SetRangeOutlineBorder.xlsx"
folder = "Ranges"
storage = ""
sheetName = "Sheet1"
rangeOperate = spirecloudexcel.models.RangeSetOutlineBorderRequest()
rangeOperate.border_edge = "outline" # Available values: outline, vertical, horizontal, edgeleft
rangeOperate.border_style = "Medium" # Available values: Medium, DashDot, MediumDashDotDot, Thin
rangeOperate.border_color = spirecloudexcel.models.Color(255, 255, 0, 0)
Range = spirecloudexcel.models.Range()
Range.first_row = 1
Range.first_column = 1
Range.row_count = 5
Range.column_count = 10
rangeOperate.range=Range
api.set_range_outline_border(name, sheetName, rangeOperate, storage = storage, folder = folder)
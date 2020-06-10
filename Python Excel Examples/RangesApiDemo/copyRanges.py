import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.ranges_api.RangesApi(configuration)

name = "CopyRanges.xlsx"
folder = "Ranges"
storage = ""
password = ""
sheetName = "Sheet1"
rangeOperate = spirecloudexcel.models.RangeCopyRequest()
rangeOperate.operate = "copystyle"  # Available values: copystyle, copydata, copyvalue, copyto
PasteOptions=spirecloudexcel.models.PasteOptions()
PasteOptions.only_visible_cells=True
range1 = spirecloudexcel.models.Range()
range1.first_column = 1
range1.first_row = 1
range1.row_count = 10
range1.column_count = 10
range2 = spirecloudexcel.models.Range()
range2.first_column = 1
range2.first_row = 30
range2.row_count = 5
range2.column_count = 8
rangeOperate.source = range1
rangeOperate.target = range2
api.copy_ranges(name, sheetName, rangeOperate, storage = storage, password=password,folder = folder)
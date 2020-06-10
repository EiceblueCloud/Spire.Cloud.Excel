import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.ranges_api.RangesApi(configuration)

name = "UnMergeRange.xlsx"
folder = "Ranges"
storage = ""
sheetName = "Sheet1"
Range = spirecloudexcel.models.Range()
Range.first_row = 3
Range.first_column = 4
Range.row_count = 2
Range.column_count = 5
api.unmerge_range(name, sheetName, Range, storage=storage, folder=folder)
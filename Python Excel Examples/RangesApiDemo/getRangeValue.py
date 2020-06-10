import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.ranges_api.RangesApi(configuration)

name = "GetRangeValue.xlsx"
sheetName = "Sheet1"
folder = "ExcelDocument"
storage = ""
firstRow = 1
firstColumn = 6
rowCount = 2
columnCount = 2
result = api.get_range_value(name, sheetName, namerange="",first_row=firstRow,first_column=firstColumn,row_count=rowCount,column_count=columnCount,storage=storage, folder=folder)
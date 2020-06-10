import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "GetNamedRangeValue.xlsx"
storage = ""
folder = "ExcelDocument"
namerange = "Name1"
result=api.get_named_range_value(name, namerange, folder=folder, storage=storage)
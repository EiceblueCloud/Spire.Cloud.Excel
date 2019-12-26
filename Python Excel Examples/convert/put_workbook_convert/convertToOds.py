import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
configuration = ExcelConfiguration(appId, appKey)
api = spirecloudexcel.api.general_api.GeneralApi(configuration)

file = "D:/data/test.xls"
password = None
outputPath = "input/result.ods"
result = api.put_workbook_convert("Ods", outputPath, workbook=file, password=password)

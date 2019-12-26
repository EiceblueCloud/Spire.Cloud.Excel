import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
configuration = ExcelConfiguration(appId, appKey)
api = spirecloudexcel.api.general_api.GeneralApi(configuration)

file = "D:/data/test.xls"
password = None
outputPath = "output/result.xlsx"
result = api.put_workbook_convert("Xlsx", outputPath, workbook=file, password=password)
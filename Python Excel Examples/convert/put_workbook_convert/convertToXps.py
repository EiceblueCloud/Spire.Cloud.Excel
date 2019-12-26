import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
configuration = ExcelConfiguration(appId, appKey)
api = spirecloudexcel.api.general_api.GeneralApi(configuration)

file = "D:/data/test.xlsx"
password = None
outputPath = "output/result.xps"
result = api.put_workbook_convert("Xps", outputPath, workbook=file, password=password)
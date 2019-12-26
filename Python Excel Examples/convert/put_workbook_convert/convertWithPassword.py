import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
configuration = ExcelConfiguration(appId, appKey)
api = spirecloudexcel.api.general_api.GeneralApi(configuration)

file = "D:/data/test1.xlsx"
password = "test"
outputPath = "output/result.pdf"
result = api.put_workbook_convert("Pdf", outputPath, workbook=file, password=password)


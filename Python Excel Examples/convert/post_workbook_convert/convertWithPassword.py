import shutil

import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
configuration = ExcelConfiguration(appId, appKey)
api = spirecloudexcel.api.general_api.GeneralApi(configuration)

file = "D:/data/test1.xlsx"
password = "test"
result = api.post_workbook_convert("Pdf", workbook=file, password=password)
file = "D:/output/result.pdf"
shutil.copyfile(result, file)
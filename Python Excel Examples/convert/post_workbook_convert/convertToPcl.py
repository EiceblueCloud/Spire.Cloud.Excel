import shutil

import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
configuration = ExcelConfiguration(appId, appKey)
api = spirecloudexcel.api.general_api.GeneralApi(configuration)

file = "D:/data/test.xlsx"
password = None
result = api.post_workbook_convert("Pcl", workbook=file, password=password)
file = "D:/output/result.pcl"
shutil.copyfile(result, file)
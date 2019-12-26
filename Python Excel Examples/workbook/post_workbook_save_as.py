import shutil

import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration


appId = "your id"
appKey = "your key"
configuration = ExcelConfiguration(appId, appKey)
api = spirecloudexcel.api.workbook_api.WorkbookApi(configuration)

#input file format： ods, xls, xlsx, xlsb, xlsm, xltm, xltx
#format： Ods, Ps, Xls, Xlsx, Xps, Xlsb, Pdf
name = "test.xlsx"
password = ""
storage = ""
format = "Pdf" #The first letter must be capitalized
folder = "input/data"
result = api.post_workbook_save_as(name, format=format, password=password, storage=storage, folder=folder)
file = "D:/output/result.pdf"
shutil.copyfile(result, file)
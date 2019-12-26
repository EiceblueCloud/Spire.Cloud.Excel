import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
configuration = ExcelConfiguration(appId, appKey)
api = spirecloudexcel.api.workbook_api.WorkbookApi(configuration)

#input file format:  ods, xls, xlsx, xlsb, xlsm, xltm, xltx
#output file format:  ods, pcl, ps, xls, xlsx, xps, pdf, xlsm, xltm, xltx
name = "test.pdf"
inputPath = "D:/data/test.xltx"
password = ""
storage = ""
folder = "output/data"
result = api.post_workbook(name, data=inputPath, input_password="", password=password, storage=storage, folder=folder)
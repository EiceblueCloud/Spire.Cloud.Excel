import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
configuration = ExcelConfiguration(appId, appKey)
api = spirecloudexcel.api.workbook_api.WorkbookApi(configuration)

#input file format： ods, xls, xlsx, xlsb, pdf, xlsm, xltm, xltx
#format： Ods, Ps, Xls, Xlsx, Xps, Pdf
name = "test.xltx"
password = ""
storage = ""
format = "Pdf" #The first letter must be capitalized
inputFolder = "input/data"
outputPath = "output/out.pdf"
api.put_workbook_save_as(name, format=format, password=password, storage=storage, folder=inputFolder, out_path=outputPath)

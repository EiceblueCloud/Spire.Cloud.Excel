import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
configuration = ExcelConfiguration(appId, appKey)
api = spirecloudexcel.api.workbook_api.WorkbookApi(configuration)

#input file format:  ods, xls, xlsx, xlsb, xlsm, xltm, xltx
name = "test.xltx"
password = ""
storage = ""
inputFolder = "input/data"
result = api.get_workbook(name, password=password, storage=storage, folder=inputFolder)

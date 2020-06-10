import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.workbook_api.WorkbookApi(configuration)

# Available formats: xlsx, xls, xlsm, xltm, xltx, ods
name = "CreateWorkbook.xlsx"
password = ""
inputPassword = ""
storage = ""
data="D:/data/test.xlsx"
folder = "ExcelDocument"
api.create_workbook(name, data=data, input_password=inputPassword, password=password, storage=storage,folder=folder)
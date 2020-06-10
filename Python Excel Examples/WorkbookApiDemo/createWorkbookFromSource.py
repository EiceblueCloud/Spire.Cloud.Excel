import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.workbook_api.WorkbookApi(configuration)

name = "CreateWorkbookFromSource.xlsx"
source_path="ExcelDocument"
source_password=""
source_storage=""
password = ""
storage = ""
folder = "Workbook"
api.create_workbook_from_source(name, source_path=source_path, source_password=source_password, source_storage=source_storage, password=password, storage=storage, folder=folder)
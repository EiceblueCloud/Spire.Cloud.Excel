import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
configuration = ExcelConfiguration(appId, appKey)
api = spirecloudexcel.api.workbook_api.WorkbookApi(configuration)

#input file format:  ods, xls, xlsx, xlsm, xltm, xltx
name = "test.xltx"
source_path = "input/data"
source_password = ""
source_storage = ""
output_password = "123"
output_storage = ""
output_folder = "output/data"
result = api.post_workbook_from_source(name, source_path=source_path,
source_password=source_password, source_storage=source_storage, password=output_password, storage=output_storage,folder=output_folder)

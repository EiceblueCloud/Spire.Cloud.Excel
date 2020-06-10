import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.workbook_api.WorkbookApi(configuration)

name = "ConvertWorkbookToPath.xlsx"
password = ""
storage = ""
format = "Xlsx"  # Available formats: Xlsx, Xls, Xlsb, Ods, Xps, Ps, Pcl, Pdf
folder = "ExcelDocument"
dest_file_path = "Output/ConvertWorkbookToPath.xlsx"
api.convert_workbook_to_path(name, dest_file_path=dest_file_path, format=format,password=password, storage=storage, folder=folder)
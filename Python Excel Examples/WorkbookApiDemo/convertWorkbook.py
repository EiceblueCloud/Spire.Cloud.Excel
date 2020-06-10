import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.workbook_api.WorkbookApi(configuration)

name = "ConvertWorkbook.xlsx"
password = ""
storage = ""
format = "Xlsx"  # Available format: Xlsx, Xls, Xlsb, Ods, Xps, Ps, Pcl, Pdf
folder = "ExcelDocument"
result = api.convert_workbook(name, format=format,password=password, storage=storage, folder=folder)
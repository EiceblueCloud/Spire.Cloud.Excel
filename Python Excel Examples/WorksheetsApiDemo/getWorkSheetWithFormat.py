import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "GetWorkSheetWithFormat.xlsx"
sheetName = "Sheet1"
format = "html"  # Available format: html, pdf, png, jpeg, bmp, emf,tiff
verticalResolution = 0
horizontalResolution = 0
folder = "ExcelDocument"
storage = ""
result = api.get_work_sheet_with_format(name, sheetName, format, vertical_resolution=verticalResolution,horizontal_resolution=horizontalResolution, folder=folder,storage=storage)
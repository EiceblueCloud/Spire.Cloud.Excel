import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

destination_name = "CopyWorksheet_destination.xlsx"
sourceWorkbook = "CopyWorksheet_original.xlsx"
sourceFolder = "ExcelDocument"
folder = "Worksheets/Copy"
storage = ""
sourceSheet = "Sheet1"
destSheet = "Sheet2"
isVisble = 'true'
api.copy_worksheet(destination_name, destSheet, sourceSheet, source_workbook=sourceWorkbook,source_folder=sourceFolder,folder=folder, storage=storage)
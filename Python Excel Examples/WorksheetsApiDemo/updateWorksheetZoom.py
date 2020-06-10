import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.models import Color

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "UpdateWorksheetZoom.xlsx"
folder = "ExcelDocument";
storage = ""
sheetName = "Sheet1"
value = 150
api.update_worksheet_zoom(name, sheetName, value, folder=folder, storage=storage)
import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "MoveWorksheet.xlsx"
folder = "ExcelDocument";
storage = ""
sheetName = "Sheet1"
moving =   spirecloudexcel.models.WorksheetMovingRequest()
moving.destination_worksheet = "Sheet4"
moving.position="after" # Can be Before, after
api.move_worksheet(name, sheetName, moving, folder=folder, storage=storage)
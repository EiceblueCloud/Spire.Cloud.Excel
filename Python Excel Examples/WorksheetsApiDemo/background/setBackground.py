import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "SetBackground.xlsx"
storage = ""
folder = "ExcelDocument"
sheet_name = "Sheet1"
pic_path = "ExcelDocument/SpireXls.png"
api.set_background(name, sheet_name, pic_path=pic_path, folder=folder, storage=storage)
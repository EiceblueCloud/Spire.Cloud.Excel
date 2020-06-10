import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "UpdateComment.xlsx"
storage = ""
folder = "ExcelDocument"
sheet_name = "Sheet1"
cellName = "C6"
comment = spirecloudexcel.models.Comment()
comment.note = "Hello,Comments"
comment.width = 160
comment.height = 30
comment.is_visible = True
comment.text_horizontal_alignment = "center"
comment.auto_size = True
api.update_comment(name, sheet_name, cellName, comment, storage=storage, folder=folder)

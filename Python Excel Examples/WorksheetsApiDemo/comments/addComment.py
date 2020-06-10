import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "AddComment.xlsx"
storage = ""
folder = "ExcelDocument"
sheet_name = "Sheet1"
cellName = "C2:D4"
comment = spirecloudexcel.models.Comment()
comment.author = "jerry"
comment.note = "hello"
comment.text_vertical_alignment = "Bottom"
comment.auto_size = True
api.add_comment(name, sheet_name, cellName, comment, storage=storage, folder=folder)
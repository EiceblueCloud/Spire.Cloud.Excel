import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.models import Color

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.worksheets_api.WorksheetsApi(configuration)

name = "UpdateWorksheetProperty.xlsx"
folder = "ExcelDocument";
storage = ""
sheetName = "Sheet1"
sheet = spirecloudexcel.models.Worksheet()
sheet.zoom = 50;
sheet.index = 3
sheet.visibility_type = "hidden"
sheet.is_gridlines_visible = False
sheet.tab_color = Color(255, 255, 0, 0)  # 设置选项卡颜色
sheet.is_protected = True
sheet.is_selected = True
api.update_worksheet_property(name, sheetName, sheet, folder=folder, storage=storage)
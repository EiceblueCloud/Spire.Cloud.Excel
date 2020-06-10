import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.ranges_api.RangesApi(configuration)

name = "UpdateRangeStyle.xlsx"
folder = "Ranges"
sheetName = "Sheet1"
range = "C2:D6"
storage = ""
style = spirecloudexcel.models.Style()
font = spirecloudexcel.models.Font()
font.color = spirecloudexcel.models.Color(0, 160, 32, 240)
font.underline = "Single"
font.size = 13
font.is_italic = True
font.is_bold = True
font.name = "楷体"
style.font = font
borders = []
topBorder = spirecloudexcel.models.Border("Medium", spirecloudexcel.models.Color(255, 255, 0, 0), "EdgeTop")
rightBorder = spirecloudexcel.models.Border("DashDot", spirecloudexcel.models.Color(255, 0, 255, 0), "EdgeRight")
borders.append(topBorder)
borders.append(rightBorder)
style.border_collection = borders
style.indent_level = 2
style.horizontal_alignment = "Left"
api.update_range_style(name, sheet_name=sheetName, range=range, style=style, folder=folder,storage=storage)
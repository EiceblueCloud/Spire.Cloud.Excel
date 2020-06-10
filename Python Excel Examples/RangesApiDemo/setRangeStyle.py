import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration

appId = "your id"
appKey = "your key"
baseUrl="https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.ranges_api.RangesApi(configuration)

name = "SetRangeStyle.xlsx"
folder = "Ranges"
storage = ""
sheetName = "Sheet1"
rangeOperate = spirecloudexcel.models.RangeSetStyleRequest()
style = spirecloudexcel.models.Style()
rangeOperate.style = style
style.font = spirecloudexcel.models.Font()
style.font.color = spirecloudexcel.models.Color(255, 255, 0, 0)
style.font.underline = "single"
style.font.size = 12
style.font.name = "楷体"
style.font.is_italic = True
style.font.IsStrikeout = True
style.font.is_bold = True
borderList = list()
border1 = spirecloudexcel.models.Border("Medium", spirecloudexcel.models.Color(255, 255, 0, 0), "EdgeTop")
border2 = spirecloudexcel.models.Border("DashDot", spirecloudexcel.models.Color(255, 0, 255, 0), "EdgeRight")
border3 = spirecloudexcel.models.Border("Thin", spirecloudexcel.models.Color(0, 160, 32, 240), "EdgeLeft")
rangeOperate.style.border_collection = borderList
borderList.append(border1)
borderList.append(border2)
borderList.append(border3)
style.horizontal_alignment = "left"
style.indent_level = 1
style.text_direction = "RightToLeft"
style.background_color = spirecloudexcel.models.Color(0, 255, 255, 0)
style.is_text_wrapped = True
range1 = spirecloudexcel.models.Range()
range1.first_row = 1
range1.first_column = 2
range1.row_count = 10
range1.column_count = 1
rangeOperate.range = range1
api.set_range_style(name, sheetName, rangeOperate, storage=storage, folder=folder)
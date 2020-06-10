import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "SetColumnStyle_1.xlsx"
sheetName = "Sheet4"
column_index = 2
style = spirecloudexcel.models.Style()
font = spirecloudexcel.models.Font()
font.underline = "Single"
font.size = 8
font.is_italic = True
font.is_bold = True
font.name = "Algerian"
style.font = font
style.is_text_wrapped = True
style.pattern = "Solid"
style.patternColor = "Black"
borders = []
topBorder = spirecloudexcel.models.Border("Medium", spirecloudexcel.models.Color(255, 255, 0, 0), "EdgeTop")
rightBorder = spirecloudexcel.models.Border("DashDot", spirecloudexcel.models.Color(255, 0, 255, 0),
                                            "EdgeRight")
borders.append(topBorder)
borders.append(rightBorder)
style.border_collection = borders
style.horizontal_alignment = "Center"
style.background_color = spirecloudexcel.models.Color(255, 0, 255, 0)
storage = ""
folder = "/Cells/Column/"
api.set_column_style(name, sheet_name=sheetName, column_index=column_index,
                          style=style,
                          folder=folder,
                          storage=storage)

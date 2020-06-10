import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "SetRowStyle_1.xlsx"
sheetName = "Sheet2"
row_index = 3
style = spirecloudexcel.models.Style()
font = spirecloudexcel.models.Font()
font.underline = "Single"
font.size = 8
font.is_italic = True
font.is_bold = True
font.name = "Algerian"
style.font = font
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
folder = "/Cells/Row/"
api.set_row_style(name, sheet_name=sheetName, row_index=row_index, style=style, folder=folder, storage=storage)

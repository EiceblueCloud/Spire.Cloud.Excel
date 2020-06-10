import spirecloudexcel
from spirecloudexcel.configuration import Configuration as ExcelConfiguration
from spirecloudexcel.api.cells_api import CellsApi

appId = "your id"
appKey = "your key"
baseUrl = "https://api.e-iceblue.cn"
configuration = ExcelConfiguration(appId, appKey,baseUrl)
api = spirecloudexcel.api.cells_api.CellsApi(configuration)
name = "SetCellStyle_1.xlsx"
folder = "/Cells/Cell/Style/"
storage = ""
sheetName = "Sheet4"
cellName = "C8"
style = spirecloudexcel.models.Style()
font = spirecloudexcel.models.Font()
font.is_bold = True
font.is_italic = True
font.underline = "Single"
font.name = "Algerian"
font.size = 8
style.font = font
style.background_color = spirecloudexcel.models.Color(255, 0, 255, 0)
style.horizontal_alignment = "Center" #Left, Right
borders = []
border1 = spirecloudexcel.models.Border("Medium", spirecloudexcel.models.Color(255, 255, 0, 0), "EdgeTop") #EdgeBottom
border2 = spirecloudexcel.models.Border("DashDot", spirecloudexcel.models.Color(255, 0, 255, 0), "EdgeRight") #EdgeLeft
borders.append(border1)
borders.append(border2)
style.border_collection = borders
api.set_cell_style(name, sheetName, cellName, style, folder=folder, storage=storage)